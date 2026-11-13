import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'core/src'))
from scan2excel.ocr_batch import iter_ocr_results
from scan2excel.ocr_resources import GIB, ResourceMonitor, choose_parallelism
from scan2excel.service import Scan2ExcelService
from scan2excel.web_app import WebApi
from scan2excel.table_structure import TableStructureEngine


class ResourceTests(unittest.TestCase):
    def test_configurable_limit_up_to_twelve(self):
        self.assertEqual(choose_parallelism(32,24*GIB,0,max_workers=12)[0],12)
        self.assertEqual(choose_parallelism(32,24*GIB,0,max_workers=3)[0],3)
        self.assertEqual(choose_parallelism(8,24*GIB,0,max_workers=12)[0],4)
        self.assertEqual(choose_parallelism(32,3*GIB,0,max_workers=12)[0],3)
        self.assertEqual(choose_parallelism(32,4*GIB,0,high_accuracy=True,max_workers=12)[0],2)
        self.assertEqual(choose_parallelism(32,GIB,0,previous=12,max_workers=12)[0],1)

    def test_resource_limits_and_hysteresis(self):
        for cpus, memory, cpu, previous, expected in [
            (12, 8*GIB, 10, 1, 2), (2, 8*GIB, 10, 1, 1),
            (12, GIB, 10, 1, 1), (12, 8*GIB, 80, 2, 1),
            (12, None, 10, 1, 1), (12, 8*GIB, None, 1, 1),
            (12, 1.8*GIB, 10, 1, 1), (12, 1.8*GIB, 10, 2, 2),
            (12, 8*GIB, 65, 1, 1), (12, 8*GIB, 65, 2, 2),
        ]:
            with self.subTest(cpus=cpus, memory=memory, cpu=cpu, previous=previous):
                self.assertEqual(choose_parallelism(cpus,memory,cpu,previous)[0], expected)
        self.assertEqual(choose_parallelism(12, 3*GIB, 10, high_accuracy=True)[0], 1)

    def test_ocr_cpu_is_excluded_from_external_load(self):
        with patch('scan2excel.ocr_resources.os.cpu_count', return_value=12), \
             patch('scan2excel.ocr_resources._cpu_totals', side_effect=[(0,0),(1200,0)]), \
             patch('scan2excel.ocr_resources.time.monotonic', side_effect=[0,1]), \
             patch('scan2excel.ocr_resources.time.process_time', side_effect=[0,6]), \
             patch('scan2excel.ocr_resources.available_memory', return_value=8*GIB):
            monitor = ResourceMonitor()
            cpus, memory, external = monitor.sample()
        self.assertEqual(cpus,12)
        self.assertAlmostEqual(external,50)
        self.assertEqual(choose_parallelism(cpus,memory,external)[0],2)


class SchedulingTests(unittest.TestCase):
    def test_twelve_jobs_can_run_without_exceeding_limit(self):
        barrier=threading.Barrier(12)
        guard=threading.Lock()
        active,peak=[0],[0]
        def process(i):
            with guard:
                active[0]+=1
                peak[0]=max(peak[0],active[0])
            try:
                barrier.wait(5)
                return i
            finally:
                with guard:active[0]-=1
        results=dict(iter_ocr_results(range(12),process,12,threading.Event(),lambda _:None))
        self.assertEqual(peak[0],12)
        self.assertEqual({i:f.result() for i,f in results.items()},dict(enumerate(range(12))))

    def test_borderless_fallback_model_is_not_called_concurrently(self):
        model=TableStructureEngine.__new__(TableStructureEngine)
        model._infer_lock=threading.Lock()
        active,peak=[0],[0]
        def infer(*args,**kwargs):
            active[0]+=1
            peak[0]=max(peak[0],active[0])
            time.sleep(0.01)
            active[0]-=1
            return SimpleNamespace(pred_htmls=['<table><tr><td>A</td><td>B</td></tr></table>'],cell_bboxes=[])
        model._engine=infer
        items=[([[0,0],[10,0],[10,10],[0,10]],'A',1)]
        with ThreadPoolExecutor(max_workers=2) as pool:
            results=list(pool.map(lambda _:model.recognize(None,items),range(2)))
        self.assertEqual(peak[0],1)
        self.assertTrue(all(result['rows']==[['A','B']] for result in results))

    def test_out_of_order_results_keep_input_indices(self):
        release = threading.Event()
        started = []
        def process(i):
            if i == 0:
                self.assertTrue(release.wait(5))
            return f'result-{i}'
        stream = iter_ocr_results([0,1,2], process, 2, threading.Event(), started.append)
        try:
            index, future = next(stream)
            self.assertEqual(index,1)
            self.assertEqual(future.result(),'result-1')
            release.set()
            results = {index:future.result()}
            results.update((i,f.result()) for i,f in stream)
            self.assertEqual(results,{0:'result-0',1:'result-1',2:'result-2'})
        finally:
            release.set()
            stream.close()

    def test_cancellation_does_not_start_queued_pages(self):
        cancel, release = threading.Event(), threading.Event()
        started = []
        def process(i):
            if i == 1:
                release.wait(5)
            return i
        stream = iter_ocr_results(range(8),process,2,cancel,started.append)
        try:
            next(stream)
            cancel.set()
            release.set()
            list(stream)
            self.assertEqual(started,[0,1])
        finally:
            release.set()
            stream.close()

    def test_reducing_capacity_waits_for_running_page(self):
        release = threading.Event()
        started, capacity = [], [2]
        def process(i):
            if i == 1:
                release.wait(5)
            return i
        stream = iter_ocr_results(range(4),process,2,threading.Event(),started.append,lambda:capacity[0])
        try:
            self.assertEqual(next(stream)[0],0)
            capacity[0] = 1
            release.set()
            self.assertEqual(next(stream)[0],1)
            self.assertEqual(started,[0,1])
            self.assertEqual([i for i,_ in stream],[2,3])
        finally:
            release.set()
            stream.close()

    def test_page_failure_is_isolated(self):
        def process(i):
            if i == 0:
                raise ValueError('bad page')
            return i
        results = dict(iter_ocr_results([0,1],process,2,threading.Event(),lambda _:None))
        with self.assertRaises(ValueError):
            results[0].result()
        self.assertEqual(results[1].result(),1)


class ApiTests(unittest.TestCase):
    def test_worker_stores_pages_in_original_order(self):
        api = WebApi.__new__(WebApi)
        api.state = dict(ocr_workers=2,ocr_parallel_limit=1,high_accuracy=False,
                         auto_rotate=True,template_auto=True,busy=True,current=-1,
                         images=[dict(path=str(i),name=str(i),status='等待') for i in range(3)],log=[])
        api._templates, api._pages = [], []
        api._cancel, api._log_guard = threading.Event(), threading.Lock()
        release = threading.Event()
        stored = []
        real_store = api._store_page
        def store(index,page):
            real_store(index,page)
            stored.append(index)
            if index == 1:
                release.set()
        api._store_page = store
        class FakeService:
            def __init__(self,ocr_threads=None):
                pass
            def process_image(self,path,**kwargs):
                if path == '0' and not release.wait(5):
                    raise RuntimeError('parallel second page did not run')
                return SimpleNamespace(path=path,rows=[[path]],error=None,geometry_saved=False,
                    mode='table',n_rows=1,n_cols=1,elapsed=0,warped=False,borderless=False,
                    template='',warning='',fill_mode='vlookup',row_mismatch={},checks=[],min_score=1)
        class FakeMonitor:
            def __init__(self):
                self.calls = 0
            def sample(self):
                memory = [8*GIB,GIB,8*GIB][min(self.calls,2)]
                self.calls += 1
                return 12,memory,10
        with patch('scan2excel.service.Scan2ExcelService',FakeService), \
             patch('scan2excel.web_app.os.cpu_count',return_value=12), \
             patch('scan2excel.ocr_resources.ResourceMonitor',FakeMonitor), \
             patch('scan2excel.web_app.time.sleep'):
            api._worker([0,1,2])
        self.assertFalse(api.state['busy'])
        self.assertEqual(stored[0],1)
        self.assertEqual([page.path for page in api._pages],['0','1','2'])
        self.assertTrue(all(info['status']=='完成' for info in api.state['images']))
        decisions = [entry['text'] for entry in api.state['log'] if '本批同时' in entry['text']]
        self.assertEqual([next(n for n in [1,2] if f'最多 {n} 张' in line) for line in decisions], [2,1,2])

    def test_missing_geometry_forces_sequential_template_capture(self):
        api = WebApi.__new__(WebApi)
        api.state = dict(ocr_workers=2,ocr_parallel_limit=1,busy=True,current=-1,
                         images=[dict(path=str(i),name=str(i),status='等待') for i in range(2)],log=[])
        api._templates = [SimpleNamespace(path='template.xlsx',pages=[
            SimpleNamespace(col_fracs=[],row_fracs=[])])]
        api._pages = []
        api._cancel,api._log_guard = threading.Event(),threading.Lock()
        service_threads = []
        class FakeService:
            def __init__(self,ocr_threads=None):service_threads.append(ocr_threads)
            def process_image(self,path,**kwargs):
                return SimpleNamespace(path=path,error=None,geometry_saved=False,
                    mode='table',n_rows=1,n_cols=1,elapsed=0,warped=False,borderless=False,
                    template='',warning='',fill_mode='vlookup',row_mismatch={},checks=[],min_score=1)
        with patch('scan2excel.service.Scan2ExcelService',FakeService), \
             patch('scan2excel.web_app.os.cpu_count',return_value=12):
            api._worker([0,1])
        self.assertEqual(service_threads,[None])
        self.assertEqual(api.state['ocr_parallel_limit'],1)
        self.assertTrue(all(info['status']=='完成' for info in api.state['images']))

    def test_window_close_cancels_batch(self):
        api=WebApi.__new__(WebApi)
        api.state=dict(busy=True)
        api._cancel=threading.Event()
        api._log=lambda _:None
        callbacks=[]
        class ClosedEvent:
            def __iadd__(self,handler):
                callbacks.append(handler)
                return self
        api.attach_window(SimpleNamespace(events=SimpleNamespace(closed=ClosedEvent())))
        callbacks[0]()
        self.assertTrue(api._cancel.is_set())

    def test_settings_guard(self):
        api = WebApi.__new__(WebApi)
        api.state = dict(busy=False)
        api._save_last_paths = lambda:None
        api._log = lambda _:None
        self.assertEqual(api.set_ocr_workers(0)['ocr_workers'],0)
        self.assertEqual(api.set_ocr_workers(12)['ocr_workers'],12)
        for invalid in [True,13,-1,'2',2.0]:
            with self.assertRaises(ValueError):api.set_ocr_workers(invalid)
        api.state['busy'] = True
        with self.assertRaises(RuntimeError):api.set_ocr_workers(2)

    def test_parallel_services_have_independent_engines(self):
        with patch('scan2excel.service.OcrEngine',side_effect=lambda **_:object()):
            a,b = Scan2ExcelService(ocr_threads=3),Scan2ExcelService(ocr_threads=3)
            self.assertIs(a._engine(),a._engine())
            self.assertIsNot(a._engine(),b._engine())


if __name__=='__main__':
    unittest.main()
