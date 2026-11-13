"""Bounded page scheduling; completed results retain their original index."""
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from collections import deque


def iter_ocr_results(indices, process, workers, cancel, on_start, capacity=None):
    """Yield (index, Future), drain running jobs on cancellation, queue no extra pages."""
    if workers <= 1:
        for index in indices:
            if cancel.is_set():
                break
            on_start(index)
            future = Future()
            try:
                future.set_result(process(index))
            except Exception as exc:
                future.set_exception(exc)
            yield index, future
        return

    pending = deque(indices)
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="ocr-page") as pool:
        active = {}

        def submit_next():
            if cancel.is_set() or not pending:
                return
            index = pending.popleft()
            on_start(index)
            active[pool.submit(process, index)] = index

        def limit():
            return max(1, min(workers, capacity() if capacity else workers))

        for _ in range(limit()):
            submit_next()
        while active:
            done, _ = wait(active, return_when=FIRST_COMPLETED)
            for future in done:
                index = active.pop(future)
                yield index, future
            allowed = limit() if pending else 0
            while len(active) < allowed and not cancel.is_set():
                count = len(active)
                submit_next()
                if len(active) == count:
                    break
