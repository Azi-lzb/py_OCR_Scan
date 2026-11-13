"""Local resource sampling and bounded OCR page admission."""
import os
import sys
import time
from pathlib import Path

GIB = 1024 ** 3


def available_memory():
    try:
        if sys.platform == 'win32':
            import ctypes

            class MemoryStatus(ctypes.Structure):
                _fields_ = [('length', ctypes.c_uint32), ('load', ctypes.c_uint32)] + [
                    (name, ctypes.c_uint64) for name in (
                        'total_phys', 'avail_phys', 'total_page', 'avail_page',
                        'total_virtual', 'avail_virtual', 'avail_extended')]

            status = MemoryStatus()
            status.length = ctypes.sizeof(status)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return status.avail_phys
        elif sys.platform.startswith('linux'):
            for line in Path('/proc/meminfo').read_text().splitlines():
                if line.startswith('MemAvailable:'):
                    return int(line.split()[1]) * 1024
    except (OSError, ValueError, AttributeError):
        pass
    return None


def _cpu_totals():
    try:
        if sys.platform == 'win32':
            import ctypes
            from ctypes.wintypes import FILETIME
            idle, kernel, user = FILETIME(), FILETIME(), FILETIME()
            if ctypes.windll.kernel32.GetSystemTimes(
                    ctypes.byref(idle), ctypes.byref(kernel), ctypes.byref(user)):
                value = lambda t: (t.dwHighDateTime << 32) | t.dwLowDateTime
                # Windows kernel time already includes idle time.
                return value(kernel) + value(user), value(idle)
        elif sys.platform.startswith('linux'):
            ticks = [int(v) for v in Path('/proc/stat').read_text().splitlines()[0].split()[1:9]]
            return sum(ticks), ticks[3] + ticks[4]
    except (OSError, ValueError, AttributeError, IndexError):
        pass
    return None


class ResourceMonitor:
    def __init__(self):
        self.cpus = os.cpu_count() or 1
        self.previous = (_cpu_totals(), time.monotonic(), time.process_time())
        self.external_cpu = None

    def sample(self):
        totals, now, own = _cpu_totals(), time.monotonic(), time.process_time()
        previous, then, old_own = self.previous
        if totals is None or previous is None:
            self.external_cpu = None
            self.previous = (totals, now, own)
        elif now - then >= 0.1:
            delta = totals[0] - previous[0]
            if delta > 0:
                busy = 100 * (1 - (totals[1] - previous[1]) / delta)
                own_busy = 100 * (own - old_own) / ((now - then) * self.cpus)
                # OCR's own CPU use is expected; react to competing work instead.
                self.external_cpu = max(0.0, min(100.0, busy - own_busy))
            self.previous = (totals, now, own)
        return self.cpus, available_memory(), self.external_cpu


def choose_parallelism(cpus, memory, external_cpu, previous=1, high_accuracy=False,
                       max_workers=2):
    """Hysteresis prevents repeated mode switching near memory/load thresholds."""
    if cpus < 4:
        return 1, 'CPU 核数较少'
    if memory is None or external_cpu is None:
        return 1, '资源信息暂不可用'
    previous = max(1, min(previous, max_workers))
    if external_cpu > (75 if previous > 1 else 55):
        return 1, f'其他任务 CPU 负载 {external_cpu:.0f}%'
    # 为每页保留 CPU 余量；内存按标准/高精度档分别预算，已运行模型给半档信用。
    cpu_limit = max(1, int(cpus * (1 - external_cpu / 100) / 2))
    upper = max(1, min(12, max_workers, cpu_limit))
    unit = 2 if high_accuracy else 1
    for count in range(upper, 1, -1):
        needed = unit * (count - min(previous - 1, count - 1) * 0.5) * GIB
        if memory >= needed:
            return count, f'可用内存 {memory / GIB:.1f} GB，其他任务 CPU 负载 {external_cpu:.0f}%'
    return 1, f'可用内存 {memory / GIB:.1f} GB，其他任务 CPU 负载 {external_cpu:.0f}%'
