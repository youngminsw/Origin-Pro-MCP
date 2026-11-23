"""Out-of-band process checks on Windows (no COM).

Shared by the daemon (confirming an empty Origin exited after Exit) and by
``close_origin`` (confirming this session's Origin is gone).
"""
from __future__ import annotations

SYNCHRONIZE = 0x00100000
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WAIT_TIMEOUT = 0x102
ERROR_ACCESS_DENIED = 5


def process_running(pid: int) -> bool:
    """True while process ``pid`` exists and has not exited. Windows only.

    A process that exited but whose handle someone still holds counts as NOT
    running (its handle is signalled). When the process cannot be opened, the
    error decides: access denied means it exists but belongs to someone else
    (e.g. an elevated Origin), so it is reported as running — never as "gone",
    which would wrongly tell the caller an Exit took. Any other failure (no
    such process) means it is not running.
    """
    if not pid:
        return False
    import ctypes
    from ctypes import wintypes

    # A private WinDLL instance: setting argtypes on the shared ctypes.windll
    # object would change them for every other caller in the process.
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    k32.WaitForSingleObject.restype = wintypes.DWORD
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    k32.CloseHandle.restype = wintypes.BOOL

    handle = k32.OpenProcess(
        SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return ctypes.get_last_error() == ERROR_ACCESS_DENIED
    try:
        return k32.WaitForSingleObject(handle, 0) == WAIT_TIMEOUT
    finally:
        k32.CloseHandle(handle)
