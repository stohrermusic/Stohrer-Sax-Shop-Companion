"""Run the test suites on a private Windows desktop, so their windows never show on,
or take focus from, the desktop the user is working on.

Matt, 2026-11-11: "any chance the tests could be done minimized? I can't use my pc for
a while when they are running." A Tk window cannot be kept from showing or from taking
focus on the desktop it is created on, so the suites are moved to a desktop of their
own: `tools/run_tests.py` re-runs itself through `rerun_hidden` on a desktop made with
CreateDesktopW, and every process it starts (the suites, their --selftest children,
the tours) inherits it. The display is the same (the suites read the same screen size),
the keyboard focus inside that desktop is theirs alone, and nothing of it is drawn on
the user's.

Only on Windows, only outside CI (GITHUB_ACTIONS set: the runners have no user to
protect, and CI stays exactly as it was), and not with --visible or SSC_VISIBLE=1.
Elsewhere `wanted` is False and nothing here runs.

Python's subprocess.STARTUPINFO has no working lpDesktop (an attribute set on it is
ignored by _winapi, the child lands on the user's desktop all the same; measured
2026-11-11), so the child is started with CreateProcessW through ctypes, inside a job
object that kills the whole tree when this process goes. Standard handles are
inherited, so the output and the exit code pass through unchanged.

As a command: `python tools/hidden_desktop.py <command...>` runs any command there
(a probe, a tour, a planted-break run).
"""
import os
import subprocess
import sys

DESKTOP_PREFIX = "ssc-tests-"
ENV_FLAG = "SSC_HIDDEN_DESKTOP"       # set in the child: the name of the desktop it runs on


def wanted(argv):
    """Should this run move to a private desktop? Windows, not CI, not asked to be visible,
    and not already there."""
    if sys.platform != "win32":
        return False
    if os.environ.get("GITHUB_ACTIONS") or os.environ.get("SSC_VISIBLE") == "1":
        return False
    if os.environ.get(ENV_FLAG):
        return False
    return "--visible" not in argv


def where(argv=None):
    """One line saying where the suites' windows are."""
    argv = argv or []
    if sys.platform != "win32":
        return "suites on this display (not Windows: no private desktop)"
    if os.environ.get("GITHUB_ACTIONS"):
        return "suites on the runner's display (CI)"
    if os.environ.get(ENV_FLAG):
        return f"suites on the private desktop {os.environ[ENV_FLAG]!r}: nothing shows on yours (--visible to watch)"
    if "--visible" in argv or os.environ.get("SSC_VISIBLE") == "1":
        return "suites on your desktop (--visible)"
    return "suites on your desktop"


def current_desktop_name():
    """The name of the desktop this thread runs on (Windows), else ''."""
    if sys.platform != "win32":
        return ""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    user32.GetThreadDesktop.restype = wintypes.HANDLE
    user32.GetThreadDesktop.argtypes = [wintypes.DWORD]
    h = user32.GetThreadDesktop(kernel32.GetCurrentThreadId())
    buf = ctypes.create_unicode_buffer(256)
    needed = wintypes.DWORD()
    UOI_NAME = 2
    if not user32.GetUserObjectInformationW(h, UOI_NAME, buf, ctypes.sizeof(buf), ctypes.byref(needed)):
        return ""
    return buf.value


def rerun_hidden(argv, cwd=None):
    """Run `argv` (a command line; argv[0] a program) on a new private desktop and
    return its exit code. The child gets ENV_FLAG so it knows, and stays visible-free
    for everything it starts. Ctrl-C here ends the whole tree (the job object)."""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    class STARTUPINFOW(ctypes.Structure):
        _fields_ = [("cb", wintypes.DWORD), ("lpReserved", wintypes.LPWSTR), ("lpDesktop", wintypes.LPWSTR),
                    ("lpTitle", wintypes.LPWSTR), ("dwX", wintypes.DWORD), ("dwY", wintypes.DWORD),
                    ("dwXSize", wintypes.DWORD), ("dwYSize", wintypes.DWORD), ("dwXCountChars", wintypes.DWORD),
                    ("dwYCountChars", wintypes.DWORD), ("dwFillAttribute", wintypes.DWORD),
                    ("dwFlags", wintypes.DWORD), ("wShowWindow", wintypes.WORD), ("cbReserved2", wintypes.WORD),
                    ("lpReserved2", ctypes.c_void_p), ("hStdInput", wintypes.HANDLE),
                    ("hStdOutput", wintypes.HANDLE), ("hStdError", wintypes.HANDLE)]

    class PROCESS_INFORMATION(ctypes.Structure):
        _fields_ = [("hProcess", wintypes.HANDLE), ("hThread", wintypes.HANDLE),
                    ("dwProcessId", wintypes.DWORD), ("dwThreadId", wintypes.DWORD)]

    class JOBOBJECT_BASIC_LIMIT_INFORMATION(ctypes.Structure):
        _fields_ = [("PerProcessUserTimeLimit", ctypes.c_int64), ("PerJobUserTimeLimit", ctypes.c_int64),
                    ("LimitFlags", wintypes.DWORD), ("MinimumWorkingSetSize", ctypes.c_size_t),
                    ("MaximumWorkingSetSize", ctypes.c_size_t), ("ActiveProcessLimit", wintypes.DWORD),
                    ("Affinity", ctypes.c_size_t), ("PriorityClass", wintypes.DWORD),
                    ("SchedulingClass", wintypes.DWORD)]

    class IO_COUNTERS(ctypes.Structure):
        _fields_ = [(n, ctypes.c_uint64) for n in ("ReadOperationCount", "WriteOperationCount",
                                                    "OtherOperationCount", "ReadTransferCount",
                                                    "WriteTransferCount", "OtherTransferCount")]

    class JOBOBJECT_EXTENDED_LIMIT_INFORMATION(ctypes.Structure):
        _fields_ = [("BasicLimitInformation", JOBOBJECT_BASIC_LIMIT_INFORMATION), ("IoInfo", IO_COUNTERS),
                    ("ProcessMemoryLimit", ctypes.c_size_t), ("JobMemoryLimit", ctypes.c_size_t),
                    ("PeakProcessMemoryUsed", ctypes.c_size_t), ("PeakJobMemoryUsed", ctypes.c_size_t)]

    GENERIC_ALL = 0x10000000
    CREATE_SUSPENDED = 0x00000004
    CREATE_UNICODE_ENVIRONMENT = 0x00000400
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
    JobObjectExtendedLimitInformation = 9
    WAIT_TIMEOUT = 0x102
    STILL_ACTIVE = 259

    user32.CreateDesktopW.restype = wintypes.HANDLE
    user32.CreateDesktopW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p, wintypes.DWORD,
                                      wintypes.DWORD, ctypes.c_void_p]
    kernel32.CreateProcessW.restype = wintypes.BOOL
    kernel32.CreateProcessW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.c_void_p, ctypes.c_void_p,
                                        wintypes.BOOL, wintypes.DWORD, ctypes.c_void_p, wintypes.LPCWSTR,
                                        ctypes.POINTER(STARTUPINFOW), ctypes.POINTER(PROCESS_INFORMATION)]
    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
    kernel32.SetInformationJobObject.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
    kernel32.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
    kernel32.ResumeThread.argtypes = [wintypes.HANDLE]
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    kernel32.TerminateJobObject.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    name = f"{DESKTOP_PREFIX}{os.getpid()}"
    desktop = user32.CreateDesktopW(name, None, None, 0, GENERIC_ALL, None)
    if not desktop:
        raise ctypes.WinError(ctypes.get_last_error())
    job = kernel32.CreateJobObjectW(None, None)
    if not job:
        user32.CloseDesktop(desktop)
        raise ctypes.WinError(ctypes.get_last_error())
    info = JOBOBJECT_EXTENDED_LIMIT_INFORMATION()
    info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    kernel32.SetInformationJobObject(job, JobObjectExtendedLimitInformation, ctypes.byref(info), ctypes.sizeof(info))

    os.environ[ENV_FLAG] = name                    # inherited: the child knows, and never re-runs itself
    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(si)
    si.lpDesktop = name
    pi = PROCESS_INFORMATION()
    cmdline = ctypes.create_unicode_buffer(subprocess.list2cmdline(argv))
    try:
        ok = kernel32.CreateProcessW(None, cmdline, None, None, True, CREATE_SUSPENDED | CREATE_UNICODE_ENVIRONMENT,
                                     None, cwd, ctypes.byref(si), ctypes.byref(pi))
        if not ok:
            raise ctypes.WinError(ctypes.get_last_error())
        kernel32.AssignProcessToJobObject(job, pi.hProcess)
        kernel32.ResumeThread(pi.hThread)
        kernel32.CloseHandle(pi.hThread)
        try:
            while kernel32.WaitForSingleObject(pi.hProcess, 500) == WAIT_TIMEOUT:
                pass                                # a slice at a time, so Ctrl-C reaches us
        except KeyboardInterrupt:
            kernel32.TerminateJobObject(job, 130)
            raise
        code = wintypes.DWORD(STILL_ACTIVE)
        kernel32.GetExitCodeProcess(pi.hProcess, ctypes.byref(code))
        kernel32.CloseHandle(pi.hProcess)
        return int(code.value)
    finally:
        del os.environ[ENV_FLAG]
        kernel32.CloseHandle(job)                   # ends anything still running in it
        user32.CloseDesktop(desktop)


def main(argv):
    if len(argv) < 2:
        print("usage: python tools/hidden_desktop.py <command...>")
        return 2
    if not wanted(["--"] + argv[1:]) and os.environ.get(ENV_FLAG):
        return subprocess.call(argv[1:])
    if sys.platform != "win32":
        return subprocess.call(argv[1:])
    return rerun_hidden(argv[1:])


if __name__ == "__main__":
    sys.exit(main(sys.argv))
