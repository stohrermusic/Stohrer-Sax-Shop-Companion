"""Worker threads beside the Tk thread, started so that they cannot deadlock it.

The deadlock (CI, 2026-11-10, found by the runner's faulthandler dump of a hung
test_calibration_dialog): the Tk thread called threading.Thread.start(), which
waits for the new thread to report that it has started. The new thread's
bootstrap allocates, the cyclic garbage collector ran on it, and it collected a
dead tkinter Variable left in a reference cycle by a closed dialog.
Variable.__del__ calls into Tcl, and a Tcl call from any thread but the Tk
thread is queued for the Tk thread and waited on; the Tk thread was still
inside start(), waiting for this very thread. Neither could move. The same can
happen wherever the Tk thread blocks on a worker: a lock the worker holds, a
join. It can freeze the app for a user at any thread start whenever a Variable
sits in an uncollected cycle, which closed dialogs leave behind; more dialogs
and more threads made it likelier, which is why it showed up now.

The rule: the collector is off while the Tk thread is blocked on a worker.
  * start_worker(): the one way the app starts a thread. Thread.start()
    returns only after the new thread has set its started event, which is
    after the bootstrap's allocations, so with the collector held off across
    start() no collection can run on the new thread while the Tk thread waits.
  * GcHeldRLock: a re-entrant lock whose holder has the collector off, for the
    locks a worker holds while the Tk thread may wait on them.
  * gc_held(): the context manager both use, for a bounded join.
gc.disable() is process-wide; holders are counted, so nested and concurrent
holds compose, and a collector that was already off stays off. The holds are
short (a thread start, a serial write, a bounded join), so collection is only
delayed, never lost. tools/test_threads.py gates all of it: an AST scan fails
on any threading.Thread( in a root module outside this file.
"""
import gc
import threading

_depth = 0
_was_enabled = True
_depth_lock = threading.Lock()


class gc_held:
    """While any holder is inside, the cyclic collector is off."""

    def __enter__(self):
        global _depth, _was_enabled
        with _depth_lock:
            if _depth == 0:
                _was_enabled = gc.isenabled()
                gc.disable()
            _depth += 1
        return self

    def __exit__(self, exc_type, exc, tb):
        global _depth
        with _depth_lock:
            _depth -= 1
            if _depth == 0 and _was_enabled:
                gc.enable()
        return False


def holds():
    """How many holders currently keep the collector off (for tests)."""
    with _depth_lock:
        return _depth


def start_worker(target, name=None, args=(), kwargs=None, daemon=True):
    """Create and start a thread with the collector held off across start().
    Returns the thread."""
    t = threading.Thread(target=target, name=name, args=args, kwargs=kwargs, daemon=daemon)
    with gc_held():
        t.start()
    return t


class GcHeldRLock:
    """A re-entrant lock whose holder has the cyclic collector off, so a
    thread waiting for it can never be waited on in turn by a finaliser
    the holder runs. Same calls as threading.RLock: acquire(blocking,
    timeout) returning whether it got the lock, release(), `with`."""

    def __init__(self):
        self._lock = threading.RLock()

    def acquire(self, blocking=True, timeout=-1):
        got = self._lock.acquire(blocking, timeout)
        if got:
            gc_held().__enter__()
        return got

    def release(self):
        gc_held().__exit__(None, None, None)
        self._lock.release()

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.release()
        return False
