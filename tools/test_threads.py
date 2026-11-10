"""Gates for threads.py (2026-11-10): the one way the app starts a worker thread, and the
locks the Tk thread may wait on, keep the cyclic collector off while the Tk thread is blocked
on a worker, so a worker can never finalise a tkinter Variable (a Tcl call from the wrong
thread, waited on by the Tk thread) while the Tk thread waits for it. The deadlock this
prevents froze test_calibration_dialog on CI (the runner's faulthandler dump named it).

  * an AST scan: no `threading.Thread(` in a root module outside threads.py, and no suite
    that owns a Tk root starts a raw thread;
  * the helper holds the collector off across Thread.start() and gives it back; holds nest
    and count; a collector that was off stays off; GcHeldRLock holds while held, is
    re-entrant, and a non-blocking acquire that fails changes nothing;
  * the deadlock itself, in child processes: Variables in cycles aged into the oldest
    generation and dropped, the collector armed to run after k more allocations, then a
    thread started the old way; for each k the parent bounds the child with a timeout. Where
    a k hangs the old way, the same k through start_worker must return. Where no k hangs on
    this interpreter (the trigger depends on the allocation order inside Thread.start, which
    differs by Python version), the check says so and passes on the scan and the unit checks.
No Tk root in this process; the children have their own.
"""
import ast
import gc
import os
import re
import subprocess
import sys
import tempfile
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
try:
    import threads
except ImportError:          # the tree before 2026-11-10: the scan below names every raw thread start
    threads = None

results = []


def check(name, fn):
    try:
        ok = fn()
        if ok is False:
            raise AssertionError("returned False")
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:  # noqa: BLE001
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        results.append(False)


# ---------------------------------------------------------------------------
# 1. The scans
# ---------------------------------------------------------------------------


def no_raw_thread_start_in_the_app():
    bad = []
    for f in sorted(x for x in os.listdir(ROOT) if x.endswith(".py")):
        if f == "threads.py":
            continue
        tree = ast.parse(open(os.path.join(ROOT, f), encoding="utf-8").read(), filename=f)
        for node in ast.walk(tree):
            if isinstance(node, ast.Call):
                fn = node.func
                if (isinstance(fn, ast.Attribute) and fn.attr == "Thread"
                        and isinstance(fn.value, ast.Name) and fn.value.id == "threading"):
                    bad.append(f"{f}:{node.lineno}")
                elif isinstance(fn, ast.Name) and fn.id == "Thread":
                    bad.append(f"{f}:{node.lineno} (Thread imported by name)")
    assert not bad, "threads started outside threads.start_worker: " + ", ".join(bad)
    users = [f for f in os.listdir(ROOT) if f.endswith(".py") and "start_worker(" in open(os.path.join(ROOT, f), encoding="utf-8").read()]
    assert {"main.py", "ui_dialogs.py", "falcon_sender.py", "toner_engine.py"} <= set(users), users


def no_suite_with_a_tk_root_starts_a_raw_thread():
    bad = []
    for f in sorted(os.listdir(os.path.join(ROOT, "tools"))):
        if not (f.startswith("test_") and f.endswith(".py")) or f == "test_threads.py":
            continue          # this file carries both strings inside CHILD, a program its children run
        src = open(os.path.join(ROOT, "tools", f), encoding="utf-8").read()
        if "threading.Thread(" in src and re.search(r"tk\.Tk\(|tkinter\.Tk\(", src):
            bad.append(f)
    assert not bad, "suites that own a Tk root and start a raw thread (use threads.start_worker): " + ", ".join(bad)


print("=== 1. Every thread start goes through threads.py ===")
check("no threading.Thread( in any root module outside threads.py; main, ui_dialogs, falcon_sender and "
      "toner_engine use start_worker", no_raw_thread_start_in_the_app)
check("no suite that owns a Tk root starts a raw thread", no_suite_with_a_tk_root_starts_a_raw_thread)

# ---------------------------------------------------------------------------
# 2. The helper
# ---------------------------------------------------------------------------


def start_worker_holds_the_collector_across_start():
    assert threads is not None, "no threads.py in this tree"
    assert gc.isenabled(), "this check needs the collector on"
    seen = []
    real_start = threading.Thread.start

    def recording_start(self):
        seen.append((gc.isenabled(), threads.holds()))
        real_start(self)
    threading.Thread.start = recording_start
    try:
        t = threads.start_worker(lambda: None, name="probe")
        t.join(5)
    finally:
        threading.Thread.start = real_start
    assert seen == [(False, 1)], f"inside start(): {seen} (collector on / holders)"
    assert gc.isenabled() and threads.holds() == 0, (gc.isenabled(), threads.holds())
    assert t.daemon and t.name == "probe" and not t.is_alive()


def holds_nest_and_count():
    assert threads is not None, "no threads.py in this tree"
    assert gc.isenabled()
    with threads.gc_held():
        assert not gc.isenabled() and threads.holds() == 1
        with threads.gc_held():
            assert not gc.isenabled() and threads.holds() == 2
        assert not gc.isenabled() and threads.holds() == 1, "the inner exit gave the collector back too early"
    assert gc.isenabled() and threads.holds() == 0
    # Two threads holding at once: the collector comes back when the last one leaves.
    go, inside = threading.Event(), threading.Event()

    def other():
        with threads.gc_held():
            inside.set()
            go.wait(5)
    t = threads.start_worker(other)
    inside.wait(5)
    with threads.gc_held():
        pass
    assert not gc.isenabled(), "the collector came back while another thread still held it"
    go.set()
    t.join(5)
    assert gc.isenabled() and threads.holds() == 0


def a_collector_that_was_off_stays_off():
    assert threads is not None, "no threads.py in this tree"
    gc.disable()
    try:
        with threads.gc_held():
            assert not gc.isenabled()
        assert not gc.isenabled(), "gc_held turned on a collector someone else had turned off"
    finally:
        gc.enable()


def the_lock_holds_while_held():
    assert threads is not None, "no threads.py in this tree"
    lock = threads.GcHeldRLock()
    assert gc.isenabled()
    assert lock.acquire() is True and not gc.isenabled() and threads.holds() == 1
    assert lock.acquire() is True and threads.holds() == 2              # re-entrant
    lock.release()
    assert not gc.isenabled()
    lock.release()
    assert gc.isenabled() and threads.holds() == 0
    with lock:
        assert not gc.isenabled()
        got = []
        t = threads.start_worker(lambda: got.append(lock.acquire(blocking=False)))
        t.join(5)
        assert got == [False], got                                       # held here: not given away
        assert not gc.isenabled() and threads.holds() == 1, "a failed acquire changed the hold"
    assert gc.isenabled() and threads.holds() == 0


print("=== 2. The helper ===")
check("start_worker holds the collector off across Thread.start() and gives it back", start_worker_holds_the_collector_across_start)
check("holds nest and count, across threads too", holds_nest_and_count)
check("a collector that was already off stays off", a_collector_that_was_off_stays_off)
check("GcHeldRLock: held means the collector is off; re-entrant; a failed non-blocking acquire changes nothing",
      the_lock_holds_while_held)

# ---------------------------------------------------------------------------
# 3. The deadlock itself, in children
# ---------------------------------------------------------------------------
CHILD = '''\
import gc, sys, threading
import tkinter as tk
k, mode, repo = int(sys.argv[1]), sys.argv[2], sys.argv[3]
sys.path.insert(0, repo)
root = tk.Tk()
root.withdraw()
keep = []
for _i in range(200):
    v = tk.StringVar(root, value="x")
    v.cycle = [v]              # a Variable in a reference cycle, as closed dialogs leave them
    keep.append(v)
gc.collect()
gc.collect()                   # aged into the oldest generation while still referenced
del v
del keep                       # garbage now, found only by a full collection
gc.set_threshold(k, 1, 1)      # a full collection after k more allocations, wherever they happen
gc.enable()
print("armed", flush=True)
if mode == "old":
    threading.Thread(target=lambda: None, daemon=True).start()
else:
    from threads import start_worker
    start_worker(lambda: None)
print("returned", flush=True)
root.destroy()
'''


def the_old_way_hangs_where_the_helper_returns():
    d = tempfile.mkdtemp(prefix="ssc-deadlock-")
    child = os.path.join(d, "child.py")
    open(child, "w", encoding="utf-8").write(CHILD)

    def run(k, mode):
        try:
            r = subprocess.run([sys.executable, child, str(k), mode, ROOT], capture_output=True, text=True, timeout=8)
            return "returned" in r.stdout
        except subprocess.TimeoutExpired:
            return False
    hung = []
    t0 = time.time()
    for k in range(1, 13):
        if not run(k, "old"):
            hung.append(k)
            if len(hung) == 2:
                break
    if not hung:
        print(f"    (no k in 1..12 hangs the old way on this interpreter, {sys.version.split()[0]}; "
              f"the trigger depends on the allocation order inside Thread.start; {time.time() - t0:.0f} s)")
        return
    for k in hung:
        assert run(k, "new"), f"k={k}: the old way hung and start_worker hung too"
    print(f"    (old way hung at k={hung}; start_worker returned at the same k; {time.time() - t0:.0f} s)")


print("=== 3. The deadlock in a child process ===")
check("Variables in aged cycles + a thread started the old way hangs (where this interpreter triggers it); "
      "the same through start_worker returns", the_old_way_hangs_where_the_helper_returns)

print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
