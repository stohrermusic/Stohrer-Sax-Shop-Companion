"""The laser may be plugged in, unplugged or re-plugged while the app runs.

Matt, 2026-11-09: "is there a way for the ssc to be ok with a user plugging
in the laser when the ssc is already running? I feel like now it has to be
plugged in first." The app used to probe every serial port every 10 s until
a laser answered and then never look again; a laser unplugged, power-cycled
or moved to another socket left a stale port and a stale homed flag.

What is pinned:
  - the probe (falcon_sender.probe_grbl_port) asks "?" first and accepts a
    Grbl status report without ever writing the soft-reset byte; a silent
    port gets the Ctrl-X and the banner wait as before; garbage is rejected
  - the watch (falcon_sender.LaserWatch), with a simulated port list and a
    fake clock: a port that appears is probed at once; a laser-looking
    silent port is asked again every 10 s; an unrelated port once per
    appearance; a Bluetooth port never on its own; nothing while a laser is
    known; a known port that disappears is lost and found again under any
    name; a rescan probes everything, Bluetooth last; known IDs first
  - through the real app (main._laser_watch_pass / _laser_watch_apply with
    fakes): found shows Frame & Cut on both tabs, lost hides them and clears
    the homed flag so Try Auto Locate is not offered, found again brings
    them back still not homed
  - no log line per tick over a simulated hour with nothing attached
  - building the app, --selftest and the tour start no watch, no thread,
    no timer; a worker result after the window is gone raises nothing

No real serial port is ever opened: every probe here is a fake.
Isolated profile (SAXSHOP_CONFIG_DIR). GUI checks self-skip without a display.
"""
import inspect
import json
import re
import logging
import os
import sys
import tempfile
import threading
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-laser-watch-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
with open(os.path.join(_PROFILE, "app_settings.json"), "w") as _f:
    json.dump({"experimental_machine_menu": True}, _f)

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import falcon_sender as fs  # noqa: E402

results = []


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        box_guard.settle()      # a question this check did not expect fails it
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}{box_guard.reset()}")
        traceback.print_exc()
        results.append(False)


# ---------------------------------------------------------------- fixtures
FALCON = {'device': 'COM7', 'description': 'USB Serial Device (COM7)', 'hwid': 'USB VID:PID=303A:4001',
          'vid': 0x303A, 'pid': 0x4001, 'manufacturer': 'Espressif'}
CH340 = {'device': 'COM3', 'description': 'USB-SERIAL CH340 (COM3)', 'hwid': 'USB VID:PID=1A86:7523',
         'vid': 0x1A86, 'pid': 0x7523, 'manufacturer': 'wch.cn'}
ARDUINO = {'device': 'COM5', 'description': 'Arduino Uno (COM5)', 'hwid': 'USB VID:PID=2341:0043',
           'vid': 0x2341, 'pid': 0x0043, 'manufacturer': 'Arduino'}
BLUETOOTH = {'device': 'COM1', 'description': 'Standard Serial over Bluetooth link (COM1)',
             'hwid': 'BTHENUM\\{00001101-0000-1000-8000-00805F9B34FB}', 'vid': None, 'pid': None,
             'manufacturer': 'Microsoft'}
FALCON_B = dict(FALCON, device='COM12', description='USB Serial Device (COM12)')
CREALITY = {'device': 'COM3', 'description': 'Creality Falcon (COM3)', 'hwid': 'USB VID:PID=1234:0001',
            'vid': 0x1234, 'pid': 0x0001, 'manufacturer': ''}       # a laser by description, unknown VID
# Laser-looking by DESCRIPTION alone: no VID:PID at all (a driver that reports none).
DESC_ONLY = [{'device': 'COM9', 'description': 'USB-SERIAL CH340 (COM9)', 'hwid': 'n/a', 'vid': None, 'pid': None,
              'manufacturer': None},
             {'device': 'COM10', 'description': 'Falcon (COM10)', 'hwid': 'n/a', 'vid': None, 'pid': None,
              'manufacturer': None},
             {'device': 'COM11', 'description': 'Creality (COM11)', 'hwid': 'n/a', 'vid': None, 'pid': None,
              'manufacturer': None}]


class Clock:
    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t


class Sim:
    """A port list and a fake probe, driven by the test."""

    def __init__(self):
        self.ports = []
        self.answers = {}           # device -> 'status' / 'banner' / None, or a callable(now)
        self.probes = []            # (device, now)
        self.clock = Clock()

    def probe(self, device):
        self.probes.append((device, self.clock.t))
        answer = self.answers.get(device)
        return answer(self.clock.t) if callable(answer) else answer

    def run(self, watch, seconds, step=2.0):
        """Looks every `step` seconds for `seconds`, probing as told; returns the events."""
        events = []
        end = self.clock.t + seconds
        while self.clock.t < end:
            for p in watch.look(self.ports, self.clock.t):
                watch.probed(p['device'], self.probe(p['device']), self.clock.t)
                if watch.known_port:
                    break           # as main._laser_watch_pass does
            watch.finish()
            events += watch.pop_events()
            self.clock.t += step
        return events


# ---------------------------------------------------------------- the probe
class FakeSerial:
    """A port with a script: what it answers to '?' and whether a reset
    brings the banner. Records every byte written."""
    opened = []

    def __init__(self, port, baud, timeout=0.2, *, status=None, banner_after_reset=False, garbage=b''):
        self.port, self.written, self.out = port, b'', b'' + garbage
        self.status, self.banner_after_reset = status, banner_after_reset
        FakeSerial.opened.append(port)

    def reset_input_buffer(self):
        pass

    def write(self, data):
        self.written += data
        if data == fs.RT_STATUS and self.status:
            self.out += self.status
        if data == fs.RT_SOFT_RESET and self.banner_after_reset:
            self.out += b"\r\nGrbl 1.1f ['$' for help]\r\n"
        return len(data)

    def read(self, n):
        chunk, self.out = self.out[:n], self.out[n:]
        return chunk

    def close(self):
        pass


def factory(**script):
    def make(port, baud, timeout=0.2):
        return FakeSerial(port, baud, timeout, **script)
    return make


def probe(script):
    return fs.probe_grbl_port("COMX", ask_timeout_s=0.05, timeout_s=0.05, settle_s=0.0, serial_factory=factory(**script))


def the_probe_asks_before_it_resets():
    # Alive: a status report answers "?", and the reset byte is never written.
    alive = {}

    def remembering(port, baud, timeout=0.2):
        alive['s'] = FakeSerial(port, baud, timeout, status=b'<Idle|MPos:0.000,0.000,0.000|FS:0,0>\r\n')
        return alive['s']
    how = fs.probe_grbl_port("COMX", ask_timeout_s=0.05, timeout_s=0.05, settle_s=0.0, serial_factory=remembering)
    assert how == 'status', how
    assert fs.RT_SOFT_RESET not in alive['s'].written, f"a laser that answered was reset: {alive['s'].written!r}"
    assert alive['s'].written == fs.RT_STATUS
    # Alive and busy: a Run state is an answer too (an SD-card job must not be reset).
    assert probe(dict(status=b'<Run|MPos:12.000,3.000,0.000|FS:1200,300>\r\n')) == 'status'
    assert probe(dict(status=b'<Alarm|MPos:0.000,0.000,0.000>\r\n')) == 'status'
    # Silent until reset: the old path, unchanged.
    quiet = {}

    def quiet_port(port, baud, timeout=0.2):
        quiet['s'] = FakeSerial(port, baud, timeout, banner_after_reset=True)
        return quiet['s']
    assert fs.probe_grbl_port("COMX", ask_timeout_s=0.05, timeout_s=0.05, settle_s=0.0, serial_factory=quiet_port) == 'banner'
    assert quiet['s'].written == fs.RT_STATUS + fs.RT_SOFT_RESET, quiet['s'].written
    # Garbage, or nothing at all: rejected.
    assert probe(dict(garbage=b'\xff\x00hello<>\r\n')) is None
    assert probe({}) is None
    # A port that cannot be opened.

    def refuses(port, baud, timeout=0.2):
        raise OSError("busy")
    assert fs.probe_grbl_port("COMX", serial_factory=refuses, settle_s=0.0) is None
    assert fs._probe_for_grbl.__doc__ and "probe_grbl_port" in fs._probe_for_grbl.__doc__


check("probe: a status report to '?' is accepted with no reset byte written; a silent port gets the reset and "
      "banner as before; garbage is rejected", the_probe_asks_before_it_resets)


# ---------------------------------------------------------------- the watch
def nothing_then_a_laser_appears():
    sim, w = Sim(), fs.LaserWatch()
    assert sim.run(w, 20) == [] and sim.probes == [], "nothing attached: nothing probed, nothing said"
    sim.ports = [FALCON]
    sim.answers['COM7'] = 'status'
    events = sim.run(w, 2)
    assert [e[:3] for e in events] == [('found', 'COM7', 'status')], events
    assert sim.probes == [('COM7', sim.clock.t - 2)], "probed on the look it appeared on, once"
    assert w.known_port == 'COM7'


check("watch: nothing attached, then a laser appears: probed on that look and found", nothing_then_a_laser_appears)


def a_silent_laser_port_is_asked_again_every_ten_seconds():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [CH340]
    t0 = sim.clock.t
    sim.answers['COM3'] = lambda now: 'banner' if now >= t0 + 25 else None     # switched on 25 s later
    events = sim.run(w, 40)
    times = [t - t0 for d, t in sim.probes]
    assert times == [0.0, 10.0, 20.0, 30.0], f"re-probed every 10 s until it answered: {times}"
    assert [e[:3] for e in events if e[0] != 'silent'] == [('found', 'COM3', 'banner')], events
    assert [e for e in events if e[0] == 'silent'] == [('silent', 'COM3', CH340)], "one 'silent' event, not one per probe"


check("watch: a laser-looking port that is silent, then answers 25 s later, is found by the 10 s re-probe",
      a_silent_laser_port_is_asked_again_every_ten_seconds)


def a_port_that_looks_like_a_laser_by_description_alone_is_retried_too():
    # No VID:PID at all: the description is the only hint (ch340 / falcon / creality).
    for p in DESC_ONLY:
        assert fs.port_priority(p) == 1 and fs.port_looks_like_a_laser(p), p['description']
        sim, w = Sim(), fs.LaserWatch()
        sim.ports = [p]
        t0 = sim.clock.t
        sim.answers[p['device']] = lambda now, t0=t0: 'banner' if now >= t0 + 25 else None
        events = sim.run(w, 40)
        times = [t - t0 for d, t in sim.probes]
        assert times == [0.0, 10.0, 20.0, 30.0], f"{p['description']}: re-probed every 10 s until it answered: {times}"
        assert [e[:3] for e in events if e[0] != 'silent'] == [('found', p['device'], 'banner')], events
    # And an unknown VID:PID with a plain description is not a laser-looking port: once, then left alone.
    plain = {'device': 'COM13', 'description': 'USB Serial Port (COM13)', 'hwid': 'USB VID:PID=0403:6001',
             'vid': 0x0403, 'pid': 0x6001, 'manufacturer': 'FTDI'}
    assert fs.port_priority(plain) == 2
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [plain]
    sim.run(w, 60)
    assert len(sim.probes) == 1, sim.probes


check("watch: a port recognised as a laser by its description alone (no VID:PID) is asked again every 10 s while "
      "silent; an unknown ID with a plain description once", a_port_that_looks_like_a_laser_by_description_alone_is_retried_too)


def a_rescan_that_finds_the_same_laser_keeps_it():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [FALCON, ARDUINO]
    sim.answers['COM7'] = 'status'
    sim.run(w, 2)
    assert w.known_port == 'COM7'
    n = len(sim.probes)
    # A connect failed for a moment; the laser is still there and answers on the same port.
    w.request_rescan()
    events = sim.run(w, 2)
    assert not [e for e in events if e[0] == 'lost'], f"the laser was lost by a rescan that found it: {events}"
    assert ('rescan', 1, 'COM7') in events, events
    assert w.known_port == 'COM7', w.known_port
    assert [d for d, _t in sim.probes[n:]] == ['COM7'], f"the rescan went on after the laser answered: {sim.probes[n:]}"
    # And the watch is quiet again afterwards: the laser is known.
    sim.run(w, 60)
    assert len(sim.probes) == n + 1, f"probed after the rescan found the laser: {sim.probes[n + 1:]}"


check("watch: a rescan that finds the laser on the port it already knew keeps it: no 'lost', the rescan names it, "
      "nothing probed afterwards", a_rescan_that_finds_the_same_laser_keeps_it)


def an_unrelated_port_is_probed_once_then_left_alone():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [ARDUINO]
    events = sim.run(w, 600)
    assert len(sim.probes) == 1, f"an unrelated port was probed {len(sim.probes)} times in ten minutes"
    assert [e[:2] for e in events] == [('silent', 'COM5')]
    # Gone and back: once more.
    sim.ports = []
    sim.run(w, 4)
    sim.ports = [ARDUINO]
    sim.run(w, 60)
    assert len(sim.probes) == 2, "a port that went away and came back is probed once more"
    # The user asks: probed again.
    w.request_rescan()
    events = sim.run(w, 2)
    assert len(sim.probes) == 3 and ('rescan', 1, None) in events, (len(sim.probes), events)


check("watch: an unrelated port that never answers is probed once in ten minutes; again after it reappears or a rescan",
      an_unrelated_port_is_probed_once_then_left_alone)


def bluetooth_is_never_probed_on_the_watchs_own_account():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [BLUETOOTH, ARDUINO]
    sim.run(w, 600)
    assert [d for d, _t in sim.probes] == ['COM5'], f"probed: {sim.probes}"
    # macOS names its Bluetooth port in the device path and reports 'n/a' for the rest.
    mac_bt = {'device': '/dev/cu.Bluetooth-Incoming-Port', 'description': 'n/a', 'hwid': 'n/a',
              'vid': None, 'pid': None, 'manufacturer': None}
    assert fs.port_is_bluetooth(mac_bt) and fs.port_is_bluetooth(BLUETOOTH)
    assert not fs.port_is_bluetooth(ARDUINO) and not fs.port_is_bluetooth(FALCON)
    sim2, w2 = Sim(), fs.LaserWatch()
    sim2.ports = [mac_bt]
    sim2.run(w2, 60)
    assert sim2.probes == [], f"the Mac's Bluetooth port was opened: {sim2.probes}"
    w.request_rescan()
    sim.run(w, 2)
    # COM1 sorts before COM5 by name: only the rule puts it last.
    assert [d for d, _t in sim.probes][1:] == ['COM5', 'COM1'], f"a rescan probes everything, Bluetooth last: {sim.probes}"


check("watch: a Bluetooth port is never opened by the watch; a manual rescan opens it, last",
      bluetooth_is_never_probed_on_the_watchs_own_account)


def found_lost_and_found_again_under_another_name():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [FALCON, ARDUINO]
    sim.answers['COM7'] = 'status'
    events = sim.run(w, 2)
    assert [e[:2] for e in events] == [('found', 'COM7')], events
    assert [d for d, _t in sim.probes] == ['COM7'], "the known ID is probed first, and the laser answered, so nothing else is"
    n = len(sim.probes)
    sim.ports = [ARDUINO]                          # unplugged
    events = sim.run(w, 2)
    # Lost on that look; with no laser known the Arduino, never probed
    # while the laser was there, is asked once (and found silent).
    assert [e[:2] for e in events] == [('lost', 'COM7'), ('silent', 'COM5')] and w.known_port is None, events
    assert len(sim.probes) == n + 1 and sim.probes[-1][0] == 'COM5', sim.probes[n:]
    sim.run(w, 20)
    assert len(sim.probes) == n + 1, "the unrelated port is left alone after its one probe"
    sim.ports = [ARDUINO, FALCON_B]                # back on another socket
    sim.answers['COM12'] = 'banner'
    events = sim.run(w, 2)
    assert [e[:3] for e in events] == [('found', 'COM12', 'banner')], events
    assert w.known_port == 'COM12'


check("watch: found, then the port goes (lost), then back as another port name: found again",
      found_lost_and_found_again_under_another_name)


def two_ports_at_once_known_id_first():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [ARDUINO, CREALITY, FALCON]
    sim.answers['COM7'] = 'status'
    order = [p['device'] for p in w.look(sim.ports, sim.clock.t)]
    assert order == ['COM7', 'COM3', 'COM5'], f"priority order: known VID:PID, then description, then the rest: {order}"
    assert [fs.port_priority(p) for p in (FALCON, CH340, CREALITY, ARDUINO, BLUETOOTH)] == [0, 0, 1, 2, 2]
    # The real pass probes in that order and stops once the laser answers.
    events = sim.run(w, 2)
    assert w.known_port == 'COM7' and [d for d, _t in sim.probes] == ['COM7'], sim.probes
    assert [e[:2] for e in events] == [('found', 'COM7')], events


check("watch: two ports appear at once, the known ID is probed first", two_ports_at_once_known_id_first)


def nothing_is_probed_while_a_laser_is_known():
    sim, w = Sim(), fs.LaserWatch()
    sim.ports = [FALCON]
    sim.answers['COM7'] = 'status'
    sim.run(w, 2)
    n = len(sim.probes)
    sim.ports = [FALCON, ARDUINO, CH340]           # new ports appear during a cut
    events = sim.run(w, 600)
    assert len(sim.probes) == n and events == [], f"ports were opened while a laser was known: {sim.probes[n:]} {events}"


check("watch: while a laser is known no port is opened, whatever appears", nothing_is_probed_while_a_laser_is_known)


def a_quiet_hour_says_nothing():
    sim, w = Sim(), fs.LaserWatch()
    events = sim.run(w, 3600)
    assert events == [] and sim.probes == []
    sim.ports = [ARDUINO, BLUETOOTH]
    events = sim.run(w, 3600)
    assert len(events) == 1 and events[0][0] == 'silent', f"an hour with one unrelated port: one line, got {len(events)}"


check("watch: a simulated hour with nothing attached produces no event (nothing to log per tick)", a_quiet_hour_says_nothing)


# ---------------------------------------------------------------- the app
def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import main as main_mod
    import ui_dialogs
    real_supported = main_mod.machine_supported
    main_mod.machine_supported = lambda: True
    threads_before = {t.name for t in threading.enumerate()}
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()

    def shown():
        pad = bool(app._frame_cut_btn.winfo_manager())
        tooling = [bool(b.winfo_manager()) for b, _o in app._tooling_frame_cut_btns]
        return pad, tooling

    machine_log = []

    class Grab(logging.Handler):
        def emit(self, record):
            machine_log.append(record.getMessage())
    logging.getLogger("ssc.machine").addHandler(Grab())
    logging.getLogger("ssc.machine").setLevel(logging.INFO)

    def building_the_app_starts_nothing():
        assert getattr(app, "_laser_watch", None) is None, "the watch exists before the launch path started it"
        new_threads = {t.name for t in threading.enumerate()} - threads_before
        assert not any("laser" in n or "falcon" in n for n in new_threads), new_threads
        assert app.falcon_port is None
        # The user asking with no watch running is a pure no-op: nothing
        # probed, nothing scheduled, the port left exactly as it was.
        afters = []
        root.after = lambda ms, fn=None, *args: afters.append((ms, fn, args))
        app.falcon_port = "COM-untouched"
        try:
            app._detect_falcon_async()
        finally:
            del root.after
        assert app.falcon_port == "COM-untouched", f"a rescan with no watch touched the port: {app.falcon_port}"
        app.falcon_port = None
        assert afters == [], f"a rescan with no watch scheduled something: {afters}"
        assert not getattr(app, "_laser_watch_busy", False), "a rescan with no watch marked a worker busy"
        assert not any("laser" in t.name for t in threading.enumerate()), "a rescan ran with no watch"
        # The one call site is the launch block, and the probe suites' gates
        # (selftest, the tour) never reach it.
        src = inspect.getsource(main_mod)
        launch = re.split(r"\nif __name__ == ['\"]__main__['\"]:", src, maxsplit=1)[1]
        assert "app._start_laser_watch" in launch, "the launch block does not start the watch"
        assert "_start_laser_watch" not in inspect.getsource(main_mod.PadSVGGeneratorApp.__init__)
        assert "_start_laser_watch" not in inspect.getsource(main_mod._selftest)
        assert "_start_laser_watch" not in inspect.getsource(main_mod.run_tour)
        assert "_detect_falcon_async" not in launch, "the launch block still starts the old one-shot detection"
    check("app: building the app starts no watch, no thread and no timer; selftest and the tour never do", building_the_app_starts_nothing)

    def the_launch_path_starts_nothing_without_the_toggle_or_pyserial():
        # The __main__ block checks too, which masked it; this is the method's own rule.
        afters = []
        root.after = lambda ms, fn=None, *args: afters.append((ms, fn, args))
        real_has = fs.HAS_PYSERIAL
        try:
            app.settings["experimental_machine_menu"] = False
            fs.HAS_PYSERIAL = True
            app._start_laser_watch()
            assert getattr(app, "_laser_watch", None) is None, "a watch started with machine integration off"
            app.settings["experimental_machine_menu"] = True
            fs.HAS_PYSERIAL = False
            app._start_laser_watch()
            assert getattr(app, "_laser_watch", None) is None, "a watch started without pyserial"
            assert afters == [] and not getattr(app, "_laser_watch_busy", False), afters
        finally:
            del root.after
            fs.HAS_PYSERIAL = real_has
            app.settings["experimental_machine_menu"] = True
    check("app: _start_laser_watch starts nothing with machine integration off, and nothing without pyserial",
          the_launch_path_starts_nothing_without_the_toggle_or_pyserial)

    def the_live_loop_ticks_frees_its_worker_and_ticks_again():
        # The real _start_laser_watch / _laser_watch_tick / _laser_watch_worker
        # chain, with root.after recording instead of scheduling and the worker
        # thread run inline, on a fake port list and probe.
        afters, lists, probes = [], [], []
        root.after = lambda ms, fn=None, *args: afters.append((ms, fn, args))
        real_thread, real_has = threading.Thread, fs.HAS_PYSERIAL
        keep = fs.list_serial_ports, fs.probe_grbl_port

        class Inline:
            def __init__(self, target=None, name=None, daemon=None, **k):
                self.target, self.name = target, name

            def start(self):
                assert self.name == 'laser-watch', self.name
                self.target()
        fs.list_serial_ports = lambda: lists.append(1) or [dict(FALCON)]
        fs.probe_grbl_port = lambda device, *a, **k: probes.append(device) or 'status'
        threading.Thread, fs.HAS_PYSERIAL = Inline, True
        del machine_log[:]
        try:
            app._start_laser_watch()
            assert isinstance(app._laser_watch, fs.LaserWatch) and any("laser watch started" in m for m in machine_log)
            # One pass ran on the worker, and it left the worker slot free.
            assert lists == [1] and probes == ['COM7'], (lists, probes)
            assert app._laser_watch_busy is False, "the worker left itself marked busy"
            applies = [a for a in afters if a[1] == app._laser_watch_apply]
            ticks = [a for a in afters if a[1] == app._laser_watch_tick]
            assert len(applies) == 1 and applies[0][0] == 0, f"the pass's events were not handed to the Tk thread: {afters}"
            assert len(ticks) == 1 and ticks[0][0] == main_mod.LASER_WATCH_PERIOD_MS, \
                f"the next tick is not scheduled at LASER_WATCH_PERIOD_MS: {afters}"
            # Delivered: the laser is found.
            app._laser_watch_apply(*applies[0][2])
            assert app.falcon_port == 'COM7'
            # The next tick starts a second pass (the laser is known, so it probes nothing, but it looks).
            ticks[0][1]()
            assert lists == [1, 1], f"the second tick did not run a pass: {lists}"
            assert probes == ['COM7'], f"probed while the laser was known: {probes}"
            assert app._laser_watch_busy is False
            assert sum(1 for a in afters if a[1] == app._laser_watch_tick) == 2, "the second tick did not reschedule"
            # The watch is not started twice.
            first = app._laser_watch
            app._start_laser_watch()
            assert app._laser_watch is first and lists == [1, 1]
        finally:
            del root.after
            threading.Thread, fs.HAS_PYSERIAL = real_thread, real_has
            fs.list_serial_ports, fs.probe_grbl_port = keep
            app._laser_watch = None
            app._laser_watch_busy = False
            app.falcon_port = None
            app._refresh_machine_ui_state()
            del machine_log[:]
    check("app: the live loop: a tick runs one pass on the worker, hands its events to the Tk thread, frees the "
          "worker and schedules the next tick, which runs again", the_live_loop_ticks_frees_its_worker_and_ticks_again)

    # Wire a watch by hand, the way the launch path does, but drive it with fakes and a fake clock.
    sim = Sim()
    app._laser_watch = fs.LaserWatch()
    app._laser_watch_busy = False

    def one_pass():
        events = app._laser_watch_pass(lambda: list(sim.ports), sim.probe, sim.clock)
        app._laser_watch_apply(events)
        root.update()
        sim.clock.t += 2.0
        return events

    def found_shows_the_buttons_on_both_tabs():
        assert shown() == (False, [False] * len(app._tooling_frame_cut_btns)), "fixture: no laser, no buttons"
        sim.ports = [FALCON]
        sim.answers['COM7'] = 'status'
        del machine_log[:]
        one_pass()
        assert app.falcon_port == 'COM7'
        assert shown() == (True, [True] * len(app._tooling_frame_cut_btns)), f"buttons after found: {shown()}"
        assert len(machine_log) == 1 and "laser found on COM7" in machine_log[0] and "303A:4001" in machine_log[0] \
            and "status report" in machine_log[0], machine_log
    check("app: a laser found shows Frame & Cut on both tabs and logs the port, its ID and which probe answered",
          found_shows_the_buttons_on_both_tabs)

    def lost_hides_them_and_forgets_the_home():
        app._falcon_homed_this_session = True
        sim.ports = []
        del machine_log[:]
        one_pass()
        assert app.falcon_port is None and app._falcon_homed_this_session is False
        assert shown() == (False, [False] * len(app._tooling_frame_cut_btns)), f"buttons after lost: {shown()}"
        assert machine_log == ["laser gone from COM7"], machine_log
        # The next run offers no Try Auto Locate: the jog dialog is told the laser is not homed.
        kws = []

        class FakeRun:
            def __init__(self, parent, sender, lines, **kw):
                kws.append(kw)
                self._final_reason = "cancelled"

        class FakeSender:
            def __init__(self, port=None):
                pass

            def connect(self):
                pass

            def unlock(self):
                pass

            def disconnect(self):
                pass
        keep = ui_dialogs.FalconRunDialog, fs.FalconSender
        ui_dialogs.FalconRunDialog, fs.FalconSender = FakeRun, FakeSender
        app.falcon_port = "COM7"            # as if the laser were back for this one click
        app.settings["units"] = "mm"
        for entry, value in ((app.width_entry, "200"), (app.height_entry, "150")):
            entry.delete(0, tk.END)
            entry.insert(0, value)
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", "20 x 2")
        for m, v in app.material_vars.items():
            v.set(m == "felt")
        app.preview_var.set(False)
        app.settings["show_engraving_warning"] = False
        keep_nest = main_mod.NestingPreviewWindow

        class Nest:
            def __init__(self, *a, **k):
                self.result = "save"
        main_mod.NestingPreviewWindow = Nest
        try:
            app.on_frame_and_cut()
        finally:
            ui_dialogs.FalconRunDialog, fs.FalconSender = keep
            main_mod.NestingPreviewWindow = keep_nest
            app.falcon_port = None
        assert kws and kws[0].get("is_homed") is False and kws[0].get("auto_locate_target") is None, kws[:1]
    check("app: the laser gone hides the buttons, clears the homed flag (no Try Auto Locate on the next run), logs once",
          lost_hides_them_and_forgets_the_home)

    def found_again_under_another_name():
        sim.ports = [ARDUINO, FALCON_B]
        sim.answers['COM12'] = 'banner'
        del machine_log[:]
        n = len(sim.probes)
        one_pass()
        assert app.falcon_port == 'COM12' and app._falcon_homed_this_session is False
        assert shown() == (True, [True] * len(app._tooling_frame_cut_btns))
        assert any("laser found on COM12" in m and "boot banner" in m for m in machine_log), machine_log
        # The known ID answered first, so the pass opened nothing else.
        assert [d for d, _t in sim.probes[n:]] == ['COM12'], f"probed after the laser answered: {sim.probes[n:]}"
        assert len(machine_log) == 1, machine_log
    check("app: back on another port name: found again, buttons back, still not homed", found_again_under_another_name)

    def a_connect_that_fails_asks_for_a_rescan():
        # Frame & Cut's connect failure: the same box as before, then a rescan is requested.
        class DeadSender:
            def __init__(self, port=None):
                pass

            def connect(self):
                raise OSError("port busy")

            def disconnect(self):
                pass
        shown_boxes = []
        keep_err = main_mod.messagebox.showerror
        main_mod.messagebox.showerror = lambda title, msg, **k: shown_boxes.append((str(title), str(msg)))
        keep = fs.FalconSender
        fs.FalconSender = DeadSender
        keep_nest = main_mod.NestingPreviewWindow

        class Nest:
            def __init__(self, *a, **k):
                self.result = "save"
        main_mod.NestingPreviewWindow = Nest
        app._laser_watch_busy = True        # so no worker thread starts in the test
        try:
            app.on_frame_and_cut()
            assert [b[0] for b in shown_boxes] == ["Connection Failed"], shown_boxes
            assert app._laser_watch.rescan_wanted is True, "a failed connect did not ask for a rescan"
            # The Machine menu's three connecting items do the same: their box, then a rescan.
            for handler, question in ((app._on_machine_home_falcon, "Home Laser?"),
                                      (app._on_machine_test_connection, None),
                                      (app._on_machine_clear_errors, None)):
                del shown_boxes[:]
                app._laser_watch.rescan_wanted = False
                if question:
                    box_guard.GUARD.expect(question, answer=True)
                handler()
                assert [b[0] for b in shown_boxes] == ["Connection Failed"], f"{handler.__name__}: {shown_boxes}"
                assert "port busy" in shown_boxes[0][1], shown_boxes
                assert app._laser_watch.rescan_wanted is True, f"{handler.__name__}: a failed connect did not ask for a rescan"
        finally:
            fs.FalconSender = keep
            main_mod.NestingPreviewWindow = keep_nest
            main_mod.messagebox.showerror = keep_err
            app._laser_watch_busy = False
        # The rescan probes every port present, the stale one included (it
        # is still listed: another program may hold it, or the laser behind
        # it is off), Bluetooth last, and reports; a stale port that does
        # not answer is lost, and the watch is free to find the laser again.
        sim.ports = [FALCON_B, ARDUINO, BLUETOOTH]
        sim.answers['COM12'] = None
        del machine_log[:]
        events = one_pass()
        assert ('lost', 'COM12') in events and ('rescan', 3, None) in events, events
        assert [d for d, _t in sim.probes][-3:] == ['COM12', 'COM5', 'COM1'], sim.probes[-3:]
        assert app.falcon_port is None and any("rescan asked for: 3 port(s) probed, no laser" in m for m in machine_log), machine_log
        assert app._laser_watch.known_port is None, "the rescan kept the stale port as the laser"
        sim.ports = [FALCON]
        sim.answers['COM7'] = 'status'
        events = one_pass()
        assert app.falcon_port == 'COM7' and [e[:2] for e in events] == [('found', 'COM7')], events
    check("app: a connect that fails (Frame & Cut, Home Laser, Test Connection, Clear Errors) shows its box and asks "
          "for a rescan, which probes every port, Bluetooth last", a_connect_that_fails_asks_for_a_rescan)

    def a_rescan_that_finds_the_same_laser_keeps_everything():
        # A connect failed for a moment (a busy port, say); the laser is still
        # there on the port the watch knew. The rescan must not lose it.
        assert app.falcon_port == 'COM7' and shown() == (True, [True] * len(app._tooling_frame_cut_btns)), "fixture"
        app._falcon_homed_this_session = True
        app._laser_watch.request_rescan()
        sim.ports = [FALCON, ARDUINO, BLUETOOTH]
        sim.answers['COM7'] = 'status'
        del machine_log[:]
        n = len(sim.probes)
        events = one_pass()
        assert not [e for e in events if e[0] == 'lost'], f"the laser was lost by a rescan that found it: {events}"
        assert ('rescan', 1, 'COM7') in events, events
        assert app.falcon_port == 'COM7' and app._laser_watch.known_port == 'COM7', (app.falcon_port, app._laser_watch.known_port)
        assert app._falcon_homed_this_session is True, "the homed flag was cleared by a rescan that found the same laser"
        assert shown() == (True, [True] * len(app._tooling_frame_cut_btns)), f"buttons after the rescan: {shown()}"
        assert [d for d, _t in sim.probes[n:]] == ['COM7'], f"the rescan went on past the laser: {sim.probes[n:]}"
        assert any("rescan asked for: 1 port(s) probed, laser on COM7" in m for m in machine_log), machine_log
        assert not any("gone" in m for m in machine_log), machine_log
        app._falcon_homed_this_session = False
    check("app: a rescan that finds the laser on the port it already knew: no 'lost', the homed flag and the buttons "
          "stay, the log names the port", a_rescan_that_finds_the_same_laser_keeps_everything)

    def a_quiet_hour_logs_nothing():
        sim.ports = []
        one_pass()                          # whatever the last check left is settled first
        del machine_log[:]
        for _ in range(1800):
            events = app._laser_watch_pass(lambda: list(sim.ports), sim.probe, sim.clock)
            app._laser_watch_apply(events)
            sim.clock.t += 2.0
        assert machine_log == [], f"{len(machine_log)} log lines over a quiet hour: {machine_log[:3]}"
        # A pass that raises every tick (port enumeration broken) is said once, not every 2 s.
        for _ in range(1800):
            app._laser_watch_apply([('error', "comports() exploded (test)")])
        assert machine_log == ["laser watch: comports() exploded (test)"], machine_log[:3]
        del machine_log[:]
    check("app: a simulated hour with nothing attached logs nothing; a pass that keeps raising is said once",
          a_quiet_hour_logs_nothing)

    def a_worker_result_after_the_window_is_gone_raises_nothing():
        root.destroy()
        app._laser_watch_apply([('found', 'COM7', 'status', FALCON)])
        # The real worker, with the port list and the probe stood in for: no
        # suite may open a real serial port (a probe can soft-reset whatever
        # is on it).
        probed, errors = [], []
        keep = fs.list_serial_ports, fs.probe_grbl_port
        fs.list_serial_ports = lambda: [dict(ARDUINO)]
        fs.probe_grbl_port = lambda device, *a, **k: probed.append(device)

        def run():
            try:
                app._laser_watch_worker()   # finds a port to probe, probes the stand-in, cannot post its result
            except BaseException as e:      # noqa: BLE001 - the point is that nothing gets here
                errors.append(e)
        # On its own thread, as in the app: with no main loop left, Tk refuses
        # the worker's after() ("main thread is not in main loop").
        t = threading.Thread(target=run, name="laser-watch-test")
        t.start()
        t.join(15)
        fs.list_serial_ports, fs.probe_grbl_port = keep
        assert not t.is_alive(), "the worker hung"
        assert not errors, f"the worker raised after the window was gone: {errors!r}"
        assert probed == ['COM5'], f"the worker did not run its pass: {probed}"
        assert app._laser_watch_busy is False, "the worker left itself marked busy"
        app._laser_watch_apply([('lost', 'COM7')])
    check("app: a worker result arriving after the root is destroyed raises nothing",
          a_worker_result_after_the_window_is_gone_raises_nothing)
    main_mod.machine_supported = real_supported


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
