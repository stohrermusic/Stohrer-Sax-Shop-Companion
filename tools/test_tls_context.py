"""Gate: the web data imports verify the server (2026-11-10, pre-3.0 housekeeping).

config.get_ssl_context() is what the three "Import Matt's..." downloads hand to urlopen. Until
2026-11-10 it fell back to an UNVERIFIED context (check_hostname False, CERT_NONE) whenever the
system's default certificates could not be loaded, which silently defeated HTTPS for them. Now:
certifi's bundle (every build ships it, the Intel Mac build since the same day), then the system's
default certificates, then it raises, and the caller's existing "Could not fetch ..." box says why.

What is pinned here, with no network:
  * the context verifies (CERT_REQUIRED, check_hostname) and has certificate authorities loaded;
    config.assert_tls_ready(), which --selftest calls in the shipped bundle, says the same;
  * without certifi the result is still a verifying context or an OSError, never CERT_NONE;
  * with no certificate authority anywhere it raises instead of returning an unverified context;
  * no root module turns verification off (no CERT_NONE, no `check_hostname = False`);
  * a certificate failure, and a missing-CA failure, reach the user as the download error box
    ("Connection Error") in each of the three importers, and nothing else is shown;
  * _selftest names assert_tls_ready, so the Intel Mac build job checks its bundle.
Isolated profile (SAXSHOP_CONFIG_DIR). The GUI checks self-skip without a display.
"""
import inspect
import os
import ssl
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-tls-")

import box_guard  # noqa: E402
guard = box_guard.install()

import config  # noqa: E402

results = []


def check(name, fn):
    try:
        ok = fn()
        if ok is False:
            raise AssertionError("returned False")
        box_guard.settle()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:  # noqa: BLE001
        print(f"  FAIL  {name}: {e}")
        print(f"        boxes pending: {box_guard.reset()}")
        results.append(False)


def verifying():
    ctx = config.get_ssl_context()
    assert ctx.verify_mode == ssl.CERT_REQUIRED, ctx.verify_mode
    assert ctx.check_hostname is True
    n = len(ctx.get_ca_certs())
    assert n > 0, "no certificate authorities loaded"
    assert config.assert_tls_ready() == n
    print(f"    ({n} certificate authorities, certifi {'present' if _has_certifi() else 'absent'})")


def _has_certifi():
    try:
        import certifi  # noqa: F401
        return True
    except ImportError:
        return False


class _NoCertifi:
    """Make `import certifi` raise ImportError inside the block."""

    def __enter__(self):
        self.saved = sys.modules.get("certifi", "<absent>")
        sys.modules["certifi"] = None
        return self

    def __exit__(self, *exc):
        if self.saved == "<absent>":
            sys.modules.pop("certifi", None)
        else:
            sys.modules["certifi"] = self.saved


def without_certifi_verifying_or_refusing():
    with _NoCertifi():
        try:
            ctx = config.get_ssl_context()
        except OSError as e:
            print(f"    (no system certificates here: refused with {e})")
            return
        assert ctx.verify_mode == ssl.CERT_REQUIRED and ctx.check_hostname is True, \
            "an unverified context came back without certifi"
        assert ctx.get_ca_certs(), "a context with no certificate authorities came back"
        print(f"    (system default: {len(ctx.get_ca_certs())} certificate authorities)")


class _EmptyContext:
    verify_mode = ssl.CERT_REQUIRED
    check_hostname = True

    def get_ca_certs(self, binary_form=False):
        return []


def no_authorities_anywhere_raises():
    real = config.ssl.create_default_context
    config.ssl.create_default_context = lambda *a, **k: _EmptyContext()
    try:
        with _NoCertifi():
            try:
                ctx = config.get_ssl_context()
            except OSError as e:
                assert "certificate" in str(e).lower(), str(e)
                return
            raise AssertionError(f"returned a context instead of raising: verify_mode={ctx.verify_mode}, "
                                 f"check_hostname={ctx.check_hostname}")
    finally:
        config.ssl.create_default_context = real


def no_root_module_turns_verification_off():
    """In code, not in comments or docstrings (config.py's own docstring names
    CERT_NONE as the thing it no longer does): no `ssl.CERT_NONE`, no
    `check_hostname = False`, no `_create_unverified_context` in any root module."""
    import ast
    bad = []
    for f in sorted(os.listdir(ROOT)):
        if not f.endswith(".py"):
            continue
        tree = ast.parse(open(os.path.join(ROOT, f), encoding="utf-8").read(), filename=f)
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and node.attr in ("CERT_NONE", "_create_unverified_context"):
                bad.append(f"{f}:{node.lineno} {node.attr}")
            elif isinstance(node, ast.Name) and node.id in ("CERT_NONE", "_create_unverified_context"):
                bad.append(f"{f}:{node.lineno} {node.id}")
            elif isinstance(node, ast.Assign):
                for t in node.targets:
                    if (isinstance(t, ast.Attribute) and t.attr == "check_hostname"
                            and isinstance(node.value, ast.Constant) and node.value.value is False):
                        bad.append(f"{f}:{node.lineno} check_hostname = False")
    assert not bad, "verification turned off in: " + ", ".join(bad)


def selftest_checks_the_shipped_bundle():
    import main as main_mod
    src = inspect.getsource(main_mod._selftest)
    assert "assert_tls_ready" in src, "--selftest no longer checks the TLS context"


def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import urllib.error
    import urllib.request
    import main as main_mod
    import library_features
    app = main_mod.PadSVGGeneratorApp(root)
    importers = [("pad sets", app.on_import_matts_pad_sets),
                 ("key heights", app.on_import_matts_key_heights),
                 ("screw specs", app.on_import_matts_specs)]

    def failure_reaches_the_user(make_failure, expect_text):
        for label, handler in importers:
            del guard.shown[:]
            restore = make_failure()
            try:
                handler()
            finally:
                restore()
            boxes = list(guard.shown)
            assert len(boxes) == 1, f"{label}: expected one box, saw {boxes}"
            kind, title, message = boxes[0]
            assert kind == "showerror" and title == "Connection Error", f"{label}: {kind} {title!r}"
            assert expect_text in message, f"{label}: the box does not say why: {message!r}"
        del guard.shown[:]

    def certificate_failure_reaches_the_user():
        def make():
            real = urllib.request.urlopen

            def failing(*a, **k):
                raise urllib.error.URLError(ssl.SSLCertVerificationError(
                    1, "[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer"))
            urllib.request.urlopen = failing

            def restore():
                urllib.request.urlopen = real
            return restore
        failure_reaches_the_user(make, "CERTIFICATE_VERIFY_FAILED")
    check("a certificate failure reaches the user as the Connection Error box in all three importers, and nothing else",
          certificate_failure_reaches_the_user)

    def no_authorities_reaches_the_user():
        def make():
            saved = (main_mod.get_ssl_context, library_features.get_ssl_context)

            def refusing():
                raise OSError("no certificate authorities are available to verify the server's certificate")
            main_mod.get_ssl_context = refusing
            library_features.get_ssl_context = refusing
            real_open = urllib.request.urlopen

            def never(*a, **k):
                raise AssertionError("urlopen was reached without a context")
            urllib.request.urlopen = never

            def restore():
                main_mod.get_ssl_context, library_features.get_ssl_context = saved
                urllib.request.urlopen = real_open
            return restore
        failure_reaches_the_user(make, "no certificate authorities")
    check("a missing-certificate-authority refusal reaches the user the same way, before any network call",
          no_authorities_reaches_the_user)

    def the_importers_hand_urlopen_a_verifying_context():
        seen = []
        real = urllib.request.urlopen

        def capture(req, *a, **k):
            seen.append(k.get("context"))
            raise urllib.error.URLError("offline")
        urllib.request.urlopen = capture
        try:
            for _label, handler in importers:
                del guard.shown[:]
                handler()
        finally:
            urllib.request.urlopen = real
            del guard.shown[:]
        assert len(seen) == 3 and all(c is not None and c.verify_mode == ssl.CERT_REQUIRED and c.check_hostname
                                      for c in seen), seen
    check("each importer hands urlopen a verifying context (no network: urlopen stubbed)",
          the_importers_hand_urlopen_a_verifying_context)
    root.destroy()


print("=== TLS: the web data imports verify the server, and a failure is shown, never hidden ===")
check("get_ssl_context() verifies and has certificate authorities; assert_tls_ready() agrees", verifying)
check("without certifi: a verifying context or an OSError, never an unverified one", without_certifi_verifying_or_refusing)
check("with no certificate authority anywhere it raises instead of returning an unverified context",
      no_authorities_anywhere_raises)
check("no root module turns certificate verification off", no_root_module_turns_verification_off)
check("--selftest checks the shipped bundle through assert_tls_ready", selftest_checks_the_shipped_bundle)
gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
