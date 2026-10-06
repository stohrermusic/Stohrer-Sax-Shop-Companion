"""No function may both call `_()` and assign a local named `_`.

`_` is gettext, installed into builtins by i18n.init_translation(). Python
decides a name is local for a WHOLE function body if it is assigned
anywhere in it, so one `for _, x in ...` or `a, _, _ = f()` in a function
that also calls `_("...")` breaks every one of those calls: UnboundLocalError
before the assignment, "'tuple' object is not callable" after it. The
message the user was meant to see becomes an "Unexpected Error" dialog —
or, when the `_()` call sits at the top of the function, the function
cannot run at all.

Found 2026-10-06 by the Tooling end-to-end test: every error path of the
die G-code handler crashed this way, and the same shape sat in the G-code
settings Save and the toner's Analyze dialog. This scan is scope-aware
(nested defs, lambdas and comprehensions have their own `_`) and runs on
every module at the repo root. CLAUDE.md's i18n notes already warn about
the pattern; this makes it a gate.
"""
import ast
import glob
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCOPES = (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef,
          ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)


def own_scope_nodes(fn):
    """Every node in fn's own scope, not descending into nested scopes."""
    out = []
    stack = list(ast.iter_child_nodes(fn))
    while stack:
        node = stack.pop()
        if isinstance(node, SCOPES):
            continue
        out.append(node)
        stack.extend(ast.iter_child_nodes(node))
    return out


def offenders(path):
    tree = ast.parse(open(path, encoding="utf-8").read(), filename=path)
    found = []
    for fn in ast.walk(tree):
        if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        nodes = own_scope_nodes(fn)
        assigns = sorted({n.lineno for n in nodes
                          if isinstance(n, ast.Name) and n.id == "_" and isinstance(n.ctx, ast.Store)})
        calls = sorted({n.lineno for n in nodes
                        if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == "_"})
        if assigns and calls:
            found.append((fn.lineno, fn.name, assigns, calls))
    return found


def main():
    print("gettext `_` shadowing scan")
    print("=" * 60)
    files = sorted(glob.glob(os.path.join(ROOT, "*.py")))
    bad = 0
    for path in files:
        hits = offenders(path)
        name = os.path.basename(path)
        if hits:
            bad += len(hits)
            for lineno, fn, assigns, calls in hits:
                print(f"  FAIL  {name}:{lineno} {fn}() assigns `_` at {assigns} and calls _() at {calls[:4]}")
        else:
            print(f"  PASS  {name}: no function shadows gettext")
    print("=" * 60)
    if bad:
        print(f"{bad} function(s) shadow `_` — rename the throwaway (e.g. `_unused`)")
        return 1
    print(f"{len(files)} modules clean")
    return 0


if __name__ == "__main__":
    sys.exit(main())
