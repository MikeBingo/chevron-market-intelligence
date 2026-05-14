"""
Run this from the dashboard root folder:
    python FIX_PYTHON_UTF8.py

It patches the two dashboard updater scripts so they read and write HTML/log text
using UTF-8. This fixes Windows 'charmap' failures caused by symbols such as:
    ✓  →  —

It is conservative: it creates .bak files first and only makes common safe replacements.
"""
from __future__ import annotations

from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
TARGETS = [
    ROOT / "Scripts" / "generate_dashboard_data.py",
    ROOT / "Scripts" / "generate_competitor_data.py",
]

REPLACEMENTS = [
    # pathlib read/write helpers
    (r"\.read_text\(\)", ".read_text(encoding='utf-8')"),
    (r"\.write_text\(([^\)]*)\)", lambda m: _patch_write_text(m.group(0))),
]


def _patch_write_text(call: str) -> str:
    if "encoding=" in call:
        return call
    return call[:-1] + ", encoding='utf-8')"


def patch_open_calls(text: str) -> str:
    """Add encoding='utf-8' to common text-mode open() calls that lack encoding."""
    pattern = re.compile(r"open\(([^\n\)]*)\)")

    def repl(match: re.Match[str]) -> str:
        call = match.group(0)
        inside = match.group(1)

        if "encoding=" in inside:
            return call
        if "'b'" in inside or '"b"' in inside or ", 'rb'" in inside or ', "rb"' in inside or ", 'wb'" in inside or ', "wb"' in inside:
            return call

        # Only patch calls that appear to specify a text read/write/append mode, or just a path.
        if any(mode in inside for mode in [", 'r'", ', "r"', ", 'w'", ', "w"', ", 'a'", ', "a"', ", mode='r'", ', mode="r"', ", mode='w'", ', mode="w"', ", mode='a'", ', mode="a"']) or "," not in inside:
            return f"open({inside}, encoding='utf-8')"
        return call

    return pattern.sub(repl, text)


def patch_file(path: Path) -> None:
    if not path.exists():
        print(f"MISSING: {path}")
        return

    original = path.read_text(encoding="utf-8", errors="replace")
    patched = original

    # Patch pathlib calls.
    patched = re.sub(r"\.read_text\(\)", ".read_text(encoding='utf-8')", patched)

    def write_text_repl(match: re.Match[str]) -> str:
        full = match.group(0)
        args = match.group(1)
        if "encoding=" in args:
            return full
        return f".write_text({args}, encoding='utf-8')"

    patched = re.sub(r"\.write_text\((.*?)\)", write_text_repl, patched)

    # Patch built-in open calls.
    patched = patch_open_calls(patched)

    # Add stdout/stderr safety near imports if absent.
    if "PYTHONIOENCODING" not in patched and "sys.stdout.reconfigure" not in patched:
        marker = "import "
        insert = "\nimport sys\ntry:\n    sys.stdout.reconfigure(encoding='utf-8')\n    sys.stderr.reconfigure(encoding='utf-8')\nexcept Exception:\n    pass\n"
        first_import = patched.find(marker)
        if first_import >= 0:
            line_end = patched.find("\n", first_import)
            patched = patched[: line_end + 1] + insert + patched[line_end + 1 :]
        else:
            patched = insert + "\n" + patched

    if patched == original:
        print(f"NO CHANGE: {path}")
        return

    backup = path.with_suffix(path.suffix + ".bak")
    backup.write_text(original, encoding="utf-8")
    path.write_text(patched, encoding="utf-8")
    print(f"PATCHED: {path}")
    print(f"BACKUP : {backup}")


if __name__ == "__main__":
    for target in TARGETS:
        patch_file(target)
