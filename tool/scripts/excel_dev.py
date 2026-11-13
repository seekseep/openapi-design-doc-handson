"""Developer helper for macOS: drive Excel for Mac through AppleScript.

    python3 tool/scripts/excel_dev.py test
        Build a scratch workbook with the VBA in tool/vba and run the
        YAML parser tests against tool/fixtures.

    python3 tool/scripts/excel_dev.py import <workbook.xlsm> <openapi.yaml>
        Open the workbook, run the import without dialogs, save and close.

Windows users do not need this file (see README).
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))


def _as_str(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def osa(script: str) -> str:
    r = subprocess.run(["osascript", "-"], input=script, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip())
    return r.stdout.rstrip("\n")


def _click_enable_macros() -> bool:
    """Press "Enable Macros" if Excel is showing the macro warning."""
    script = """
tell application "System Events" to tell process "Microsoft Excel"
  repeat with w in windows
    repeat with b in buttons of w
      set n to name of b as string
      if n contains "有効" or n contains "Enable" then
        click b
        return "clicked"
      end if
    end repeat
  end repeat
end tell
return ""
"""
    return osa(script) == "clicked"


def open_book(path: Path) -> None:
    import time

    # "open" blocks while the macro warning is shown, so do not wait for it.
    osa(
        'tell application "Microsoft Excel"\n'
        "  ignoring application responses\n"
        f"    open (POSIX file {_as_str(str(path))})\n"
        "  end ignoring\n"
        "end tell"
    )
    for _ in range(60):
        time.sleep(0.5)
        _click_enable_macros()
        try:
            names = osa('tell application "Microsoft Excel" to get name of every workbook')
        except RuntimeError:
            continue
        if path.name in names.split(", "):
            return
    raise RuntimeError(f"could not open {path}")


def close_book(path: Path, save: bool) -> None:
    osa(
        f'tell application "Microsoft Excel" to close workbook {_as_str(path.name)} '
        f'saving {"yes" if save else "no"}'
    )


def save_book(path: Path) -> None:
    osa(f'tell application "Microsoft Excel" to save workbook {_as_str(path.name)}')


def dismiss_dialogs() -> str:
    """Read the text of any dialog Excel / VBE shows, press OK, return the text.
    Also reports the VBE window title and the selected text (the place of a
    compile error)."""
    script = """
set found to ""
tell application "System Events" to tell process "Microsoft Excel"
  repeat with w in windows
    try
      set texts to value of every static text of w
      repeat with t in texts
        set found to found & (t as string) & " "
      end repeat
      repeat with b in buttons of w
        if (name of b as string) is in {"OK", "終了", "End"} then
          click b
          exit repeat
        end if
      end repeat
    end try
  end repeat
end tell
delay 1
tell application "System Events" to tell process "Microsoft Excel"
  repeat with w in windows
    try
      set found to found & " | window: " & (name of w as string)
      set areas to entire contents of w
      repeat with e in areas
        try
          set sel to value of attribute "AXSelectedText" of e
          if sel is not missing value and sel is not "" then set found to found & " | selected: " & sel
        end try
      end repeat
    end try
  end repeat
end tell
return found
"""
    return osa(script).strip()


def run_macro(book: Path, macro: str, *args: str, timeout: int = 30) -> str:
    arg_text = "".join(f" arg{i} {_as_str(a)}" for i, a in enumerate(args, start=1))
    name = _as_str(f"'{book.name}'!{macro}")
    try:
        return osa(
            f"with timeout of {timeout} seconds\n"
            f'  tell application "Microsoft Excel" to run VB macro {name}{arg_text}\n'
            "end timeout"
        )
    except RuntimeError as e:
        dialog = dismiss_dialogs()
        raise RuntimeError(f"{macro} failed: {e}\n  dialog: {dialog}") from None


def build_scratch_book(out: Path) -> None:
    import openpyxl

    from inject_vba import inject

    base = out.with_suffix(".xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Test"
    wb.save(base)
    inject(base, ROOT / "tool" / "vba", out)
    base.unlink()


def cmd_test() -> int:
    fixtures = ROOT / "tool" / "fixtures"
    tmp = Path(tempfile.mkdtemp(prefix="vba-test-"))
    book = tmp / "vba-test.xlsm"
    build_scratch_book(book)
    open_book(book)
    failed = 0
    print("Excel にリポジトリのフォルダへのアクセス許可を求めます（初回のみダイアログが出ます）")
    granted = run_macro(book, "modTest.GrantAccess", str(ROOT), timeout=180)
    if granted != "True":
        print("アクセスが許可されませんでした")
        close_book(book, save=False)
        return 1
    try:
        cases = [(p, p.with_suffix(".expected")) for p in sorted(fixtures.glob("*.yaml"))]
        for yaml_path, expected in cases:
            r = run_macro(book, "modTest.TestYamlFile", str(yaml_path), str(expected))
            print(f"{'PASS' if r == 'OK' else 'FAIL'}  {yaml_path.name}")
            if r != "OK":
                failed += 1
                print("      " + r.replace("\n", "\n      "))
        for yaml_path in sorted((fixtures / "errors").glob("*.yaml")):
            r = run_macro(book, "modTest.TestYamlError", str(yaml_path))
            ok = r != "NO ERROR"
            failed += 0 if ok else 1
            print(f"{'PASS' if ok else 'FAIL'}  errors/{yaml_path.name}: {r}")
    finally:
        close_book(book, save=False)
    print("all passed" if failed == 0 else f"{failed} failed")
    return 1 if failed else 0


def cmd_import(book: Path, yaml_path: Path) -> int:
    book = book.resolve()
    open_book(book)
    try:
        r = run_macro(book, "modMain.ImportFromPath", str(yaml_path.resolve()))
        print(r)
        save_book(book)
    finally:
        close_book(book, save=False)
    return 0 if r.startswith("OK") else 1


def cmd_compile() -> int:
    """Find the module that fails to compile (Excel compiles on demand)."""
    import shutil

    from inject_vba import inject
    import openpyxl

    src = ROOT / "tool" / "vba"
    failed = 0
    for target in sorted([*src.glob("*.bas"), *src.glob("*.cls")]):
        tmp = Path(tempfile.mkdtemp(prefix="vba-compile-"))
        vba = tmp / "vba"
        shutil.copytree(src, vba)
        # A probe module that touches only the target module.
        if target.suffix == ".bas":
            text = (vba / target.name).read_text(encoding="utf-8")
            text += "\nPublic Function Probe__() As String\n    Probe__ = \"ok\"\nEnd Function\n"
            (vba / target.name).write_text(text, encoding="utf-8")
            macro = f"{target.stem}.Probe__"
        else:
            (vba / "modProbe__.bas").write_text(
                f'Attribute VB_Name = "modProbe__"\nPublic Function Probe__() As String\n'
                f"    Dim x As {target.stem}\n    Set x = New {target.stem}\n    Probe__ = \"ok\"\nEnd Function\n",
                encoding="utf-8",
            )
            macro = "modProbe__.Probe__"
        base = tmp / "b.xlsx"
        wb = openpyxl.Workbook()
        wb.save(base)
        book = tmp / f"compile-{target.stem}.xlsm"
        inject(base, vba, book)
        open_book(book)
        try:
            r = run_macro(book, macro, timeout=15)
        except RuntimeError as e:
            r = str(e)
            failed += 1
        finally:
            try:
                close_book(book, save=False)
            except RuntimeError:
                pass
        print(f"{target.name}: {r}")
    return 1 if failed else 0


def cmd_probe(extra: Path, calls: list[str]) -> int:
    """Build a scratch book with tool/vba plus <extra>.bas and run each
    call ("Module.Func" or "Module.Func|arg1|arg2")."""
    import shutil

    import openpyxl

    from inject_vba import inject

    tmp = Path(tempfile.mkdtemp(prefix="vba-probe-"))
    vba = tmp / "vba"
    shutil.copytree(ROOT / "tool" / "vba", vba)
    shutil.copy(extra, vba / extra.name)
    base = tmp / "b.xlsx"
    openpyxl.Workbook().save(base)
    book = tmp / "probe.xlsm"
    inject(base, vba, book)
    open_book(book)
    try:
        for call in calls:
            name, *args = call.split("|")
            print(f"{call} -> ", end="", flush=True)
            print(run_macro(book, name, *args, timeout=20))
    finally:
        try:
            close_book(book, save=False)
        except RuntimeError:
            pass
    return 0


def main(argv: list[str]) -> int:
    if len(argv) >= 3 and argv[0] == "probe":
        return cmd_probe(Path(argv[1]), argv[2:])
    if len(argv) >= 1 and argv[0] == "compile":
        return cmd_compile()
    if len(argv) >= 1 and argv[0] == "test":
        return cmd_test()
    if len(argv) == 3 and argv[0] == "import":
        return cmd_import(Path(argv[1]), Path(argv[2]))
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
