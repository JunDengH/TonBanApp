"""Project-local Python startup fixes.

Python imports this module automatically when commands run from the project
root.  The Windows Python 3.14 virtual environment used here needs Tcl to be
given its executable path before tkinter creates the first interpreter.
"""
import ctypes
import os
import sys
from pathlib import Path


_TCL_DLL = None


def _is_pyinstaller_frontend():
    executable_name = Path(sys.argv[0]).name.lower()
    return executable_name in {
        "pyinstaller.exe",
        "pyinstaller-script.py",
        "pyi-makespec.exe",
        "pyi-makespec-script.py",
    }


def _bootstrap_windows_tcl_tk():
    global _TCL_DLL
    if os.name != "nt":
        return
    if _is_pyinstaller_frontend():
        return

    base = Path(sys.base_prefix)
    tcl_dll = base / "DLLs" / "tcl86t.dll"
    if tcl_dll.exists():
        try:
            _TCL_DLL = ctypes.WinDLL(str(tcl_dll))
            _TCL_DLL.Tcl_FindExecutable.argtypes = [ctypes.c_wchar_p]
            _TCL_DLL.Tcl_FindExecutable(str(Path(sys.executable)))
        except (AttributeError, OSError):
            _TCL_DLL = None

    tcl_dir = base / "tcl" / "tcl8.6"
    tk_dir = base / "tcl" / "tk8.6"
    if (tcl_dir / "init.tcl").exists():
        os.environ.setdefault("TCL_LIBRARY", str(tcl_dir))
    if (tk_dir / "tk.tcl").exists():
        os.environ.setdefault("TK_LIBRARY", str(tk_dir))


_bootstrap_windows_tcl_tk()
