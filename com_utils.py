import os
import shutil
import sys


def repair_com_cache():
    """Clear the win32com gen_py cache to recover from COM cache corruption."""
    try:
        import win32com
        candidates = []
        try:
            candidates.append(win32com.__gen_path__)
        except Exception:
            pass
        local_app = os.environ.get("LOCALAPPDATA", "")
        if local_app:
            candidates.append(os.path.join(local_app, "Temp", "gen_py"))
        candidates.append(
            os.path.join(os.path.dirname(win32com.__file__), "gen_py")
        )

        for path in candidates:
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
                os.makedirs(path, exist_ok=True)
                init = os.path.join(path, "__init__.py")
                with open(init, "w") as f:
                    f.write("")
                print(f"[INFO] COM cache cleared: {path}")
    except Exception as e:
        print(f"[WARN] repair_com_cache failed: {e}")

    stale = [
        k for k in sys.modules
        if k.startswith("win32com.gen_py.") and k != "win32com.gen_py"
    ]
    for k in stale:
        del sys.modules[k]
