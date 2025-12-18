import os
import shutil
import sys
import tempfile

import pythoncom  # noqa: F401
import win32com.client
import win32com.client.gencache

# Ensure COM cache is writable in frozen app; place gen_py in temp.
if hasattr(sys, "frozen"):
    gen_dir = os.path.join(tempfile.gettempdir(), "gen_py")
    os.environ["PYTHONCOMGENERATEDDIR"] = gen_dir
    try:
        os.makedirs(gen_dir, exist_ok=True)
    except Exception:
        pass
    win32com.client.gencache.is_readonly = False
    try:
        win32com.client.gencache.Rebuild()
    except Exception:
        try:
            shutil.rmtree(gen_dir, ignore_errors=True)
            os.makedirs(gen_dir, exist_ok=True)
        except Exception:
            pass
