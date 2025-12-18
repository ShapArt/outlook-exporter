from __future__ import annotations

import sys
import threading
import time

from config import default_config
from ui.app import run_ui


def main():
    cfg = default_config()
    # Run UI in a thread and close after timeout
    t = threading.Thread(target=lambda: run_ui(cfg), daemon=True)
    t.start()
    time.sleep(10)
    print("UI smoke: launched and alive for 10s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
