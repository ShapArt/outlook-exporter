from config import default_config
from ui.app import run_ui


def main():
    cfg = default_config()
    run_ui(cfg)


if __name__ == "__main__":
    main()
