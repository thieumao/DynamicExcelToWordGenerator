"""PyInstaller entry point for the desktop application (windowed)."""

import multiprocessing
import sys

if __name__ == "__main__":
    multiprocessing.freeze_support()  # must run first: worker processes re-launch this executable
    from docgen.ui.app import main

    sys.exit(main())
