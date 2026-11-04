"""PyInstaller entry point for the command-line tool (console)."""

import multiprocessing
import sys

if __name__ == "__main__":
    multiprocessing.freeze_support()
    from docgen.cli.main import main

    sys.exit(main())
