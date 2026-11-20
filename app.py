"""Punto de entrada del programa (también usado por PyInstaller)."""
import sys

from inventario.gui import main

if __name__ == "__main__":
    sys.exit(main())
