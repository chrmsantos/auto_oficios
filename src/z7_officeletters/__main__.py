"""Entry point for ``python -m z7_officeletters``.

Launches the graphical interface.  The module is also used as the
PyInstaller analysis script so that the exe bundle starts the GUI.
"""

from __future__ import annotations


def main() -> None:
    """Start the GUI application."""
    from z7_officeletters.gui.app import AutoOficiosApp

    app = AutoOficiosApp()
    app.mainloop()


if __name__ == "__main__":
    main()
