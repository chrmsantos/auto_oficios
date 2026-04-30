"""Entry point for ``python -m z7_officeletters``.

Launches the graphical interface.  The module is also used as the
PyInstaller analysis script so that the exe bundle starts the GUI.
"""

from __future__ import annotations

import sys
from pathlib import Path

# When executed directly (python __main__.py) the src/ directory is not on
# sys.path; add it so that the z7_officeletters package can be found.
_src = Path(__file__).parent.parent
if str(_src) not in sys.path:
    sys.path.insert(0, str(_src))


def main() -> None:
    """Start the GUI application."""
    from z7_officeletters.gui.app import AutoOficiosApp

    app = AutoOficiosApp()
    app.mainloop()


if __name__ == "__main__":
    main()
