from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from geo_documents.window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("GEO Documents")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
