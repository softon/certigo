from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QTimer
import sys
import certigo_gui

app = QApplication(sys.argv)
splash_pix = QPixmap("icon.png")
splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
splash.show()

# Load main GUI after delay
QTimer.singleShot(2000, lambda: (
    splash.close(),
    certigo_gui.CertigoGUI().show()
))

sys.exit(app.exec_())
