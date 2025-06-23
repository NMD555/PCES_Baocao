import sys
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow
import qdarkstyle

if __name__ == '__main__':
	app = QApplication(sys.argv)
	app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
	font = QFont("Segoe UI", 20)  # hoặc "Arial", 10pt tùy bạn
	app.setFont(font)
	window = MainWindow()
	app.setStyle("Fusion")  # Buộc dùng giao diện Fusion
	window.resize(500,250)
	window.show()
	sys.exit(app.exec_())