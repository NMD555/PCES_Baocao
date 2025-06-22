import sys
from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow
import qdarkstyle

if __name__ == '__main__':
	app = QApplication(sys.argv)
	app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
	window = MainWindow()
	app.setStyle("Fusion")  # Buộc dùng giao diện Fusion
	window.resize(500,250)
	window.show()
	sys.exit(app.exec_())