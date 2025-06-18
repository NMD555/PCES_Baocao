import sys
from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow

if __name__ == '__main__':
	app = QApplication(sys.argv)
	window = MainWindow()
	app.setStyle("Fusion")  # Buộc dùng giao diện Fusion
	window.resize(500,250)
	window.show()
	sys.exit(app.exec_())