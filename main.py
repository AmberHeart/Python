import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
import ui


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    gui = ui.Ui_mainWindow()
    gui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
    
    
