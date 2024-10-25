from PyQt5.QtWidgets import QApplication
from gui.aura_snu import *
 
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AURA_SNU()
    ex.show()
    sys.exit(app.exec_())