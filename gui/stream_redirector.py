from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QObject

class StreamRedirector(QObject):
    textWritten = pyqtSignal(str)  # 새로운 시그널

    def __init__(self, widget):
        super().__init__()  # QObject의 초기화
        self.widget = widget
        self.textWritten.connect(self.widget.append)  # 시그널 연결

    def write(self, text):
        self.textWritten.emit(text)  # 시그널 발생

    def flush(self):
        pass
