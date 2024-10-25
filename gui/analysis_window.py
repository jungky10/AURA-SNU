  
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QPushButton,QWidget, QMessageBox, QTextEdit
from PyQt5.QtCore import  pyqtSlot, QTimer
import sys

from .stream_redirector import *
from threads.analysis_thread import AnalysisThread

class AnalysisWindow(QMainWindow):
    def __init__(self, data_path, event_path, result_path, parent=None):
        super().__init__(parent)
        self.data_path = data_path
        self.event_path = event_path
        self.result_path = result_path  # 결과 경로 추가
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Analysis Progress')
        self.setGeometry(350, 350, 300, 200)

        layout = QVBoxLayout()

        self.btnClose = QPushButton('Close', self)  # 닫기 버튼 추가
        self.btnClose.clicked.connect(self.close)
        self.btnClose.setEnabled(False)  # 초기에는 비활성화
        layout.addWidget(self.btnClose)

        self.consoleOutput = QTextEdit(self)
        self.consoleOutput.setReadOnly(True)
        layout.addWidget(self.consoleOutput)

        # 터미널 출력을 QTextEdit 위젯으로 리디렉션
        sys.stdout = StreamRedirector(self.consoleOutput)
        sys.stderr = StreamRedirector(self.consoleOutput)

        self.startAnalysis()

        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)

    def startAnalysis(self):
        self.analysisThread = AnalysisThread(self.data_path, self.event_path, self.result_path)
        self.analysisThread.analysisFinished.connect(self.onAnalysisFinished)
        self.analysisThread.analysisError.connect(self.onAnalysisError)
        self.analysisThread.start()

    @pyqtSlot()
    def onAnalysisFinished(self):
        self.showMessageBox("Analysis Completed", "The analysis has been completed.")
        self.btnClose.setEnabled(True)

    @pyqtSlot(str)
    def onAnalysisError(self, error_message):
        self.showMessageBox("Analysis Error", f"An error occurred during the analysis:\n{error_message}")
        self.btnClose.setEnabled(True)

    def showMessageBox(self, title, message):
        # QMessageBox 호출을 메인 스레드의 이벤트 루프에 예약
        QTimer.singleShot(0, lambda: QMessageBox.information(self, title, message))