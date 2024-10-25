from .analysis_window import *
import traceback
import sys
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QFileDialog, QLabel, QPushButton, QVBoxLayout, QWidget, QMainWindow

class AURA_SNU(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_path = ''
        self.event_path = ''
        self.result_path = ''  # 결과 폴더 경로 추가
        self.initUI()
        sys.excepthook = self.globalExceptionHandler

    def globalExceptionHandler(self, exc_type, exc_value, exc_traceback):
        traceback_str = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(f"Unhandled exception: {traceback_str}")

    def initUI(self):
        self.setWindowTitle('RWA Analysis Tool')
        self.setGeometry(300, 300, 600, 500)
        self.setStyleSheet("background-color: #f0f0f0;")

        layout = QVBoxLayout()

        # Data File Selection
        dataLayout = QVBoxLayout()
        dataLayout.setSpacing(5)
        self.lblDataFile = QLabel('No Data File Selected', self)
        self.lblDataFile.setFont(QFont("Arial", 10))
        dataLayout.addWidget(self.lblDataFile)

        self.btnSelectData = QPushButton('Select Data File (.edf)', self)
        self.btnSelectData.clicked.connect(self.openDataFileDialog)
        self.btnSelectData.setStyleSheet("QPushButton {"
                                         "  background-color: #4CAF50;"
                                         "  color: white;"
                                         "  border-radius: 5px;"
                                         "  padding: 10px;"
                                         "  font-size: 14px;"
                                         "}"
                                         "QPushButton:hover {"
                                         "  background-color: #45a049;"
                                         "}")
        self.btnSelectData.setMinimumHeight(40)
        dataLayout.addWidget(self.btnSelectData)
        layout.addLayout(dataLayout)

        # Event File Selection
        eventLayout = QVBoxLayout()
        eventLayout.setSpacing(5)
        self.lblEventFile = QLabel('No Event File Selected', self)
        self.lblEventFile.setFont(QFont("Arial", 10))
        eventLayout.addWidget(self.lblEventFile)

        self.btnSelectEvent = QPushButton('Select Event File (.xlsx)', self)
        self.btnSelectEvent.clicked.connect(self.openEventFileDialog)
        self.btnSelectEvent.setStyleSheet("QPushButton {"
                                          "  background-color: #008CBA;"
                                          "  color: white;"
                                          "  border-radius: 5px;"
                                          "  padding: 10px;"
                                          "  font-size: 14px;"
                                          "}"
                                          "QPushButton:hover {"
                                          "  background-color: #007bb8;"
                                          "}")
        self.btnSelectEvent.setMinimumHeight(40)
        eventLayout.addWidget(self.btnSelectEvent)
        layout.addLayout(eventLayout)

        # Result Folder Selection
        resultLayout = QVBoxLayout()
        resultLayout.setSpacing(5)
        self.lblResultFolder = QLabel('No Result Folder Selected', self)
        self.lblResultFolder.setFont(QFont("Arial", 10))
        resultLayout.addWidget(self.lblResultFolder)

        self.btnSelectResultFolder = QPushButton('Select Result Folder', self)
        self.btnSelectResultFolder.clicked.connect(self.openResultFolderDialog)
        self.btnSelectResultFolder.setStyleSheet("QPushButton {"
                                                 "  background-color: #4A90E2;"
                                                 "  color: white;"
                                                 "  border-radius: 5px;"
                                                 "  padding: 10px;"
                                                 "  font-size: 14px;"
                                                 "}"
                                                 "QPushButton:hover {"
                                                 "  background-color: #357ABD;"
                                                 "}")
        self.btnSelectResultFolder.setMinimumHeight(40)
        resultLayout.addWidget(self.btnSelectResultFolder)
        layout.addLayout(resultLayout)

        # Start Analysis Button
        self.btnStartAnalysis = QPushButton('Start Analysis', self)
        self.btnStartAnalysis.clicked.connect(self.startAnalysis)
        self.btnStartAnalysis.setEnabled(False)
        self.btnStartAnalysis.setStyleSheet("QPushButton:enabled { background-color: #f44336; color: white; border-radius: 5px; }"
                                            "QPushButton:disabled { background-color: #e7e7e7; color: black; }"
                                            "QPushButton:hover:enabled { background-color: #d32f2f; }")
        layout.addWidget(self.btnStartAnalysis)

        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)

    def openDataFileDialog(self):
        options = QFileDialog.Options()
        self.data_path, _ = QFileDialog.getOpenFileName(self, 'Select Data File', filter="EDF files (*.edf)", options=options)
        if self.data_path:
            self.lblDataFile.setText(f"Data File: {self.data_path.split('/')[-1]}")
        self.checkFilesSelected()

    def openEventFileDialog(self):
        options = QFileDialog.Options()
        self.event_path, _ = QFileDialog.getOpenFileName(self, 'Select Event File', filter="Text files (*.xlsx)", options=options)
        if self.event_path:
            self.lblEventFile.setText(f"Event File: {self.event_path.split('/')[-1]}")
        self.checkFilesSelected()

    def openResultFolderDialog(self):
        options = QFileDialog.Options()
        self.result_path = QFileDialog.getExistingDirectory(self, 'Select Result Folder', options=options)
        if self.result_path:
            self.lblResultFolder.setText(f"Result Folder: {self.result_path}")
        self.checkFilesSelected()

    def checkFilesSelected(self):
        if self.data_path and self.event_path and self.result_path:
            self.btnStartAnalysis.setEnabled(True)

    def startAnalysis(self):
        # AnalysisWindow 클래스에 parent 매개변수로 self를 전달
        self.analysisWindow = AnalysisWindow(self.data_path, self.event_path, self.result_path, self)
        self.analysisWindow.show()
