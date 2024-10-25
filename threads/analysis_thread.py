import traceback
from mne import io as ioo
import lazy_loader
from PyQt5.QtCore import QThread, pyqtSignal
from utils.aura_main import AURA_main

class AnalysisThread(QThread):
    analysisFinished = pyqtSignal()  # 분석 완료 신호
    analysisError = pyqtSignal(str)  # 오류 신호
    def __init__(self, data_path, event_path, result_path):
        super().__init__()
        self.data_path = data_path
        self.event_path = event_path
        self.result_path = result_path

    def run(self):
        try:
            data = ioo.read_raw_edf(self.data_path, preload=True)
            info = data.info
            f_s = int(info["sfreq"])
            data_start = info['meas_date']
            channels = data.ch_names
            data_length = len(data[0][1])
            # Call the main function from main.py
            AURA_main(data, self.result_path , f_s, data_start, channels, data_length, self.event_path)
            self.analysisFinished.emit()
        except Exception as e:
            traceback_str = ''.join(traceback.format_exception(None, e, e.__traceback__))
            print(f"Error during analysis: {traceback_str}")
            self.analysisError.emit(str(e))