import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from frontend.gui_qrcode_window import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv) # QT 객체 생성
    window = MainWindow() # 메인 윈도우 객체 생성
    window.show() # 윈도우 표시
    app.exec() # 실행.