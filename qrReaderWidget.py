import cv2
import re
import os
import sys
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QStackedLayout, QLabel, QPushButton, QComboBox
from PyQt6.QtGui import QImage, QPixmap, QIcon, QFont, QMovie
from PyQt6.QtCore import QTimer, pyqtSignal, Qt, QThread, QByteArray
from pyzbar.pyzbar import decode
from datetime import datetime
import logging
import pandas as pd
import numpy as np
from PIL import ImageFont, ImageDraw, Image
date_format = r"^\(\d{6}\)$"

class CameraThread(QThread):
    frameCaptured = pyqtSignal(np.ndarray)
    cameraReady = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.capture = None
        self.running = False

    def run(self):
        self.capture = cv2.VideoCapture(0)  # 카메라 장치 열기
        self.running = True
        
        while not self.capture.isOpened():
            cv2.waitKey(100)
        
        self.cameraReady.emit()

        while self.running:
            ret, frame = self.capture.read()
            if ret:
                # OpenCV BGR 포맷을 QImage의 RGB 포맷으로 변환
                frame = cv2.flip(frame, 1)
                
                # 신호를 통해 메인 스레드에 프레임 전달
                self.frameCaptured.emit(frame)

        self.capture.release()

    def stop(self):
        self.running = False
        self.wait()

# logging.basicConfig(level=logging.DEBUG)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, relative_path)

font_path = resource_path("meiryo.ttc")
logo_path = resource_path("scan_logo.png")
loading_path = resource_path("loading.gif")

logging.debug("font_path:", font_path)
try:
    font = ImageFont.truetype(font_path,25)
    message_font = ImageFont.truetype(font_path,15)
except IOError:
    font = ImageFont.load_default()

META_FILE = "metainfo.json"

class CameraViewer(QDialog):
    qrProcessed = pyqtSignal()
    last_processed_time = 0
    processing_interval = 2
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.resize(600,600)
        self.current_sheet = self.parent().current_sheet
        self.file_path = self.parent().file_path
        self.setWindowTitle("Camera Viewer")
        self.setWindowIcon(QIcon(logo_path))
        logging.debug("logo_path:", logo_path)
        self.selected = None
        self.current_no = ""
        self.current_name = ""

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.isChanged = False

        self.df_sheet = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
        self.sheet_label = QLabel(self)
        self.sheet_label.setFont(QFont(font_path,15))
        self.sheet_label.setText(f"Current Class: {self.current_sheet}")
        layout.addWidget(self.sheet_label)

        self.message_label = QLabel(self)
        self.message_label.setStyleSheet("color: green;")
        self.message_label.setFont(QFont(font_path,20))
        self.message_label.setText(" ")
        layout.addWidget(self.message_label)

        self.select_time = QComboBox()
        self.load_times()
        self.select_time.currentIndexChanged.connect(self.update_times)
        layout.addWidget(self.select_time)

        self.stack_layout = QStackedLayout()
        self.stack_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # video_layout = QHBoxLayout()
        # video_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        ## 카메라 프레임 표시
        self.video_label = QLabel(self)
        self.video_label.setMinimumSize(640,480)
        self.video_label.setStyleSheet("border: 2px solid black;")
        self.stack_layout.addWidget(self.video_label)

         # Loading label
        self.loading_label = QLabel(self)
        self.movie = QMovie(loading_path, QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheMode.CacheAll)

        self.loading_label.setMovie(self.movie)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stack_layout.addWidget(self.loading_label)
        # self.loading_label.setVisible(False)
        self.movie.start()

        layout.addLayout(self.stack_layout)
        
        self.camera_thread = CameraThread(parent=self)
        self.camera_thread.frameCaptured.connect(self.update_frame)
        self.camera_thread.cameraReady.connect(self.on_camera_ready)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close_camera)
        layout.addWidget(close_btn)

        # Create a QTimer to hide the message after the specified duration
        self.message_timer = QTimer()
        self.message_timer.setSingleShot(True)
        self.message_timer.timeout.connect(lambda: self.message_label.setText(" "))


    def show_loading(self, show=True):
        if show:
            self.stack_layout.setCurrentWidget(self.loading_label)
        else:
            self.stack_layout.setCurrentWidget(self.video_label)
    def on_camera_ready(self):
        self.show_loading(False)

    def update_times(self):
        self.selected = self.select_time.currentText().strip()
        self.df_sheet[self.selected] = self.df_sheet[self.selected].astype(str)
        if self.selected != "Select 回目":
            if not self.camera_thread.isRunning():
                self.show_loading(True)
            self.start_camera()

    def load_times(self):
        self.select_time.addItem("Select 回目")
        total_time = [col for col in self.df_sheet.columns if "回目" in col]
        self.select_time.addItems(total_time)
    
    def start_camera(self):
        if not self.camera_thread.isRunning():
            self.camera_thread.start()

    def update_frame(self, q_img:np.ndarray):
        frame = self.read_frame(q_img)
        rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb_image.shape
        q_img = QImage(rgb_image.data, w, h, ch * w, QImage.Format.Format_RGB888)
        self.video_label.setPixmap(QPixmap.fromImage(q_img))

    def read_frame(self, frame):
        try:
            barcodes = decode(frame)
            for barcode in barcodes:
                x, y, w, h = barcode.rect
                barcode_info = barcode.data.decode('utf-8')
                barcode_array = barcode_info.split(',') # 학번, 이름
                self.current_no = barcode_array[0] # 번호
                self.current_name = barcode_array[1] # 이름
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
                if (self.message_label.text() == " "):
                    self.updateAttendance()
            return frame
        except Exception as e:
            print(e)

    def close_camera(self):
        self.camera_thread.stop()
        self.camera_thread.quit()
        self.close()

    def closeEvent(self, event):
        self.close_camera()
        super().closeEvent(event)

        # 변경사항이 있으면 x 채우고 저장,없으면 그냥 닫기
        if not self.isChanged:
            super().closeEvent(event)
            return

        if self.current_sheet and self.selected:
            for idx, row in self.df_sheet.iterrows():
                if row[self.selected] != str('o'):
                    self.df_sheet.at[idx, self.selected] = str('x')
        if self.df_sheet[self.selected][0] == str('x'):
            date_now = datetime.now().strftime('%Y/%m/%d')
            self.df_sheet.at[0,self.selected] = str(date_now)
        with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            self.df_sheet.to_excel(writer, sheet_name=self.current_sheet, index=False)
            
        self.qrProcessed.emit()

    def update_column_name(self, col_name):
        start_idx = col_name.find('(')
        end_idx = col_name.find(')')
        current_date = datetime.now().strftime("%y%m%d")
        if start_idx != -1 and end_idx != -1:
            new_name = f"{col_name[:start_idx]({current_date})}"
            return new_name
        else:
            new_name = f"{col_name}({current_date})"
            return new_name

    def show_temporary_message(self, message, duration=4000):
        if self.message_label.isVisible():
            # If a message is already visible, stop the current timer
            self.message_timer.stop()

        self.message_label.setText(message)
        self.message_timer.start(duration)

    def updateAttendance(self):
        student_id = self.current_no
        student_name = self.current_name

        if not student_id or not student_name:
            return
        
        # Ensure IDs and names are stripped of extra whitespace
        self.df_sheet['学籍番号'] = self.df_sheet['学籍番号'].astype(str).str.strip()
        self.df_sheet['氏名'] = self.df_sheet['氏名'].astype(str).str.strip()

        student_row = self.df_sheet[
            (self.df_sheet['学籍番号'] == student_id) & (self.df_sheet['氏名'] == student_name)
        ]

        if not student_row.empty:
            student_index = student_row.index[0]
            current_value = self.df_sheet.at[student_index, self.selected]

            if current_value != str('o'):
                self.df_sheet.at[student_index, self.selected] = str('o')
                # self.processed_students.add(student_id)
                # self.qr_label.setText(f"{student_id} Attendance has been recorded.")
                self.show_temporary_message(f"{student_id}, {student_name} Attendance has been recorded.")
            else:
                self.show_temporary_message(f"{student_id}, {student_name} Already Checked.")
                # self.show_temporary_message(f"{student_id}, {student_name} Attendance has been recorded.")
                return
        else:
            self.show_temporary_message("Not existed in Attendance list, Please contact the administrator.")
            return
        if self.isChanged == False:
            self.isChanged = True
    