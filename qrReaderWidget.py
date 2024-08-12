import cv2
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QComboBox
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtCore import QTimer, pyqtSignal
from pyzbar.pyzbar import decode
from datetime import datetime

import pandas as pd
import numpy as np
from PIL import ImageFont, ImageDraw, Image




font = cv2.FONT_HERSHEY_SIMPLEX

class CameraViewer(QDialog):
    qrProcessed = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.resize(800,600)
        self.current_sheet = self.parent().current_sheet
        self.file_path = self.parent().file_path
        self.setWindowTitle("Camera Viewer")

        self.selected = None
        self.current_no = None
        self.current_name = None

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.isChanged = False

        self.df_sheet = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
        self.sheet_label = QLabel(self)
        self.sheet_label.setText(f"Current Class: {self.current_sheet}")
        layout.addWidget(self.sheet_label)

        self.message_label = QLabel(self)
        self.message_label.setStyleSheet("color: green;")
        self.message_label.setVisible(False)
        layout.addWidget(self.message_label)

        self.select_time = QComboBox()
        self.select_time.addItem("Select 回目")
        self.select_time.currentIndexChanged.connect(self.update_times)
        self.load_times()
        layout.addWidget(self.select_time)

        self.video_label = QLabel(self)
        # self.video_label.setFixedSize(300,300)
        layout.addWidget(self.video_label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close_camera)
        layout.addWidget(close_btn)

        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)

        self.message_timer = QTimer()
        self.message_timer.setSingleShot(True)
        self.message_timer.timeout.connect(lambda: self.message_label.setVisible(False))


    def update_times(self):
        self.selected = self.select_time.currentText()
        self.df_sheet[self.selected] = self.df_sheet[self.selected].astype(str)
        if self.selected != "Select 回目":
            self.start_camera()

    def load_times(self):
        total_time = [col for col in self.df_sheet.columns if "回目" in col]
        self.select_time.addItems(total_time)

    def start_camera(self):
        if self.capture and self.capture.isOpened():
            self.capture.release()  # Release the previous capture if open
        self.capture = cv2.VideoCapture(0)
        if self.capture.isOpened():
            # self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 400)
            # self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT,400)
            self.timer.start(100)  # Update every 30 ms
        else:
            print("Error: Unable to open camera")

    def update_frame(self):
        if not self.capture or not self.capture.isOpened():
            return

        ret, frame = self.capture.read()
        if ret:
            frame = cv2.flip(frame, 1)
            frame = self.read_frame(frame)

            rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil_image = Image.fromarray(rgb_image)

            try:
                font = ImageFont.truetype("./font/NotoSansJP-Regular.ttf",35)
            except IOError:
                font = ImageFont.load_default()
            
            draw = ImageDraw.Draw(pil_image)
            draw.text((30,30), f"{self.current_no}_{self.current_name}", font=font, fill=(0, 0, 0))
            # print( f"{self.current_no}_{self.current_name}")
            image = np.asarray(pil_image)

            h, w, ch = image.shape

            q_img = QImage(image.data, w, h, ch * w, QImage.Format.Format_RGB888)
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
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                self.updateAttendance()
            return frame
        except Exception as e:
            print(e)

    def close_camera(self):
        if self.capture and self.capture.isOpened():
            self.capture.release()
        self.timer.stop()
        self.close()

    def closeEvent(self, event):
        if self.capture and self.capture.isOpened():
            self.capture.release()
        self.timer.stop()

        # 변경사항이 있으면 x 채우고 저장,없으면 그냥 닫기
        if not self.isChanged:
            super().closeEvent(event)
            return

        self.df_sheet.rename(index = {self.selected:self.update_column_name(self.selected)}, inplace = True)

        if self.current_sheet and self.selected:
            for idx, row in self.df_sheet.iterrows():
                if row[self.selected] != str('o'):
                    self.df_sheet.at[idx, self.selected] = str('x')
        with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            self.df_sheet.to_excel(writer, sheet_name=self.current_sheet, index=False)
        
        super().closeEvent(event)
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

    def show_temporary_message(self, message, duration=3000):
        if self.message_label.isVisible():
            # If a message is already visible, stop the current timer
            self.message_timer.stop()

        self.message_label.setText(message)
        self.message_label.setVisible(True)

        # Create a QTimer to hide the message after the specified duration
        self.message_timer = QTimer()
        self.message_timer.setSingleShot(True)
        self.message_timer.timeout.connect(lambda: self.message_label.setVisible(False))
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
                self.show_temporary_message(f"{student_id} Attendance has been recorded.")
            else:
                self.show_temporary_message(f"{student_id} Already Checked.")
                # self.show_temporary_message(f"{student_id} Attendance has been recorded.")
                return
        else:
            self.show_temporary_message("Not existed in Attendance list, Please contact the administrator.")
            return
        if self.isChanged == False:
            self.isChanged = True
        

