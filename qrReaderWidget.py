import cv2
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QComboBox
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtCore import QTimer, pyqtSignal
from pyzbar.pyzbar import decode
from datetime import datetime
import pandas as pd

font = cv2.FONT_HERSHEY_SIMPLEX

class CameraViewer(QDialog):
    qrProcessed = pyqtSignal(str,str)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.resize(800,600)
        self.current_sheet = self.parent().current_sheet
        self.file_path = self.parent().file_path
        self.setWindowTitle("Camera Viewer")

        # self.processed_students = set()
        self.selected = None

        layout = QVBoxLayout()
        self.setLayout(layout)


        self.df_sheet = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
        self.sheet_label = QLabel(self)
        self.sheet_label.setText(f"Current: {self.current_sheet}")
        self.qr_label = QLabel(self)
        layout.addWidget(self.qr_label)

        self.select_time = QComboBox()
        self.select_time.addItem("Select 回目")
        self.select_time.currentIndexChanged.connect(self.update_times)
        self.load_times()
        layout.addWidget(self.select_time)

        self.video_label = QLabel(self)
        layout.addWidget(self.video_label)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close_camera)
        layout.addWidget(close_btn)

        self.capture = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)

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
            self.timer.start(30)  # Update every 30 ms
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
                timestamp = datetime.now()
                self.current_day = timestamp.strftime("%A")  # 요일
                self.current_date = timestamp.strftime("%Y-%m-%d")  # 년월일
                display_text = f"{barcode_array[0]}_{barcode_array[1]}: Checked"
                self.current_no = barcode_array[0]
                self.current_name = barcode_array[1]
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                cv2.putText(frame, display_text, (x, y - 20), font, 0.5, (0, 0, 255), 1)
                # if self.current_no not in self.processed_students:
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

        if self.current_sheet and self.selected:
            for idx, row in self.df_sheet.iterrows():
                if row[self.selected] != str('o'):
                    self.df_sheet.at[idx, self.selected] = str('x')
        with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            self.df_sheet.to_excel(writer, sheet_name=self.current_sheet, index=False)
        
        super().closeEvent(event)

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
                self.qr_label.setText(f"{student_id} Attendance has been recorded.")
            else:
                self.qr_label.setText(f"{student_id} Already Checked.")
        else:
            self.qr_label.setText("Not existed in Attendance list, Please contact the administrator.")
        
        self.qrProcessed.emit(student_id, student_name)

