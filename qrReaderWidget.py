import cv2

from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
from PyQt6.QtGui import QImage, QPixmap
from PyQt6.QtCore import QTimer
from pyzbar.pyzbar import decode
from datetime import datetime


import pandas as pd
font = cv2.FONT_HERSHEY_SIMPLEX

class CameraViewer(QDialog):
    # attendanceUpdated = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        # 시트 가져오기
        self.current_sheet = self.parent().current_sheet
        self.setWindowTitle("Camera Viewer")

        self.processed_students = set()

        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 선택된 파일 열기
        self.df_daily = pd.read_excel(self.file_path, sheet_name='Student_Info&Daily_Attendance')
        self.df_semester = pd.read_excel(self.file_path,sheet_name='Semester_Attendance')

        self.sync_semester_with_daily()

        self.qr_label = QLabel(self)
        layout.addWidget(self.qr_label)

        # QLabel to display video
        self.video_label = QLabel(self)
        layout.addWidget(self.video_label)

        # QLabel to display QR code data

        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close_camera)
        layout.addWidget(close_btn)

        # Initialize VideoCapture and QTimer
        self.capture = cv2.VideoCapture(0)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(30)  # Update every 30 ms
        self.loading_widget.close()

    def sync_semester_with_daily(self):
        """Sync df_semester with df_daily"""
        # Create a DataFrame for students from df_daily
        daily_students_df = self.df_daily[['No', 'Name']].drop_duplicates()

        # Create a dictionary from df_semester for quick lookups
        if 'No' in self.df_semester.columns and 'Name' in self.df_semester.columns:
            semester_students_df = self.df_semester[['No', 'Name']]
            semester_students_dict = semester_students_df.set_index(['No', 'Name']).to_dict(orient='index')
        else:
            semester_students_dict = {}

        # Iterate over daily students and ensure they are in df_semester
        new_rows = []
        for _, student in daily_students_df.iterrows():
            student_id = student['No']
            student_name = student['Name']
            key = (student_id, student_name)

            if key not in semester_students_dict:
                new_row = pd.Series(index=self.df_semester.columns, dtype='object')
                new_row['No'] = student_id
                new_row['Name'] = student_name
                new_rows.append(new_row)

        if new_rows:
            self.df_semester = pd.concat([self.df_semester, pd.DataFrame(new_rows)], ignore_index=True)

        # Ensure columns for dates are present
        current_date = datetime.now().strftime("%m/%d")
        if current_date not in self.df_semester.columns:
            self.df_semester[current_date] = pd.NA

    def update_frame(self):
        if not self.capture.isOpened():
            return

        ret, frame = self.capture.read()
        if ret:

            frame = self.read_frame(frame)
            rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            flip_image = cv2.flip(rgb_image, 1)
            h, w, ch = flip_image.shape

            q_img = QImage(flip_image.data, w, h, ch * w, QImage.Format.Format_RGB888)
            # Set the image to the QLabel
            self.video_label.setPixmap(QPixmap.fromImage(q_img))

    def read_frame(self,frame):
        try:
            barcodes = decode(frame)
            for barcode in barcodes:
                x,y,w,h = barcode.rect
                barcode_info = barcode.data.decode('utf-8')
                barcode_array = barcode_info.split(',')

                timestamp = datetime.now()
                self.current_day = timestamp.strftime("%A")
                self.current_date = timestamp.strftime("%m/%d")
                self.current_time = timestamp.strftime("%H:%M")
                #barcode_array[0]는 학번, barcode_array[1]은 이름
                display_text = f"{barcode_array[0]}_{barcode_array[1]}:{self.current_time}"
                self.current_no = barcode_array[0]
                self.current_name = barcode_array[1]
                cv2.rectangle(frame, (x,y), (x+w,y+h), (0,0,255),2)
                cv2.putText(frame,display_text, (x,y - 20), font, 0.5, (0,0,255),1)
                if self.current_no not in self.processed_students:
                    self.updateAttendance()
            return frame
        except Exception as e:
            print(e)

    def close_camera(self):
        if self.capture.isOpened():
            self.capture.release()
        self.timer.stop()
        self.close()

    def closeEvent(self,event):
        with pd.ExcelWriter(self.file_path, engine='openpyxl') as writer:
            self.df_daily.to_excel(writer, sheet_name='Student_Info&Daily_Attendance', index=False)
            self.df_semester.to_excel(writer, sheet_name='Semester_Attendance', index=False)
        if self.capture.isOpened():
            self.capture.release()
        self.timer.stop()
        super().closeEvent(event)

    def updateAttendance(self):
        student_id = self.current_no
        student_name = self.current_name
        timestamp = self.current_time

        if not student_id or not student_name:
            return

        # 출석하지 않은 학생중, 출석부에 있는지 확인
        self.df_daily['No'] = self.df_daily['No'].astype(str).str.strip()
        self.df_daily['Name'] = self.df_daily['Name'].astype(str).str.strip()
        self.df_daily['TimeStamp'] = self.df_daily['TimeStamp']
        
        student_row = self.df_daily[(self.df_daily['No'] == student_id) & (self.df_daily['Name'] == student_name)]

        
        if not student_row.empty:
            student_index = student_row.index[0]
            timestamp_column = 'TimeStamp'

            if pd.isna(self.df_daily.at[student_index, timestamp_column]) or self.df_daily.at[student_index, timestamp_column] == "":
                self.df_daily.at[student_index, timestamp_column] = str(timestamp)
                self.processed_students.add(student_id)
                self.qr_label.setText(f"{student_id} Attendance has been recorded. Timestamp: {timestamp}")
            else:
                self.qr_label.setText(f"{student_id} Already Checked. Timestamp: {self.df_daily.at[student_index, timestamp_column]}")
                return
        else:
            self.qr_label.setText("Not existed in Attendance list, Please contact the administrator.")
            return

        # Find or add the column in the semester attendance DataFrame
        date_column = self.current_date
        if date_column not in self.df_semester.columns:
            self.df_semester[date_column] = pd.NA

        # Update the semester attendance DataFrame
        if student_index in self.df_semester.index:
            self.df_semester.at[student_index, date_column] = str(timestamp)
        else:
            new_row = pd.Series(index=self.df_semester.columns, dtype='object')
            new_row['No'] = student_id
            new_row['Name'] = student_name
            new_row[date_column] = str(timestamp)
            
            self.df_semester = pd.concat([self.df_semester, pd.DataFrame([new_row])], ignore_index=True)