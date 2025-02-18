import cv2
import re
import os
import sys
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QStackedLayout, QLabel, QPushButton, QComboBox
from PyQt6.QtGui import QImage, QPixmap, QIcon, QFont, QMovie
from PyQt6.QtCore import QTimer, pyqtSignal, Qt, QThread, QByteArray, QUrl
from PyQt6.QtMultimedia import QSoundEffect
from pyzbar.pyzbar import decode
from datetime import datetime
import logging
import pandas as pd
import numpy as np
from PIL import ImageFont, ImageDraw, Image
import json

# 락 상태 해결을 위해 스레드 사용
class CameraThread(QThread):
    # 프레임 캡처 시그널, 카메라 준비 시그널
    frameCaptured = pyqtSignal(np.ndarray)
    cameraReady = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.capture = None
        self.running = False

    def run(self):
        self.capture = cv2.VideoCapture(0)  # 카메라 장치 열기
        self.running = True
        
        # 열릴때까지 대기
        while not self.capture.isOpened():
            cv2.waitKey(100)
        
        # 카메라 준비완료 시그널
        self.cameraReady.emit()

        while self.running:
            ret, frame = self.capture.read()
            if ret:
                # 프레임 좌우 반전
                frame = cv2.flip(frame, 1)
                
                # 메인 스레드로 프레임 전달
                self.frameCaptured.emit(frame)
        # 카메라 리소스 해제
        self.capture.release()

    def stop(self):
        # 카메라 스레드 종료 요청
        self.running = False
        self.wait()

# 리소스 경로를 절대 경로로 반환
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS # PyInstaller 로 패키징된 경우 사용하는 base_path
    except AttributeError:
        base_path = os.path.abspath('.') # 개발 중에서 사용하는 base_path
    return os.path.join(base_path, relative_path)

font_path = resource_path("meiryo.ttc")
logo_path = resource_path("scan_logo.png")
loading_path = resource_path("loading.gif")
sound_path = resource_path("beep.wav")

logging.debug("font_path:", font_path)

# 폰트 설정, 파일을 로드할 수 없으면 기본 폰트 사용
try:
    font = ImageFont.truetype(font_path,25)
    message_font = ImageFont.truetype(font_path,15)
except IOError:
    font = ImageFont.load_default()

class CameraViewer(QDialog):
    qrProcessed = pyqtSignal()
    last_processed_time = 0
    processing_interval = 2

    def __init__(self, parent=None):
        super().__init__(parent)
        self.json_decoder = json.decoder.JSONDecoder()
        # 윈도우 크기 및 타이틀 설정
        self.resize(600,600)
        self.current_sheet = self.parent().current_sheet
        self.file_path = self.parent().file_path
        self.setWindowTitle("Camera Viewer")
        self.setWindowIcon(QIcon(logo_path))
        logging.debug("logo_path:", logo_path) # 아이콘 설정

        # 현재 선택된 학생 정보를 초기화
        self.selected = None
        self.current_no = ""
        self.current_name = ""

        # 출석부 실행 시간 저장
        self.current_date = datetime.now().strftime("%Y/%m/%d")

        # 레이아웃 설정
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.isChanged = False # 출석 변경 여부 확인 -> 변경사항확인용

        # 시트 데이터 읽기(parent로 부터 받아옴)
        self.df_sheet = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
        self.sheet_label = QLabel(self)
        self.sheet_label.setFont(QFont(font_path,15))
        self.sheet_label.setText(f"Current Class: {self.current_sheet}")
        layout.addWidget(self.sheet_label)

        # QR인증시 보여지는 라벨 설정(QR 정보 출력)
        self.message_label = QLabel(self)
        self.message_label.setStyleSheet("color: green;")
        self.message_label.setFont(QFont(font_path,20))
        self.message_label.setText(" ")
        layout.addWidget(self.message_label)

        # 몇 차시인지 선택하는 콤보박스 -> 250211 에서는 필요하지 않아 제거.
        # self.select_time = QComboBox()
        # self.load_times() # 
        # self.select_time.currentIndexChanged.connect(self.update_times)
        # layout.addWidget(self.select_time)
        
        # 카메라 화면 및 로딩 화면을 스택하기 위한 레이아웃 ( 카메라 화면 앞에 로딩 이미지를 겹쳐 불러옴)
        self.stack_layout = QStackedLayout()
        self.stack_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
    
        ## 카메라 프레임 표시
        self.video_label = QLabel(self)
        self.video_label.setMinimumSize(640,480)
        self.video_label.setStyleSheet("border: 2px solid black;")
        self.stack_layout.addWidget(self.video_label)

        # 로딩 화면 표시
        self.loading_label = QLabel(self)
        self.movie = QMovie(loading_path, QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheMode.CacheAll)

        self.loading_label.setMovie(self.movie)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stack_layout.addWidget(self.loading_label)
        self.movie.start()
        layout.addLayout(self.stack_layout)

        # 카메라 스레드 초기화 설정        
        self.camera_thread = CameraThread(parent=self)
        self.camera_thread.frameCaptured.connect(self.update_frame) # 카메라 프레임 시그널 처리
        self.camera_thread.cameraReady.connect(self.on_camera_ready) # 카메라 준비완료 시그널 처리

        # 종료 버튼
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close_camera)
        layout.addWidget(close_btn)

        # 소리 설정( 출석 체크 시 소리 재생 )
        self.sound = QSoundEffect()
        self.sound.setSource(QUrl.fromLocalFile(sound_path))
        self.sound.setLoopCount(1)    

        # 일정 시간 후 QR 정보창을 숨김
        self.message_timer = QTimer()
        self.message_timer.setSingleShot(True)
        self.message_timer.timeout.connect(lambda: self.message_label.setText(" "))

        self.camera_Run()


    def show_loading(self, show=True):
        # 로딩 화면 표시
        if show:
            self.stack_layout.setCurrentWidget(self.loading_label)
        else:
            self.stack_layout.setCurrentWidget(self.video_label)
    
    def on_camera_ready(self):
        # 카메라 준비 완료 시 로딩 화면 숨기기
        self.show_loading(False)

    # 차수 선택하는 기능 제거
    def update_times(self):
        # 시간 선택 변경 시
        self.selected = self.select_time.currentText().strip()
        self.df_sheet[self.selected] = self.df_sheet[self.selected].astype(str)
        if self.selected != "Select 回目":
            if not self.camera_thread.isRunning():
                self.show_loading(True)
            self.start_camera()
    def camera_Run(self):
        if not self.camera_thread.isRunning():
            self.show_loading(True)
        self.start_camera()

    def load_times(self):
        # 시간 목록 로딩
        self.select_time.addItem("Select 回目")
        total_time = [col for col in self.df_sheet.columns if "回目" in col]
        self.select_time.addItems(total_time)
    
    def start_camera(self):
        # 카메라 시작
        if not self.camera_thread.isRunning():
            self.camera_thread.start()

    def update_frame(self, q_img:np.ndarray):
        # 카메라 프레임 업데이트
        frame = self.read_frame(q_img)
        if frame is None:
            # print("Error: Frame empty")
            return
        try:
            rgb_image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        except cv2.error as e:
            print(e)
            return
        h, w, ch = rgb_image.shape
        q_img = QImage(rgb_image.data, w, h, ch * w, QImage.Format.Format_RGB888)
        self.video_label.setPixmap(QPixmap.fromImage(q_img))

    def read_frame(self, frame):
        # 프레임에서 QR 코드 디코딩
        try:
            barcodes = decode(frame)
            for barcode in barcodes:
                x, y, w, h = barcode.rect
                barcode_info = barcode.data.decode('unicode_escape')
                info = self.json_decoder.decode(barcode_info)
                grade = info['学年']
                student_no = info['学籍番号']
                name1 = info['氏名']
                name2 = info['カナ']
                email = info['メールアドレス']
                self.current_no = student_no
                self.current_name = name1
                # barcode_array = barcode_info.split(',') # 학번, 이름
                # self.current_no = barcode_array[0] # 번호
                # self.current_name = barcode_array[1] # 이름
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
                # QR 코드가 읽혔을 때 출석 처리
                if (self.message_label.text() == " "):
                    self.updateAttendance()
            return frame
        except Exception as e:
            print(e)

    def close_camera(self):
        # 카메라 스레드 종료
        self.camera_thread.stop()
        self.camera_thread.quit()
        self.close()

    def closeEvent(self, event):
        # 창 닫을 때 카메라 종료 및 출석체크 처리
        self.close_camera()
        super().closeEvent(event)

        # 변경사항이 있으면 시트에 반영, 없으면 그냥 종료
        if not self.isChanged:
            super().closeEvent(event)
            return

        # 종합 출결 시트(出席調査)에 변경사항이 있으면 변경사항 저장 -> 나중에 구현
        # self.df_sheet 의 내용을 出席調査시트에 반영
        attendance_sheet = pd.read_excel(self.file_path, sheet_name="出席調査")
        for index, row in self.df_sheet.iterrows():
            student_id = row['学籍番号']
            student_name = row['氏名']
            class_name = self.current_sheet
            # 学籍番号(학번), 氏名(이름), クラス名(과목명)으로 학생 찾기기
            student_row = attendance_sheet[(attendance_sheet['学籍番号'] == student_id) & (attendance_sheet['氏名'] == student_name) & (attendance_sheet['クラス名'] == class_name)]
            if not student_row.empty:
                student_index = student_row.index[0]
                # 
                attendance_sheet.at[student_index, '授業回数'] = row['授業回数'] # 출석 일수 업데이트
                attendance_sheet.at[student_index, '欠席数'] = row['欠席数'] # 결석 일수 업데이트

        # 변경사항 반영
        with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            self.df_sheet.to_excel(writer, sheet_name=self.current_sheet, index=False)
            attendance_sheet.to_excel(writer, sheet_name="出席調査", index=False)
        self.qrProcessed.emit()

    # def update_column_name(self, col_name):
    #     start_idx = col_name.find('(')
    #     end_idx = col_name.find(')')
    #     current_date = datetime.now().strftime("%y%m%d")
    #     if start_idx != -1 and end_idx != -1:
    #         new_name = f"{col_name[:start_idx]({current_date})}"
    #         return new_name
    #     else:
    #         new_name = f"{col_name}({current_date})"
    #         return new_name

    def show_temporary_message(self, message, duration=3500):
        # 출석 체크 메세지 표시 
        if self.message_label.isVisible():
            # 메세지가 비활성화일때 타이머 중지
            self.message_timer.stop()
        self.message_label.setText(message)
        self.message_timer.start(duration)
    # 授業回数(출석일수) 열 업데이트, 결석 일수 업데이트는 시트 종료시에 처리(?) 잘 모르겠음
    def updateAttendance(self):
        # 출석 처리
        student_id = self.current_no
        student_name = self.current_name
        print(student_id, student_name)
        if not student_id or not student_name:
            return
        
        # 학번, 이름 공백 제거
        self.df_sheet['学籍番号'] = self.df_sheet['学籍番号'].astype(str).str.strip()
        self.df_sheet['氏名'] = self.df_sheet['氏名'].astype(str).str.strip()

        # 시트(데이터 프레임)에서 일치하는 학생 찾기
        student_row = self.df_sheet[
            (self.df_sheet['学籍番号'] == student_id) & (self.df_sheet['氏名'] == student_name)
        ]
        if not student_row.empty:
            student_index = student_row.index[0]
            current_value = student_row['출석시간'].values[0]

            try:
                current_value = current_value.split()[0]
            except:
                print("날짜 데이터가 아님.")

            # 시간 기준으로 출석 처리 같은 날 출석 처리되었으면 다시 출석 처리하지 않음
            if current_value != self.current_date:
                # 출석 처리
                self.df_sheet.at[student_index, '출석시간'] = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
                # 초기값이 없는 경우 0으로 초기화
                if pd.isna(self.df_sheet.at[student_index, '授業回数']):
                    self.df_sheet.at[student_index, '授業回数'] = 0  # 초기값 0으로 설정
                self.df_sheet.at[student_index, '授業回数'] += 1
                # self.processed_students.add(student_id)
                # self.qr_label.setText(f"{student_id} Attendance has been recorded.")
                self.show_temporary_message(f"{student_id}, {student_name} Attendance has been recorded.")
                self.sound.play()
                self.qrProcessed.emit()
            else:
                self.show_temporary_message(f"{student_id}, {student_name} Already Checked.")
                # self.show_temporary_message(f"{student_id}, {student_name} Attendance has been recorded.")
                return
        else:
            self.show_temporary_message("Not existed in Attendance list, Please contact the administrator.")
            return
        
        if self.isChanged == False:
            self.isChanged = True
    