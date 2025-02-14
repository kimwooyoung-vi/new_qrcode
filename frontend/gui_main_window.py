from PyQt6.QtCore import QMetaObject, QSize, Qt
from PyQt6.QtGui import QIcon, QPixmap, QFont
from PyQt6.QtWidgets import QLabel, QPushButton, QHBoxLayout, QWidget, QVBoxLayout, QFileDialog, QStatusBar, QSizePolicy
import openpyxl.drawing
import openpyxl.drawing.image
import qrcode.constants

from frontend.gui_email_window import ExcelDialog
from static.styles.styles import application_style, button_style, statusbar_style
from static.resources.resource_pathes.resource_pathes import save_icon_path, folder_icon_path

import shutil
import sys
import os
import pandas as pd
import openpyxl
from io import BytesIO
import qrcode
from qrcode.image.styledpil import StyledPilImage
from qrcode.image.styles.moduledrawers.pil import RoundedModuleDrawer
import json
from PIL import Image,ImageDraw, ImageFont

# 리소스 경로를 절대 경로로 반환
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS # PyInstaller 로 패키징된 경우 사용하는 base_path
    except AttributeError:
        base_path = os.path.abspath('.') # 개발 중에서 사용하는 base_path
    return os.path.join(base_path, relative_path)
font_path = resource_path("meiryo.ttc")
try:
    font = ImageFont.truetype(font_path,35)
except IOError:
    font = ImageFont.load_default()

def _get_icon(icon_path: str) -> QIcon:
    icon: QIcon = QIcon()
    icon.addPixmap(
        QPixmap(icon_path),
        QIcon.Mode.Normal,
        QIcon.State.Off
    )
    return icon

class GuiMainWindow:
    def setup_gui(self, MainWindow: QWidget) -> None:
        MainWindow.setObjectName("main_window")
        MainWindow.setWindowTitle("VisionInside QR")

        main_icon: QIcon = _get_icon("path/to/icon.png")  # 아이콘 경로
        icon_size = QSize(24, 24)
        MainWindow.setWindowIcon(main_icon)
        MainWindow.setIconSize(icon_size)

        self.central_widget = QWidget(MainWindow)
        self.central_widget.setObjectName("central_widget")
        self.central_widget.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
        )

        # 중앙 위젯의 레이아웃 설정
        self.main_layout = QVBoxLayout(self.central_widget)

        # 상단 레이아웃 (QHBoxLayout)
        self.header_layout = QHBoxLayout()

        # 제목 라벨 추가
        self.title_label = QLabel("VisionInside QR 시스템", self.central_widget)
        self.title_label.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.header_layout.addWidget(self.title_label)

        # 상단 레이아웃을 중앙 레이아웃에 추가
        self.main_layout.addLayout(self.header_layout)

        # 버튼 레이아웃
        self.button_layout = QVBoxLayout()

        # 엑셀 파일 선택 버튼 추가
        self.load_button = QPushButton("엑셀 파일 선택", self.central_widget)
        self.load_button.setIcon(_get_icon(icon_path=folder_icon_path))  # 버튼 아이콘
        self.load_button.setFont(QFont("Arial", 12))
        self.load_button.clicked.connect(self.select_excel_file)
        self.button_layout.addWidget(self.load_button)
        
        # 선택된 파일 표시를 위한 레이블 추가
        self.loaded_file_label = QLabel("선택된 파일: 아직 선택되지 않음", self.central_widget)
        self.loaded_file_label.setFont(QFont("Arial", 12))
        self.button_layout.addWidget(self.loaded_file_label)

        # 저장할 위치 선택 버튼 추가
        self.save_button = QPushButton("저장 위치 선택", self.central_widget)
        self.save_button.setIcon(_get_icon(icon_path=save_icon_path))  # 버튼 아이콘
        self.save_button.setFont(QFont("Arial", 12))
        self.save_button.clicked.connect(self.select_save_location)
        self.button_layout.addWidget(self.save_button)

        # 저장 위치 표시를 위한 레이블 추가
        self.save_location_label = QLabel("저장 위치: 아직 선택되지 않음", self.central_widget)
        self.save_location_label.setFont(QFont("Arial", 12))
        self.button_layout.addWidget(self.save_location_label)

        # SEPERATE 파일 생성 버튼 추가
        self.seperate_button = QPushButton("SEPERATE 파일 생성", self.central_widget)
        self.seperate_button.setFont(QFont("Arial", 12))
        self.seperate_button.clicked.connect(self.create_seperate_file)
        self.button_layout.addWidget(self.seperate_button)

        # 메일 전송 버튼 추가
        self.send_email_button = QPushButton("send mail", self.central_widget)
        self.send_email_button.clicked.connect(self.open_mail_window)
        self.button_layout.addWidget(self.send_email_button)

        # 종료 버튼 추가
        self.exit_button = QPushButton("종료", self.central_widget)
        self.exit_button.setIcon(_get_icon("path/to/exit_icon.png"))  # 버튼 아이콘
        self.exit_button.setFont(QFont("Arial", 12))
        self.exit_button.clicked.connect(MainWindow.close)
        self.button_layout.addWidget(self.exit_button)

        # 버튼 레이아웃을 중앙 레이아웃에 추가
        self.main_layout.addLayout(self.button_layout)

        # 상태 표시줄 (QStatusBar) 추가
        self.status_bar = QStatusBar(self.central_widget)
        self.main_layout.addWidget(self.status_bar)

        # 메인 윈도우에 중앙 위젯 설정
        MainWindow.setCentralWidget(self.central_widget)

        # 메타 객체로 연결 설정
        QMetaObject.connectSlotsByName(MainWindow)

        # 스타일 적용
        self.apply_styles(MainWindow)

    def open_mail_window(self):
        email_window = ExcelDialog(self)
        email_window.show()
    def apply_styles(self, MainWindow: QWidget) -> None:
        # 윈도우 스타일 설정
        MainWindow.setStyleSheet(application_style)

        # 버튼 스타일 설정
        self.load_button.setStyleSheet(button_style)
        self.save_button.setStyleSheet(button_style)
        self.exit_button.setStyleSheet(button_style)

        # 상태바 스타일 설정
        self.status_bar.setStyleSheet(statusbar_style)

    def select_excel_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        self.file_path = file_path
        if file_path:
            # self.status_bar.showMessage(f"선택된 파일: {file_path}", 5000)
            self.loaded_file_label.setText(f"선택된 파일: {file_path}")

    def select_save_location(self) -> None:
        folder_path = QFileDialog.getExistingDirectory(
            None,
            "저장할 위치 선택"
        )
        self.folder_path = folder_path
        if folder_path:
            # 저장 위치 레이블 업데이트
            self.save_location_label.setText(f"저장 위치: {folder_path}")

    def create_seperate_file(self) -> None:
        # SEPERATE 파일 생성 로직
        # 1. 원본 파일 복사 -> 
        #    원본 파일은 변경되지 않도록 복사본을 생성하여 사용
        self.status_bar.showMessage("원본 파일 복사 중...", 1000)
        new_file_name = f"{os.path.basename(self.file_path).split('.xlsx')[0]}_SEPERATE.xlsx"
        new_file_path = os.path.join(self.folder_path, new_file_name)
        shutil.copy2(self.file_path, new_file_path)
        self.status_bar.showMessage(f"원본 파일 복사 완료: {new_file_path}", 1000)
        # 2. "出席調査" 이름의 시트에서 "クラス名"을 기준으로 그룹화
        self.status_bar.showMessage("과목 시트 분리 중...", 5000)

        xls = pd.ExcelFile(new_file_path)

        if "出席調査" not in xls.sheet_names:
            self.status_bar.showMessage("出席調査 이름의 시트가 없습니다.", 2000)
            return
        
        df = pd.read_excel(new_file_path, sheet_name="出席調査")
        # 2-1. "クラス名"을 기준으로 그룹화
        grouped = df.groupby("クラス名")
        with pd.ExcelWriter(new_file_path, engine="openpyxl", mode='a') as writer:
            for name, group in grouped:
                # 2-2. 각 시트를 16개 칸으로 구분하여 저장(기존에서 마지막 열만 삭제하면 p(16)까지 있음)
                group.drop("@std.nagaokauniv.ac.jp", axis=1, inplace=True)
                group['출석시간']= pd.NA
                group.to_excel(writer, sheet_name=name, index=False)
        self.status_bar.showMessage(f"과목별 시트 분리 완료: {new_file_path}", 1000)
        # 3. 제공받은 원본에서 당담교수명, 학년, 학번, 이름, 이메일주소(担当教員名, 学年,学籍番号,氏名カナ)를 추출하고 중복 필드(학번기준(学籍番号))를 제거한 내용을 QR시트를추가하고 QR를 자동생성한다.
        # 3. 전체 시트에서 중복되지 않은 학생 정보를 기반으로 QR코드 생성성\
        self.status_bar.showMessage("QR코드 생성 중...", 5000)
        df_students = df[["担当教員名", "学年", "学籍番号", "氏名","カナ","学生メールアドレス"]].drop_duplicates(subset="学籍番号")
        
        # pandas로 엑셀 파일에 학생 정보를 먼저 저장
        with pd.ExcelWriter(new_file_path, engine="openpyxl", mode='a') as writer:
            df_students['QR']=pd.NA
            df_students.to_excel(writer, sheet_name="QR", index=False)
        self.status_bar.showMessage(f"QR코드 포함 시트 생성 완료: {new_file_path}", 1000)

        # openpyxl로 엑셀 파일 열기 (이미지 삽입을 위해)
        wb = openpyxl.load_workbook(new_file_path)
        qr_sheet = wb["QR"]

        # QR 코드 이미지를 엑셀 셀에 삽입
        row_num = 2  # 데이터는 첫 번째 행에 헤더가 있으므로 두 번째 행부터 시작
        teacher_name_image = None
        for index, row in df_students.iterrows():
            qr_code_data = {
                "担当教員名": row["担当教員名"], # 담당교수명
                "学年": row["学年"], # 학생
                "学籍番号": row["学籍番号"], # 학번
                "氏名": row["氏名"],  # 이름
                "カナ": row["カナ"],  # 이름 (カナ)
                "学生メールアドレス": row["学生メールアドレス"] # 학생 이메일
            }
            if teacher_name_image == None:
                teacher_name_image = self.create_text_to_image(qr_code_data["担当教員名"])
                teacher_name_image = teacher_name_image.convert("RGBA")
            # 직렬화(문자열로 변환)
            qr_code = json.dumps(qr_code_data)
            
            # qr 코드 생성
            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L)
            qr.add_data(qr_code)
            qr_image = qr.make_image(image_factory=StyledPilImage, module_drawer = RoundedModuleDrawer())
            # QR 코드의 중앙에 텍스트 이미지 삽입
            qr_image = qr_image.convert("RGBA")

            qr_image.paste(teacher_name_image, ((qr_image.size[0] - teacher_name_image.size[0]) // 2, (qr_image.size[1] - teacher_name_image.size[1]) // 2), teacher_name_image)

            # qr = qrcode.make(data=qr_code, box_size=9, border=4, version=1,error_correction=qrcode.constants.ERROR_CORRECT_L)

            # QR 코드를 이미지로 변환하여 BytesIO로 저장
            img_byte_arr = BytesIO()
            qr_image.save(img_byte_arr, "PNG")
            img_byte_arr.seek(0)

            # QR 이미지를 openpyxl Image 객체로 변환
            img = openpyxl.drawing.image.Image(img_byte_arr)
            # 이미지 크기 조정 (옵션, 필요시 크기를 조정)
            img.width = 50
            img.height = 50

            # 엑셀 셀에 이미지 삽입
            qr_sheet.add_image(img, f"G{row_num}")  # F열에 QR 코드 삽입
            row_num += 1
        wb.save(new_file_path)
        self.status_bar.showMessage(f"모든 작업 완료: {new_file_path}")
        # 3.2 이메일로 교수한테 전송 -> 이메일과 패스워드를 받기
          # 버튼 눌러 gui_email_window 창 띄우기
        # 4 수업별 QR 코드 체크기능은 이전에 구현된 코드를 사용. -> qr_reader
        pass

    def create_text_to_image(self,info):
        # getbbox()를 사용하여 텍스트의 크기를 계산합니다.
        bbox = font.getbbox(info)  # 텍스트 경계 상자의 크기
        width, height = bbox[2] - bbox[0], bbox[3] - bbox[1]  # (좌, 상, 우, 하)의 차이를 이용해 너비와 높이 계산

        # 텍스트를 담을 이미지를 생성합니다.
        img = Image.new("RGB", (width + 20, height + 10), color="white")
        draw = ImageDraw.Draw(img)

        # 텍스트를 이미지에 그립니다.
        draw.text((10, 0), info, font=font, fill="black")
        
        return img