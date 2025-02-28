import openpyxl
from openpyxl.drawing.image import Image as openpyxl_Image
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from PyQt6.QtWidgets import QTableWidgetItem, QTableWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QDialog, QMessageBox, QComboBox
import pandas as pd
from io import BytesIO
from PIL import Image

import ssl
from smtplib import SMTP_SSL

class ExcelDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('エクセルファイル読み込み')
        self.setGeometry(150, 150, 800, 600)

        layout = QVBoxLayout()

        # 엑셀 파일 열기 버튼
        self.open_button = QPushButton('セクセルファイル選択', self)
        self.open_button.clicked.connect(self.load_excel)
        layout.addWidget(self.open_button)

        # 이메일, 비밀번호, 제목 입력 필드
        self.email_label = QLabel('メールアドレス：', self)
        layout.addWidget(self.email_label)
        self.email_input = QLineEdit(self)
        layout.addWidget(self.email_input)

        self.password_label = QLabel('パスワード：', self)
        layout.addWidget(self.password_label)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input)

        self.subject_label = QLabel('メール件名：', self)
        layout.addWidget(self.subject_label)
        self.subject_input = QLineEdit(self)
        layout.addWidget(self.subject_input)

        self.body_label = QLabel('メール本文：', self)
        layout.addWidget(self.body_label)
        self.body_input = QLineEdit(self)
        layout.addWidget(self.body_input)

        # 이메일 전송 버튼
        self.send_button = QPushButton('メール送信', self)
        self.send_button.clicked.connect(self.send_email)
        layout.addWidget(self.send_button)

        # 전체 선택 버튼
        self.select_all_button = QPushButton('全選択', self)
        self.select_all_button.clicked.connect(self.select_all_cells)
        layout.addWidget(self.select_all_button)

        # QTableWidget을 사용하여 엑셀 데이터를 표시
        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.setLayout(layout)
        self.df = None  # 엑셀 데이터프레임을 저장할 변수

    def load_excel(self):
        # 파일 다이얼로그를 사용해 엑셀 파일 열기
        self.file, _ = QFileDialog.getOpenFileName(self, "エクセルファイル選択", "", "Excel Files (*.xlsx; *.xls)")
        if self.file:
            self.load_data(self.file)

    def load_data(self, file_path):
        # 엑셀 파일을 pandas DataFrame으로 불러오기
        xls = pd.ExcelFile(file_path)
        try:
            self.df = pd.read_excel(xls, "QR")  # "QR" 시트를 가져옴
        except:
            self.show_error("選択したファイルにQRシートがありません。 SEPERATEで生成されたファイルを選択してください。")
            return
        # 데이터 로딩 후 테이블에 표시
        self.display_table()

    def display_table(self):
        # QTableWidget을 사용하여 데이터를 표시
        self.table_widget.setRowCount(len(self.df))
        self.table_widget.setColumnCount(len(self.df.columns))
        self.table_widget.setHorizontalHeaderLabels(self.df.columns)

        # 데이터를 테이블에 채우기
        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                self.table_widget.setItem(row, col, QTableWidgetItem(str(self.df.iloc[row, col])))

    def select_all_cells(self):
        # 전체 셀 선택
        self.table_widget.selectAll()

    def send_email(self):
        # 메일 서버 연결에 사용되는 값
        user_email = self.email_input.text()
        user_password = self.password_input.text()
        # 메일 제목, 내용
        subject = self.subject_input.text()
        body = self.body_input.text()

        # SMTP 서버 설정 (예: Gmail)
        smtp_server = "smtp.gmail.com"
        smtp_port = 465
        # smtp_port = 587

        try:

            # 선택된 셀들에 대해서 이메일 전송
            selected_rows = set()  # 선택된 행 번호를 저장할 집합

            # 선택된 셀들을 확인
            selected_indexes = self.table_widget.selectedIndexes()
            for index in selected_indexes:
                row = index.row()  # 선택된 셀의 행 번호
                selected_rows.add(row)

            # 이곳에서 선택한 인원을 검증하는 기능 추가
            # 몇명이 선택되었는지, 선택된 인원이 없는지, 선택된 인원의 이름을 확인하는 기능 실행
            cancel_send = self.check_selected_rows(selected_rows)

            if cancel_send:
                return
            
            # SMTP 서버에 연결하고 로그인
            # server = smtplib.SMTP(smtp_server, smtp_port)
            # server.starttls()
            context = ssl.create_default_context()
            server = SMTP_SSL(smtp_server, smtp_port, context=context)
            
            server.login(user_email, user_password)
                        
            # 선택된 행에 대해 이메일 전송
            for row_idx in selected_rows:
                row = self.df.iloc[row_idx]

                # 이메일 주소와 QR 코드 이미지 파일 가져오기
                to_email = row['学生メールアドレス']  # 이메일 컬럼

                # QR 코드 이미지 가져오기 (QR 컬럼에서 이미지 데이터를 가져옴)
                # def extract_image_from_excel(row_idx):
                    # for image in sheet._images:  # 시트에서 이미지들을 순회
                        # if image.anchor._from.row == row_idx:
                    # ... 위의 방법을 사용할때 인덱스가 1부터 시작하기에 row_idx에 1을 추가해서 탐색
                img_byte_arr = self.extract_image_from_excel(row_idx + 1)
                if img_byte_arr is None:
                    print("正しいイメージ形式ではありません。エクセルファイルと一緒にお問合せしてください。")
                    return
                

                # 이메일 메시지 구성
                message = MIMEMultipart()
                message['From'] = user_email
                message['To'] = to_email
                # 'Subject'는 제목임. 여기서는 사용자가 입력한 제목을 사용
                message['Subject'] = subject

                # 이메일 본문 작성
                
                message.attach(MIMEText(body, 'plain'))

                # QR 이미지를 첨부
                image = MIMEImage(img_byte_arr.read(), _subtype="png")
                image.add_header('Content-Disposition', 'attachment', filename='qr_image.png') # 첨부 파일의 이름을 설정
                message.attach(image)

                # 이메일 전송
                server.send_message(message)
                print("From: ", user_email, "To: ", to_email, "メール送信完了")

            # 서버 종료
            server.quit()

            print("メール送信完了！")
            self.show_error("メールが正しく送信できました！")

        except Exception as e:
            print(f"送信エラー： {e}")
            # 오류 메시지 표시
            self.show_error(f"メール送信失敗： {e}")

    def extract_image_from_excel(self, row_idx):
        wb = openpyxl.load_workbook(self.file)  # with 문을 사용하지 않고 파일 열기
        try:
            sheet = wb['QR']  # QR 시트 가져오기

            for image in sheet._images:  # 시트에서 이미지들을 순회
                if image.anchor._from.row == row_idx:  # 선택된 행에 삽입된 이미지인지 확인
                    img_byte_arr = BytesIO()
                    pil_img = Image.open(image.ref)  # openpyxl.Image 객체의 ref 속성을 통해 이미지를 읽음
                    pil_img.save(img_byte_arr, format='PNG')  # 이미지를 PNG 형식으로 바이트 스트림으로 저장
                    img_byte_arr.seek(0)  # 바이트 스트림의 시작으로 이동
                    return img_byte_arr
        finally:
            wb.close()  # workbook을 작업한 후 명시적으로 닫아줍니다.
        
        return None           


    def show_error(self, message):
        # 오류 메시지를 표시하는 함수
        error_msg = QMessageBox(self)
        error_msg.setIcon(QMessageBox.Icon.Critical)
        error_msg.setWindowTitle("エラー")
        error_msg.setText(message)
        error_msg.exec()

    def check_selected_rows(self, selected_rows:set):
        # 선택된 인원 이름 확인
        names = [self.df.iloc[row]['氏名'] for row in selected_rows]  # '학생이름' 컬럼에서 이름을 가져옴
        selected_names = "\n".join(names)
        
        # 확인 메시지 창 표시
        reply = QMessageBox.question(self, "選択された学生確認", 
                                     f"選択された学生:\n{selected_names}\n\nへメールを送信しますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.No:
            # 이메일 전송 취소
            return True  # 전송 취소
        return False  # 전송 진행