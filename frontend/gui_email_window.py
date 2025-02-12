import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from PyQt6.QtWidgets import QTableWidgetItem, QTableWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QDialog, QMessageBox, QComboBox
import pandas as pd
from io import BytesIO

class ExcelDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('엑셀 파일 불러오기')
        self.setGeometry(150, 150, 800, 600)

        layout = QVBoxLayout()

        # 엑셀 파일 열기 버튼
        self.open_button = QPushButton('엑셀 파일 선택', self)
        self.open_button.clicked.connect(self.load_excel)
        layout.addWidget(self.open_button)

        # 이메일, 비밀번호, 제목 입력 필드
        self.email_label = QLabel('이메일:', self)
        layout.addWidget(self.email_label)
        self.email_input = QLineEdit(self)
        layout.addWidget(self.email_input)

        self.password_label = QLabel('비밀번호:', self)
        layout.addWidget(self.password_label)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input)

        self.subject_label = QLabel('제목:', self)
        layout.addWidget(self.subject_label)
        self.subject_input = QLineEdit(self)
        layout.addWidget(self.subject_input)

        # 이메일 전송 버튼
        self.send_button = QPushButton('이메일 보내기', self)
        self.send_button.clicked.connect(self.send_email)
        layout.addWidget(self.send_button)

        # 전체 선택 버튼
        self.select_all_button = QPushButton('전체 선택', self)
        self.select_all_button.clicked.connect(self.select_all_cells)
        layout.addWidget(self.select_all_button)

        # QTableWidget을 사용하여 엑셀 데이터를 표시
        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.setLayout(layout)
        self.df = None  # 엑셀 데이터프레임을 저장할 변수

    def load_excel(self):
        # 파일 다이얼로그를 사용해 엑셀 파일 열기
        file, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx; *.xls)")
        if file:
            self.load_data(file)

    def load_data(self, file_path):
        # 엑셀 파일을 pandas DataFrame으로 불러오기
        xls = pd.ExcelFile(file_path)
        self.df = pd.read_excel(xls, "QR")  # "QR" 시트를 가져옴

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
        user_email = self.email_input.text()
        user_password = self.password_input.text()
        subject = self.subject_input.text()

        # SMTP 서버 설정 (예: Gmail)
        smtp_server = "smtp.gmail.com"
        smtp_port = 587

        try:
            # SMTP 서버에 연결하고 로그인
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()  # TLS 암호화 시작
            server.login(user_email, user_password)

            # 선택된 셀들에 대해서 이메일 전송
            selected_rows = set()  # 선택된 행 번호를 저장할 집합

            # 선택된 셀들을 확인
            selected_indexes = self.table_widget.selectedIndexes()
            for index in selected_indexes:
                row = index.row()  # 선택된 셀의 행 번호
                selected_rows.add(row)

            # 선택된 행에 대해 이메일 전송
            for row_idx in selected_rows:
                row = self.df.iloc[row_idx]

                # 이메일 주소와 QR 코드 이미지 파일 가져오기
                to_email = row['学生メールアドレス']  # 이메일 컬럼

                # QR 코드 이미지 가져오기 (QR 컬럼에서 이미지 데이터를 가져옴)
                qr_image = row['QR']  # 이미 QR 이미지 데이터가 컬럼에 저장되어 있음

                # 이미지가 bytes 형식으로 저장되어 있다고 가정하고 처리
                if isinstance(qr_image, BytesIO):
                    img_byte_arr = qr_image
                else:
                    # 만약 이미지가 다른 형식이라면, 이를 BytesIO로 변환
                    img_byte_arr = BytesIO(qr_image)

                # 이메일 메시지 구성
                message = MIMEMultipart()
                message['From'] = user_email
                message['To'] = to_email
                # 'Subject'는 제목임. 여기서는 사용자가 입력한 제목을 사용
                message['Subject'] = subject

                # 이메일 본문 작성
                body = "이 이메일은 PyQt6을 사용하여 전송되었습니다. 아래에 출석에 사용될 QR 이미지를 첨부합니다."
                message.attach(MIMEText(body, 'plain'))

                # QR 이미지를 첨부
                image = MIMEImage(img_byte_arr.read())
                image.add_header('Content-ID', '<qr_image>')
                message.attach(image)

                # 이메일 전송
                server.sendmail(user_email, to_email, message.as_string())

            # 서버 종료
            server.quit()

            print("이메일 전송 완료!")
            self.show_error("모든 이메일이 성공적으로 전송되었습니다!")

        except Exception as e:
            print(f"오류 발생: {e}")
            # 오류 메시지 표시
            self.show_error(f"이메일 전송 실패: {e}")

    def show_error(self, message):
        # 오류 메시지를 표시하는 함수
        error_msg = QMessageBox(self)
        error_msg.setIcon(QMessageBox.Icon.Critical)
        error_msg.setWindowTitle("오류")
        error_msg.setText(message)
        error_msg.exec()
