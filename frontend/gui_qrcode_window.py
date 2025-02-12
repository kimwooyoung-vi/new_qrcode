import sys
import json
import qrcode
import os
from PyQt6.QtWidgets import (
    QMainWindow,
    QApplication,
    QWidget,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QComboBox,
    QTableView,
    QMessageBox,
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QAbstractTableModel, Qt
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook

from core.qr_reader.qrReaderWidget import CameraViewer

# 설정 파일 경로
SETTINGS_FILE = "settings.json"

class PandasTableModel(QAbstractTableModel):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self._dataframe = dataframe
    # QAbstractTableModel의 메서드를 오버라이드
    def rowCount(self, parent=None):
        return len(self._dataframe)
    # QAbstractTableModel의 메서드를 오버라이드
    def columnCount(self, parent=None):
        return len(self._dataframe.columns)
    # QAbstractTableModel의 메서드를 오버라이드
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._dataframe.iloc[index.row(), index.column()]
            if pd.isna(value):
                return ""
            return str(value)
        return None
    # QAbstractTableModel의 메서드를 오버라이드
    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self._dataframe.columns[section]
            else:
                return str(self._dataframe.index[section])
        return None

class MainWindow(QMainWindow):
    # 메인 윈도우 클래스
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Attendance Help") # 창 제목
        self.setWindowIcon(QIcon("./vision_logo.png")) # 창 아이콘 설정
        self.resize(800, 600) # 창 크기 설정
        self.central_widget = QWidget(self) # 중앙 위젯 설정
        self.setCentralWidget(self.central_widget) # 중앙 위젯 지정

        # 레이아웃 설정
        self.layout = QVBoxLayout(self.central_widget)

        self.label = QLabel("Select Excel File") # 파일선택 문구
        self.layout.addWidget(self.label)

        # 파일 선택 버튼 설정
        self.selectFileBtn = QPushButton("Select Attendance File")
        self.selectFileBtn.clicked.connect(self.select_file_dialog) # 클릭 시 엑셀 파일 선택 창 표시
        self.layout.addWidget(self.selectFileBtn)

        # 엑셀 시트 선택 콤보박스
        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("Select Sheet")  # 기본 항목
        self.sheet_combo.currentIndexChanged.connect(self.load_sheet_data) # 선택한 엑셀파일의 시트목록 로드
        self.layout.addWidget(self.sheet_combo)

        # 테이블 설정
        self.table_view = QTableView()
        self.table_view.setSizeAdjustPolicy(QTableView.SizeAdjustPolicy.AdjustToContents)
        self.table_view.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.layout.addWidget(self.table_view)

        # 출석 통계 결과 저장 버튼 설정 -> 누르면 몇차시 수업을 출석했는지 표기됨.
        # self.saveResultsBtn = QPushButton("Save Results")
        # self.saveResultsBtn.clicked.connect(self.save_results)
        # self.layout.addWidget(self.saveResultsBtn)

        # # QR 코드 생성 버튼 설정
        # self.qr_generate_btn = QPushButton("QR Generate")
        # self.qr_generate_btn.clicked.connect(self.qr_generate_start)
        # self.qr_generate_btn.setEnabled(False) # 엑셀 시트 선택이전까지 비활성화
        # self.layout.addWidget(self.qr_generate_btn)

        # QR 코드 읽기 버튼 설정
        self.qr_btn = QPushButton("QR Reader")
        self.qr_btn.clicked.connect(self.qr_window)
        self.qr_btn.setEnabled(False) # 엑셀 시트 선택이전까지 비활성화
        self.layout.addWidget(self.qr_btn)

        # 초기 값 설정
        self.file_path = None
        self.df = None
        self.sheet_names = []
        self.current_sheet= None

        # 저장된 설정 로드 ( 마지막 선택한 엑셀 파일경로, 시트명 )
        # self.load_settings()
        self.update_qr_button_state()

    def qr_window(self):
        # QR 읽기 화면(창)을 띄움. -> qrReaderWidget.py에 정의됨.
        qr_reader_window = CameraViewer(self)
        qr_reader_window.qrProcessed.connect(self.handle_qr_processed)
        qr_reader_window.show()
        # pass

    def handle_qr_processed(self):
        # QR 코드가 읽힐때마다 시트를 새로고침
        self.load_sheet_data()
        
    def qr_generate_start(self):
        # QR 코드 생성시작
        path = Path(self.file_path).parent
        qr_folder_path = os.path.join(path,"qr_folder",self.current_sheet)
        if not os.path.exists(qr_folder_path):
            os.makedirs(qr_folder_path) # 이미지를 저장할 폴더 생성
        for row_idx, row in self.df.iloc[1:].iterrows():
            self.generateQR(row[1:3], qr_folder_path) # 차례대로 읽어 QR 코드 생성
        QMessageBox.information(self,"Information","QR Generate Completed",QMessageBox.StandardButton.Ok)
        
    # def generateQR(self,data:pd.Series, qr_folder_path):
    #     # qrcode 라이브러리를 사용하는 위치.
    #     qr = qrcode.QRCode(
    #         version=1,
    #         error_correction=qrcode.constants.ERROR_CORRECT_L,
    #         box_size=10,
    #         border=4,
    #     )
    #     file_name = f"{data.iloc[0]}_{data.iloc[1]}.png"

    #     # QR 생성값(data)은 번호+이름
    #     qr.add_data(','.join(data[0:2]))
    #     qr.make(fit=True)
        
    #     img = qr.make_image(fill_color="black", back_color="white")
        
    #     file_path = os.path.join(qr_folder_path, file_name)
    #     img.save(file_path)
    #     return file_path

    def select_file_dialog(self):
        # 엑셀 파일 선택 창 호출
        file, _ = QFileDialog.getOpenFileName(self, 'Select File', "", "Excel Files (*.xlsx)")
        if file:
            file_without_extension = Path(file).stem
            self.label.setText(f"Selected File: {file_without_extension}")
            self.file_path = file
            # self.save_settings()
            self.load_sheet_names()

    def load_sheet_names(self):
        # 엑셀 파일에서 시트 목록을 불러옴 
        if self.file_path:
            try:
                excel_file = pd.ExcelFile(self.file_path)
                self.sheet_names = excel_file.sheet_names
                # for i in self.sheet_names:
                #     if '出席調査' not in i:
                #         self.add_date_column(i)
                self.sheet_combo.clear()
                self.sheet_combo.addItem("Select Sheet")
                self.sheet_combo.addItems(self.sheet_names)
                self.update_qr_button_state()
            except Exception as e:
                print(f"Error loading sheet names: {e}")

    def load_sheet_data(self):
        # 선택된 시트에서 데이터를 읽어옴
        if self.file_path:
            self.current_sheet = self.sheet_combo.currentText()
            if self.current_sheet != "Select Sheet" and self.current_sheet != "":
                # 헤더 바로 밑에 날짜 추가
                # 날짜 넣을 곳 있으면 넘어가고
                # 없으면 추가
                    
                try:
                    self.df = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
                    # 읽어온 pandas 데이터를 QT 형식의 테이블로 변환
                    model = PandasTableModel(self.df)
                    self.table_view.setModel(model) # 테이블 뷰에 QAbstractTableModel 구현체 전달
                    self.table_view.resizeColumnsToContents() # 열 너비 자동 조절
                    if "出席調査" in self.current_sheet:
                        self.update_qr_button_state(False) # 출석 조사가 포함된 시트명은 QR 버튼 비활성화
                    else:
                        self.update_qr_button_state(True) # 
                except Exception as e:
                    print(f"Error loading sheet data: {e}")

    # def add_date_column(self, sheet_name):
    #     # 출석부 첫 행에 요일 정보 추가
    #     df = pd.read_excel(self.file_path, sheet_name = sheet_name)
        
    #     if df['学年'][0] != '年月日':
    #         # print(df['学年'][0])
    #         # 이곳에 빈열 추가
    #         empty = pd.DataFrame(index=range(0,1))
    #         empty.loc[0,'学年'] = '年月日'
    #         df = pd.concat([empty,df], axis=0)
    #         with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #             df.to_excel(writer,sheet_name=sheet_name, index=False)

    def update_qr_button_state(self, enabled=False):
        # QR 버튼 상태 업데이트(활성/비활성) - 出席調査 가 들어가있으면 비활성화 TODO[출석부 형태의 규칙을 따르면 활성화되도록 변경]
        if enabled:
            # self.qr_generate_btn.setEnabled(True)
            # self.qr_generate_btn.setStyleSheet("background-color: #4CAF50; color: white;")  # Active style
            self.qr_btn.setEnabled(True)
            self.qr_btn.setStyleSheet("background-color: #4CAF50; color: white;")  # Active style
        else:
            # self.qr_generate_btn.setEnabled(False)
            # self.qr_generate_btn.setStyleSheet("background-color: lightgray; color: gray;")  # Disabled style
            self.qr_btn.setEnabled(False)
            self.qr_btn.setStyleSheet("background-color: lightgray; color: gray;")  # Disabled style


    def save_results(self):
        # 결과를 저장하는 메서드
        if self.file_path:
            results = []
            # today_date = datetime.now().strftime("%Y-%m-%d")
            for sheet_name in self.sheet_names:
                if "出席調査" not in sheet_name:
                    try:
                        # 학년, 학번, 이름 
                        df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                        
                        student_grade = df.columns[0]
                        student_id = df.columns[1]
                        student_name = df.columns[2]
                        attendance_columns = [col for col in df.columns if "回目" in col]
                        total_day = len(attendance_columns)
                        for _, row in df.iloc[1:].iterrows():
                            grade = row[student_grade]
                            id = row[student_id]
                            name = row[student_name]
                            count_x = row[attendance_columns].eq('x').sum() # 결석 수 계산
                            results.append([sheet_name,grade, id,name, count_x, total_day])
                    except Exception as e:
                        print(f"Error processing sheet {sheet_name}: {e}")

            #     # 결과를 DataFrame으로 변환
            # result_df = pd.DataFrame(results, columns=['クラス名', '学年', '学籍番号','氏名','欠席数','授業回数']) # 수업명, 학년, 학번, 이름, [가타가나 이름], 결석 수, [총 수업일 수]
            
            # # 결과를 새로운 시트로 저장
            # with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            #     # writer.book = book  # 이미 로드한 워크북을 writer에 설정
            #     result_df.to_excel(writer, sheet_name=f'出席調査_{today_date}', index=False)
            self.load_sheet_names() # 시트 목록 갱신 -> 콤보박스에 出席調査_{today_date} 추가
        else:
            print("No file selected.")

    # 주석 처리하였음.(미사용)
    def save_settings(self):
        # 설정 저장하는 메서드 -> 마지막 선택한 파일 경로, 시트명 저장
        settings = {
            "file_path": self.file_path,
            "current_sheet": self.current_sheet
        }
        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f)

    # (미사용)
    def load_settings(self):
        # 설정을 불러오는 메서드
        if Path(SETTINGS_FILE).exists():
            with open(SETTINGS_FILE, "r") as f:
                settings = json.load(f)
                self.file_path = settings.get("file_path")
                if self.file_path:
                    file_without_extension = Path(self.file_path).stem
                    self.label.setText(f"Selected File: {file_without_extension}")
                    self.load_sheet_names()

# 테스트할때만 사용하는 부분
# if __name__ == "__main__":
#     app = QApplication(sys.argv) # QT 객체 생성
#     window = MainWindow() # 메인 윈도우 객체 생성
#     window.show() # 윈도우 표시
#     app.exec() # 실행.
