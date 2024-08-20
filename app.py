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

from qrReaderWidget import CameraViewer

# 설정 파일 경로
SETTINGS_FILE = "settings.json"

class PandasTableModel(QAbstractTableModel):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self._dataframe = dataframe

    def rowCount(self, parent=None):
        return len(self._dataframe)

    def columnCount(self, parent=None):
        return len(self._dataframe.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._dataframe.iloc[index.row(), index.column()]
            if pd.isna(value):
                return ""
            return str(value)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return self._dataframe.columns[section]
            else:
                return str(self._dataframe.index[section])
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Attendance Help")
        self.setWindowIcon(QIcon("./vision_logo.png"))
        self.resize(800, 600)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.label = QLabel("Select Excel File")
        self.layout.addWidget(self.label)

        self.selectFileBtn = QPushButton("Select Attendance File")
        self.selectFileBtn.clicked.connect(self.select_file_dialog)
        self.layout.addWidget(self.selectFileBtn)

        self.sheet_combo = QComboBox()
        self.sheet_combo.addItem("Select Sheet")  # Default item
        self.sheet_combo.currentIndexChanged.connect(self.load_sheet_data)
        self.layout.addWidget(self.sheet_combo)

        self.table_view = QTableView()
        self.table_view.setSizeAdjustPolicy(QTableView.SizeAdjustPolicy.AdjustToContents)
        self.table_view.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)

        self.layout.addWidget(self.table_view)

        self.saveResultsBtn = QPushButton("Save Results")
        self.saveResultsBtn.clicked.connect(self.save_results)
        self.layout.addWidget(self.saveResultsBtn)

        self.qr_generate_btn = QPushButton("QR Generate")
        self.qr_generate_btn.clicked.connect(self.qr_generate_start)
        self.qr_generate_btn.setEnabled(False)
        self.layout.addWidget(self.qr_generate_btn)

        self.qr_btn = QPushButton("QR Reader")
        self.qr_btn.clicked.connect(self.qr_window)
        self.qr_btn.setEnabled(False)
        self.layout.addWidget(self.qr_btn)

        self.file_path = None
        self.df = None
        self.sheet_names = []
        self.current_sheet= None

        # Load saved settings
        self.load_settings()
        self.update_qr_button_state()

    def qr_window(self):
        qr_reader_window = CameraViewer(self)
        qr_reader_window.qrProcessed.connect(self.handle_qr_processed)
        qr_reader_window.show()
        # pass

    def handle_qr_processed(self):
        self.load_sheet_data()
        
    def qr_generate_start(self):
        path = Path(self.file_path).parent
        qr_folder_path = os.path.join(path,"qr_folder",self.current_sheet)
        if not os.path.exists(qr_folder_path):
            os.makedirs(qr_folder_path)
        for row_idx, row in self.df.iloc[1:].iterrows():
            self.generateQR(row[1:3], qr_folder_path)
        QMessageBox.information(self,"Information","QR Generate Completed",QMessageBox.StandardButton.Ok)
        
    def generateQR(self,data:pd.Series, qr_folder_path):
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        file_name = f"{data.iloc[0]}_{data.iloc[1]}.png"

        # QR 생성값(data)은 번호+이름
        qr.add_data(','.join(data[0:2]))
        qr.make(fit=True)
        
        img = qr.make_image(fill_color="black", back_color="white")
        
        file_path = os.path.join(qr_folder_path, file_name)
        img.save(file_path)
        return file_path

    def select_file_dialog(self):
        file, _ = QFileDialog.getOpenFileName(self, 'Select File', "", "Excel Files (*.xlsx)")
        if file:
            file_without_extension = Path(file).stem
            self.label.setText(f"Selected File: {file_without_extension}")
            self.file_path = file
            self.save_settings()
            self.load_sheet_names()

    def load_sheet_names(self):
        if self.file_path:
            try:
                excel_file = pd.ExcelFile(self.file_path)
                self.sheet_names = excel_file.sheet_names
                for i in self.sheet_names:
                    if '出席調査' not in i:
                        self.add_date_column(i)
                self.sheet_combo.clear()
                self.sheet_combo.addItem("Select Sheet")
                self.sheet_combo.addItems(self.sheet_names)
                self.update_qr_button_state()
            except Exception as e:
                print(f"Error loading sheet names: {e}")

    def load_sheet_data(self):
        if self.file_path:
            self.current_sheet = self.sheet_combo.currentText()
            if self.current_sheet != "Select Sheet" and self.current_sheet != "":
                # 헤더 바로 밑에 날짜 추가
                # 날짜 넣을 곳 있으면 넘어가고
                # 없으면 추가
                    
                try:
                    self.df = pd.read_excel(self.file_path, sheet_name=self.current_sheet)
                    model = PandasTableModel(self.df)
                    self.table_view.setModel(model)
                    self.table_view.resizeColumnsToContents()
                    if "出席調査" in self.current_sheet:
                        self.update_qr_button_state(False)
                    else:
                        self.update_qr_button_state(True)
                except Exception as e:
                    print(f"Error loading sheet data: {e}")

    def add_date_column(self, sheet_name):
        # try:
            df = pd.read_excel(self.file_path, sheet_name = sheet_name)
            if df['学年'][0] != '年月日':
                # print(df['学年'][0])
                # 이곳에 빈열 추가
                empty = pd.DataFrame(index=range(0,1))
                empty.loc[0,'学年'] = '年月日'
                df = pd.concat([empty,df], axis=0)
                with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer,sheet_name=sheet_name, index=False)


        # except Exception as e:
        #     print(e)

    def update_qr_button_state(self, enabled=False):
        if enabled:
            self.qr_generate_btn.setEnabled(True)
            self.qr_generate_btn.setStyleSheet("background-color: #4CAF50; color: white;")  # Active style
            self.qr_btn.setEnabled(True)
            self.qr_btn.setStyleSheet("background-color: #4CAF50; color: white;")  # Active style
        else:
            self.qr_generate_btn.setEnabled(False)
            self.qr_generate_btn.setStyleSheet("background-color: lightgray; color: gray;")  # Disabled style
            self.qr_btn.setEnabled(False)
            self.qr_btn.setStyleSheet("background-color: lightgray; color: gray;")  # Disabled style


    def save_results(self):
        if self.file_path:
            results = []
            today_date = datetime.now().strftime("%Y-%m-%d")
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
                            count_x = row[attendance_columns].eq('x').sum()
                            results.append([sheet_name,grade, id,name, count_x, total_day])
                    except Exception as e:
                        print(f"Error processing sheet {sheet_name}: {e}")

                # 결과를 DataFrame으로 변환
            result_df = pd.DataFrame(results, columns=['クラス名', '学年', '学籍番号','氏名','欠席数','授業回数']) # 수업명, 학년, 학번, 이름, [가타가나 이름], 결석 수, [총 수업일 수]
            
            # 결과를 새로운 시트로 저장
            with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                # writer.book = book  # 이미 로드한 워크북을 writer에 설정
                result_df.to_excel(writer, sheet_name=f'出席調査_{today_date}', index=False)
            self.load_sheet_names()
        else:
            print("No file selected.")

    def save_settings(self):
        settings = {
            "file_path": self.file_path,
            "current_sheet": self.current_sheet
        }
        with open(SETTINGS_FILE, "w") as f:
            json.dump(settings, f)


    def load_settings(self):
        if Path(SETTINGS_FILE).exists():
            with open(SETTINGS_FILE, "r") as f:
                settings = json.load(f)
                self.file_path = settings.get("file_path")
                if self.file_path:
                    file_without_extension = Path(self.file_path).stem
                    self.label.setText(f"Selected File: {file_without_extension}")
                    self.load_sheet_names()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()
