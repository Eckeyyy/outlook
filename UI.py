import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QTextEdit, QComboBox
from main import Outook
import pandas as pd


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()  # 修复super()调用，传递self参数

        self.initUI()

    def get_sheets(self, file):
        excel_file = pd.ExcelFile(file)
        sheet_names = excel_file.sheet_names
        return sheet_names

    def initUI(self):
        self.setWindowTitle('AutoOutlook')
        self.setGeometry(500, 100, 400, 600)

        layout = QVBoxLayout()
        label1 = QLabel('请选择配置')

        self.config_name_combo = QComboBox()
        self.config_name_combo.setFixedWidth(200)
        sheets = self.get_sheets('email.xlsx')
        new_sheets = [sheet for sheet in sheets if sheet != 'data']
        for s in new_sheets:
            self.config_name_combo.addItem(s)  # 添加下拉选项1
        confirm_button = QPushButton('确认')
        confirm_button.setFixedWidth(200)
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)

        layout.addWidget(label1)
        layout.addWidget(self.config_name_combo)
        layout.addWidget(confirm_button)
        layout.addWidget(self.output_text)

        confirm_button.clicked.connect(self.confirm_button_clicked)

        self.setLayout(layout)

    def confirm_button_clicked(self):
        excel_file = "email.xlsx"
        df = pd.read_excel(excel_file, sheet_name='data')
        config_name = self.config_name_combo.currentText()
        configs = pd.read_excel(excel_file, sheet_name=config_name, index_col=0)
        outlook = Outook(configs, df)
        outlook.send()
        result = config_name
        self.output_text.append(result)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
