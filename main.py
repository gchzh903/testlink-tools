import sys

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtWidgets import QPushButton, QHBoxLayout, QVBoxLayout, QWidget, QTextEdit, QFileDialog, QMessageBox

from excel_generator import *
from xml_generator import *


class MainUI(QWidget):

    def __init__(self):
        super(MainUI, self).__init__()
        self.status_edit = None
        self.file_edit = None
        self.init_ui()
        self.windowList = []

    def init_ui(self):
        layout = QVBoxLayout()
        dialog = QDialog()
        self.resize(400, 300)
        self.setWindowTitle('TestLink Tools')
        btn1 = QPushButton('XML --> EXCEL')
        btn2 = QPushButton('EXCEL --> XML')

        btn1.setFixedSize(200, 100)
        btn2.setFixedSize(200, 100)
        layout.setAlignment(Qt.AlignCenter)

        btn1.clicked.connect(self.show_xml_ui)
        btn2.clicked.connect(self.show_excel_ui)

        layout.addWidget(btn1)
        layout.addWidget(btn2)
        dialog.setLayout(layout)
        self.setLayout(layout)

    def show_xml_ui(self):

        window = QWidget()
        window.setWindowTitle('XML转换Excel工具 v1.0')
        window.setFixedSize(600, 350)
        window.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)

        box_layout = QHBoxLayout()
        file_edit = QTextEdit()
        file_edit.setFixedHeight(30)
        file_edit.setReadOnly(True)
        self.file_edit = file_edit

        open_button = QPushButton('打开测试用例XML')
        open_button.clicked.connect(self.on_open_xml_clicked)

        box_layout.addWidget(file_edit)
        box_layout.addWidget(open_button)

        status_edit = QTextEdit()
        status_edit.setEnabled(False)
        self.status_edit = status_edit
        status_edit.setStyleSheet('color:green')

        main_layout = QVBoxLayout()
        main_layout.addLayout(box_layout)
        main_layout.addWidget(status_edit)
        window.setLayout(main_layout)
        self.windowList.append(window)
        self.close()
        window.show()

    def show_excel_ui(self):

        window = QWidget()
        window.setWindowTitle('Excel转换XML工具 v1.0')
        window.setFixedSize(600, 350)
        window.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)

        hori_layout = QHBoxLayout()
        file_edit = QTextEdit()
        file_edit.setFixedHeight(30)
        file_edit.setReadOnly(True)
        self.file_edit = file_edit

        open_button = QPushButton('打开测试用例Excel')
        open_button.clicked.connect(self.on_open_excel_clicked)

        hori_layout.addWidget(file_edit)
        hori_layout.addWidget(open_button)

        status_edit = QTextEdit()
        status_edit.setEnabled(False)
        self.status_edit = status_edit
        status_edit.setStyleSheet('color:green')

        main_layout = QVBoxLayout()
        main_layout.addLayout(hori_layout)
        main_layout.addWidget(status_edit)
        window.setLayout(main_layout)
        self.windowList.append(window)
        self.close()
        window.show()

    def on_open_xml_clicked(self):
        filepath, filetype = QFileDialog.getOpenFileName(self, "选取文件", "./", "XML Files (*.xml)")
        if filepath == '':
            return
        self.file_edit.setText(filepath)
        outdir = os.path.dirname(filepath)

        try:
            trees = read_xml_and_build_cases(filepath)
            k = list(trees.keys())[0]
            v = list(trees.values())[0]
            excel_file = generate_excel(outdir, k, v)
            self.status_edit.insertPlainText('生成成功：%s \n' % excel_file)
        except Exception as e:
            QMessageBox.critical(self, '生成Excel出错', '请检查XML文件是否正确！')

    def on_open_excel_clicked(self):
        filepath, filetype = QFileDialog.getOpenFileName(self, "选取文件", "./", "Excel Files (*.xlsx)")
        if filepath == '':
            return
        self.file_edit.setText(filepath)
        outdir = os.path.dirname(filepath)

        try:
            trees = read_excel_and_build_trees(filepath)
            for k in trees:
                xml_file = generate_xml(outdir, k, trees[k])
                self.status_edit.insertPlainText('生成成功：' + xml_file)
        except Exception as e:
            QMessageBox.critical(self, '生成xml出错', '请检查excel文件中格式是否正确！')


def main():
    app = QApplication(sys.argv)
    ui = MainUI()
    ui.show()
    app.exec_()


if __name__ == '__main__':
    main()
