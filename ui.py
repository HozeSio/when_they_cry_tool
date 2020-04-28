from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import converter

commands = {
    "export_text": [
        {
            'text': 'Update folder',
            'type': 'folder',
        }
    ],
    "replace_text": [
        {
            'text': 'Update folder',
            'type': 'folder',
        },
        {
            'text': 'Translation xlsx folder',
            'type': 'folder',
        },
        {
            'text': 'Actor xlsx file',
            'type': 'file',
        }
    ],
    "validate_folder": [

    ],
    "extract_text": [

    ],
    "combine_xlsx": [

    ],
    "insert_actor_column": [

    ],
    "remove_key_column": [

    ],
    "compare_line_count": [

    ],
    "insert_new_rows": [

    ],
    "get_actors": [

    ],
    "download": [

    ],
    "upload": [

    ],
    "insert_papago": [

    ],
    "unique_characters": [

    ],
    "find_old_format": [

    ],
    "export_text_onscript": [

    ],
    "export_text_steam": [

    ],
}


class Ui_Dialog(object):
    def __init__(self):
        self.selectedCommand = 'export_text'
        self.lineEditList = []

        self.app = QtWidgets.QApplication(sys.argv)
        self.dialog = QtWidgets.QDialog()
        self.setupUI()

    def exec(self):
        return self.app.exec()

    def _command_change(self, text):
        self.selectedCommand = text
        self.reDrawCommand(commands[text])

    def _open_file_dialog(self, lineEdit):
        result = str(QtWidgets.QFileDialog.getOpenFileName())
        lineEdit.setText(result)

    def _open_folder_dialog(self, lineEdit):
        result = str(QtWidgets.QFileDialog.getExistingDirectory())
        lineEdit.setText(result)
    
    def ex1(self):
        argv = ['ui', self.selectedCommand]
        argv.extend(lineEdit.text() for lineEdit in self.lineEditList)
        converter.convert(argv)
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle('Notice')
        msg.setText('작업 완료')
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        result = msg.exec_()
    
    def setupUI(self):
        self.dialog.setWindowTitle('When they cry tool')
        self.dialog.setObjectName("Dialog")
        self.dialog.setEnabled(True)
        self.dialog.resize(333, 300)
        frame = QtWidgets.QFrame(self.dialog)
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setFrameShadow(QtWidgets.QFrame.Raised)
        frame.setMinimumHeight(40)

        comboBox = QtWidgets.QComboBox(frame)
        comboBox.setEnabled(True)
        comboBox.setGeometry(QtCore.QRect(20, 10, 191, 22))
        #comboBox.setObjectName("commandBox")
        for command in commands.keys():
            comboBox.addItem(command)
        comboBox.currentTextChanged.connect(self._command_change)
        comboBox.setCurrentIndex(0)

        self.verticalLayout = QtWidgets.QVBoxLayout(self.dialog)
        self.verticalLayout.addWidget(frame)

        self.reDrawCommand(commands[self.selectedCommand])
        self.dialog.show()

    def reDrawCommand(self, command):
        self.lineEditList.clear()
        for idx in range(1, self.verticalLayout.count()):
            widget = self.verticalLayout.takeAt(1).widget()
            if widget:
                widget.deleteLater()
            #self.verticalLayout.takeAt(1).widget().deleteLater()

        frame = QtWidgets.QFrame(self.dialog)
        frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame.setFrameShadow(QtWidgets.QFrame.Raised)
        frame.setMinimumHeight(60 + 50 * len(command))

        for idx, argv in enumerate(command):
            label = QtWidgets.QLabel(frame)
            label.setGeometry(QtCore.QRect(20, 10 + 50 * idx, 150, 12))
            #self.label.setObjectName("label")
            label.setText(argv['text'])

            lineEdit = QtWidgets.QLineEdit(frame)
            lineEdit.setGeometry(QtCore.QRect(20, 30 + 50 * idx, 191, 20))
            self.lineEditList.append(lineEdit)
            #lineEdit.setObjectName("lineEdit")

            toolButton = QtWidgets.QToolButton(frame)
            toolButton.setGeometry(QtCore.QRect(230, 30 + 50 * idx, 61, 20))
            #self.toolButton.setObjectName("toolButton")
            toolButton.clicked.connect(lambda: self._open_file_dialog(lineEdit) if argv['type'] == 'file' else self._open_folder_dialog())
            toolButton.setText('경로')

        if command:
            pushButton = QtWidgets.QPushButton(frame)
            pushButton.setGeometry(QtCore.QRect(20, 10 + 50 * len(command), 280, 23))
            pushButton.setObjectName("pushButton")
            pushButton.clicked.connect(self.ex1)
            pushButton.setText('실행')

        self.verticalLayout.addWidget(frame)
        self.verticalLayout.addStretch()
        self.dialog.repaint()


def initializeUI():
    import sys
    ui = Ui_Dialog()
    sys.exit(ui.exec())


if __name__ == "__main__":
    initializeUI()