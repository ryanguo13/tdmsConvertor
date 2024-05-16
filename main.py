import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel
from convertor import tdms_to_xls

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'TDMS to Excel Converter'
        self.left = 10
        self.top = 10
        self.width = 320
        self.height = 240
        self.tdms_path = "" 
        self.export_path = "" 
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        
        layout = QVBoxLayout()

        self.label = QLabel('Select file and folder to convert', self)
        layout.addWidget(self.label)

        # Button to open TDMS File
        self.buttonOpen = QPushButton('Open TDMS File', self)
        self.buttonOpen.clicked.connect(self.openFileNameDialog)
        layout.addWidget(self.buttonOpen)

        # Button to select Export Folder
        self.buttonSelectFolder = QPushButton('Select Export Folder', self)
        self.buttonSelectFolder.clicked.connect(self.folderDialog)
        layout.addWidget(self.buttonSelectFolder)

        # Button to convert
        self.buttonConvert = QPushButton('Convert', self)
        self.buttonConvert.clicked.connect(self.convert)
        layout.addWidget(self.buttonConvert)

        self.setLayout(layout)
    
    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"Open TDMS File", "","TDMS Files (*.tdms);;All Files (*)", options=options)
        if fileName:
            self.tdms_path = fileName
            self.label.setText(f"TDMS File Selected: {self.tdms_path}")

    def folderDialog(self):
        folder = str(QFileDialog.getExistingDirectory(self, "Select Export Directory"))
        if folder:
            self.export_path = folder
            self.label.setText(f"Selected Export Folder: {self.export_path}")

    def convert(self):
        if self.tdms_path and self.export_path:
            xls_path = f"{self.export_path}/OutputFile.xlsx"  # Define the output filename
            tdms_to_xls(self.tdms_path, xls_path)
            self.label.setText(f"Conversion Complete: {xls_path}")
        else:
            self.label.setText("Please select a TDMS file and export folder.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())