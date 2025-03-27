import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QTextEdit
)
from convertion_tools import utils


class FileSelectorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setGeometry(300, 300, 500, 500)
        self.setWindowTitle("Word Files Consolidator App")
        
        # Select Folder
        self.selectFolderButton = QPushButton('Select Folder', self)
        self.selectFolderButton.clicked.connect(self.showFolderDialog)
        
        # Textbox with selected Folder
        self.txtTargetFolder = QTextEdit(self)
        
        # Select Files
        self.selectButton = QPushButton('Select Files', self)
        self.selectButton.clicked.connect(self.showFileDialog)
        
        # Display selected files
        self.textEdit = QTextEdit(self)
        self.textEdit.setReadOnly(True)
        
        # Button to proccess files
        self.btnProccessFiles = QPushButton('Convert Files', self)
        self.btnProccessFiles.clicked.connect(self.processFiles)
        
        
        layout = QVBoxLayout()
        layout.addWidget(self.selectFolderButton)
        layout.addWidget(self.txtTargetFolder)
        layout.addWidget(self.selectButton)
        layout.addWidget(self.textEdit)
        layout.addWidget(self.btnProccessFiles)
        
        self.setLayout(layout)
        self.show()
        
    def showFileDialog(self):
        file_dialog = QFileDialog()
        file_paths, _ = file_dialog.getOpenFileNames(self, "Select Files", "", "")
        
        if file_paths:
            self.textEdit.clear()
            for file_path in file_paths:
                self.textEdit.append(file_path)
                
    def showFolderDialog(self):
        folder_dialog = QFileDialog()
        folder_path = folder_dialog.getExistingDirectory(self, "Select Folder")
        self.txtTargetFolder.setText(folder_path)
        
    def processFiles(self):
        target_folder = self.txtTargetFolder.toPlainText()
        files_selected_text = self.textEdit.toPlainText()
        # Convert to list
        files_selected_list = files_selected_text.splitlines()
        # Convert "docx" files to "doc"
        utils.process_files(files_selected_list, target_folder)
        # Read new "doc" files
        doc_files = os.listdir(target_folder)
        doc_files_fixed = []
        for doc in doc_files:
            doc_files_fixed.append(os.path.join(target_folder, doc))
        # Convert each "doc" file to "pdf"
        utils.convert_files_to_pdf(doc_files_fixed, target_folder)
        # Merge PDF files into a single file
        utils.merge_pdf_files(target_folder)
        # Convert the pdf to "docx" file
        utils.convert_pdf_to_word(target_folder)
        
        print("\nConversion of files Completed....Have a nice day!!!!")
        
        
                
if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = FileSelectorGUI()
    sys.exit(app.exec_())