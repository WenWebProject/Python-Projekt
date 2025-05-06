import os
from PySide6.QtWidgets import QFileDialog, QMessageBox, QWidget
from PySide6.QtCore import QTimer
from PyPDF2 import PdfMerger  # Library for merging PDFs
import shutil



class PDFMerger:
    def __init__(self, ui, parent_widget: QWidget):
        self.ui = ui
        self.parent_widget = parent_widget # Pass the main window or relevant parent QWidget
       
        # Storage for uploaded files
        self.saved_files = [None, None]  # Track file paths for two files

        # Initialize attributes to track upload states
        self.uploaded_files = [None, None]  # Keep track of uploaded files
        self.merged_pdf_path = None
        self.current_progress = 0

        # Connect buttons to methods
        self.ui.pushButton_30.clicked.connect(self.select_file)  # Select file
        self.ui.pushButton_31.clicked.connect(self.upload_file)  # Upload file

        # For the first file
        self.ui.pushButton.clicked.connect(lambda: self.save_file(0))  # Save button for file 1
        self.ui.pushButton_5.clicked.connect(lambda: self.discard_file(0))  # Discard button for file 1

        # For the second file
        self.ui.pushButton_2.clicked.connect(lambda: self.save_file(1))  # Save button for file 2
        self.ui.pushButton_6.clicked.connect(lambda: self.discard_file(1))  # Discard button for file 2

        self.ui.pushButton_4.clicked.connect(self.merge_pdfs)  # Merge PDFs
        self.ui.pushButton_3.clicked.connect(self.download_merged_pdf)  # Download merged PDF
        self.ui.pushButton_32.clicked.connect(self.delete_merged_pdf)  # Delete merged PDF
       
    def select_file(self):
        # Let user select a PDF file
        file_path, _ = QFileDialog.getOpenFileName(self.parent_widget, "Select PDF File", "", "PDF Files (*.pdf)")
        if file_path:
            # Display selected file path in the line edit
            self.ui.lineEditFilePath_6.setText(file_path)

    def upload_file(self):
        file_path = self.ui.lineEditFilePath_6.text()
        if not file_path:
            QMessageBox.warning(self, "Warning", "Please select a file to upload!")
            return

        # Simulate upload with a progress bar
        self.current_progress = 0
        self.timer = QTimer(self.parent_widget)
        self.timer.timeout.connect(self.update_progress)
        self.timer.start(100)  # Update every 100ms

        # Assign the upload to the first or second slot
        if self.uploaded_files[0] is None:
            self.target_index = 0
            self.ui.progressBar_26.setValue(0)
        elif self.uploaded_files[1] is None:
            self.target_index = 1
            self.ui.progressBar_27.setValue(0)
        else:
            QMessageBox.warning(self, "Warning", "You can only upload two files!")
            return

        # Store the file path temporarily
        self.uploaded_files[self.target_index] = file_path

    def update_progress(self):
        self.current_progress += 10  # Simulate 10% progress increment
        if self.target_index == 0:
            self.ui.progressBar_26.setValue(self.current_progress)
        elif self.target_index == 1:
            self.ui.progressBar_27.setValue(self.current_progress)

        
        if self.current_progress >= 100:
            self.timer.stop()
            file_name = os.path.basename(self.uploaded_files[self.target_index])
            if self.target_index == 0:
                self.ui.textBrowser.setText(file_name)  # Display file name
            elif self.target_index == 1:
                self.ui.textBrowser_2.setText(file_name)   # Display file name
           # QMessageBox.information(self, "Success", f"File '{file_name}' uploaded successfully!")

    # Save the selected file and update the LCD display and saved_files list
    def save_file(self, file_index):
      file_path = self.ui.lineEditFilePath_6.text()
      if file_path:  # Ensure a file path is provided
        if file_index == 0:
            self.ui.lcdNumber.display("1")  # File 1 saved
            self.saved_files[0] = file_path  # Save file path for the first file
            
        elif file_index == 1:
            self.ui.lcdNumber_2.display("1")  # File 2 saved
            self.saved_files[1] = file_path  # Save file path for the 2nd file

    #  Discard the file and reset the display and saving path for the specified file index.
    def discard_file(self, file_index):
            if file_index == 0:
               self.ui.textBrowser.clear()         # Clear the display of the first file name
               self.ui.lineEditFilePath_6.clear()  # clear the file path input
               self.ui.lcdNumber.display(0)        # Set the QLCDNumber to 0 for the first file
               self.ui.progressBar_26.setValue(0)  # Reset the progress bar
               self.uploaded_files[0] = None       # clear the 1st file reference

            elif file_index == 1:
               self.ui.textBrowser_2.clear()      # Clear the display of the 2nd file name
               self.ui.lineEditFilePath_6.clear()  # clear the file path input
               self.ui.lcdNumber_2.display(0)     # Set the QLCDNumber to 0 for the 2nd file
               self.ui.progressBar_27.setValue(0) # Reset the progress bar
               self.uploaded_files[1] = None      # clear the 2nd file reference
            
    def merge_pdfs(self):
        if None in self.uploaded_files:
            QMessageBox.warning(self, "Warning", "Please upload two files before merging!")
            return

        merger = PdfMerger()
        for file in self.uploaded_files:
            merger.append(file)

        self.merged_pdf_path = os.path.join(os.getcwd(), "merged_PDF.pdf")
        merger.write(self.merged_pdf_path)
        merger.close()

        self.ui.progressBar_5.setValue(100)
        self.ui.textBrowser_3.setText("merged_PDF.pdf")
        self.ui.lcdNumber_3.display("1")
        QMessageBox.information(self.parent_widget, "Success", "PDF files merged successfully!")

    def download_merged_pdf(self):
        if not self.merged_pdf_path or not os.path.exists(self.merged_pdf_path):
            QMessageBox.warning(self, "Warning", "No merged PDF to download!")
            return

        download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        download_path = os.path.join(download_folder, "merged_PDF.pdf")

        self.current_progress = 0
        self.timer = QTimer(self.parent_widget)   # Use parent_widget if available
        self.timer.timeout.connect(lambda: self.update_download_progress(download_path))
        self.timer.start(100)

    def update_download_progress(self, download_path):
        self.current_progress += 10
        self.ui.progressBar_6.setValue(self.current_progress)

        if self.current_progress >= 100:
            self.timer.stop()
            shutil.copy(self.merged_pdf_path, download_path)
            QMessageBox.information(None, "Success", f"Merged PDF downloaded to {download_path}")

    def delete_merged_pdf(self):
        if not self.merged_pdf_path or not os.path.exists(self.merged_pdf_path):
            QMessageBox.warning(self, "Warning", "No merged PDF to delete!")
            return

        os.remove(self.merged_pdf_path)
        self.merged_pdf_path = None

        self.ui.textBrowser_3.clear()
        self.ui.lcdNumber_3.display("0")
        self.ui.progressBar_5.setValue(0)
        QMessageBox.information(None, "Deleted", "Merged PDF deleted successfully!")


