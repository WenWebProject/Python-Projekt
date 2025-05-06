from PySide6.QtWidgets import QApplication, QMainWindow
from ui_design import Ui_Dialog  # Import the generated UI class
from pdf_merger import PDFMerger
from word_to_pdf import WordToPDF

class PDFTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)

        # Set the software's title
        self.setWindowTitle("PDFTool")

        # Ensure the first tab ("PDF Merge") is shown at startup
        self.ui.tabWidget.setCurrentIndex(0)

        # Initialize PDF merging feature
        self.pdf_merger = PDFMerger(self.ui, self)  # tab_1

        # Initialize Word-to-PDF feature
        self.word_to_pdf = WordToPDF(self.ui)   # tab_2



if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)

    window = PDFTool()
    window.show()

    sys.exit(app.exec())