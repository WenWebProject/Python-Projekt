from PySide6.QtWidgets import QApplication, QLabel

# Create an application
app = QApplication([])

# Create a label
label = QLabel("Hello, PySide6!")
label.show()

# Run the application
app.exec()