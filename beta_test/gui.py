import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QStackedWidget, QMessageBox, QFileDialog, QLabel
import test as cr

class ReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Report Generator')
        self.file_path = 0

        # Create widgets
        self.import_button = QPushButton('Import Data OR Drag and Drop', self)
        self.import_button.clicked.connect(self.import_data)
        self.import_button.setMinimumHeight(60)  # Set the minimum height to 60 (adjust as needed)
        self.import_button.setMinimumWidth(200)

        # Customize the appearance of the import button
        font = self.import_button.font()
        font.setPointSize(12)  # Set the font size to 12
        font.setBold(True)    # Set the font to bold
        self.import_button.setFont(font)

        # Create a table widget for data preview
        self.data_preview_table = QTableWidget(self)


        self.header_preview_label = QLabel('Header Preview:', self)
        self.header_preview_label.setFont(font)

        self.header_preview_table = QTableWidget(self)
        self.header_preview_table.setMaximumHeight(80)  # Set maximum height (adjust as needed)


        self.default_button = QPushButton('Process', self)
        self.default_button.clicked.connect(lambda: cr.main(self.data, self.supp_df))
        self.default_button.clicked.connect(self.show_done_message)
        self.default_button.setMinimumHeight(60)  # Set the minimum height to 60 (adjust as needed)
        self.default_button.setMinimumWidth(200)
        self.default_button.setFont(font)
        self.default_button.setEnabled(False)

        self.setAcceptDrops(True)  # Enable drag and drop

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.import_button)
        main_layout.addWidget(self.data_preview_table)
        main_layout.addWidget(self.header_preview_label)
        main_layout.addWidget(self.header_preview_table)
        main_layout.addWidget(self.default_button)  # Align the button at the bottom

        self.main_widget = QWidget()
        self.main_widget.setLayout(main_layout)

        # Stacked widget to switch between layouts (no change here)
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.addWidget(self.main_widget)
        # Add other layouts as needed

        self.setCentralWidget(self.stacked_widget)


    def import_data(self):
        # Open a file dialog to allow the user to select an Excel file for import
        file_dialog = QFileDialog()

        # If no file has been selected previously, set the file_path to the selected file's path
        if self.file_path == 0:
            self.file_path, _ = file_dialog.getOpenFileName(self, 'Open Excel File', '', 'Excel Files (*.xlsx);;All Files (*)')

        if self.file_path:
            try:
                # Read the Excel data and supplier data from the selected file and "List Example.xlsx"
                self.data = pd.read_excel(self.file_path)
                self.supp_df=pd.read_excel("List Example.xlsx")

                # Display import success message
                self.show_import_message()

                # Display data preview
                self.display_data_preview(self.data.head(10))

                # Display header preview
                self.display_header_preview(self.supp_df.columns)

                # Enable Continue button
                self.default_button.setEnabled(True)

            except Exception as e:
                error_message = f"An error occurred while importing the Excel file:\n{str(e)}"
                self.show_error_message("Import Error", error_message)


    def display_header_preview(self, headers):
        # Clear previous data from the header preview table
        self.header_preview_table.clear()
        self.header_preview_table.setRowCount(0)
        self.header_preview_table.setColumnCount(0)

        # Display headers in the header preview table
        self.header_preview_table.setColumnCount(len(headers))
        self.header_preview_table.setRowCount(1)
        self.header_preview_table.setHorizontalHeaderLabels(headers)

        for col in range(len(headers)):
            item = QTableWidgetItem(headers[col])
            self.header_preview_table.setItem(0, col, item)


    def display_data_preview(self, data):
        # Clear previous data from the table
        self.data_preview_table.clear()
        self.data_preview_table.setRowCount(0)
        self.data_preview_table.setColumnCount(0)

        # Display data in the table widget
        self.data_preview_table.setColumnCount(len(data.columns))
        self.data_preview_table.setRowCount(len(data))
        self.data_preview_table.setHorizontalHeaderLabels(data.columns)

        for row in range(len(data)):
            for col in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[row, col]))
                self.data_preview_table.setItem(row, col, item)


    def show_import_message(self):
        # Display a message box to inform the user about the successful import
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('Import Successful')
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(f'Data imported from:\n{self.file_path}')
        msg_box.exec_()


    def show_done_message(self):
        # Display a message box to inform the user that the process is complete
        QMessageBox.information(self, 'All Done', 'All done!')
        QApplication.quit()


    def show_error_message(self, title, message):
        # Display an error message box
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText(message)
        msg_box.exec_()


    def dragEnterEvent(self, event):
        # Accept drag-and-drop events if a valid .xlsx file is being dragged
        if event.mimeData().hasUrls() and event.mimeData().urls()[0].toString().endswith(".xlsx"):
            event.acceptProposedAction()


    def dropEvent(self, event):
        # Handle a drop event when a valid .xlsx file is dropped
        self.file_path = event.mimeData().urls()[0].toLocalFile()
        self.import_data()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.setGeometry(100, 100, 800, 600)
    window.show()
    sys.exit(app.exec_())
