# The AP Reports Automation Program is a user-friendly GUI program that streamlines the import, processing, and generating of formatted Excel reports.
# Author: Sean(Sunghyun) Park
# Version: 1.2
# Last Updated: 16-08-2023

import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QTableWidget, QTableWidgetItem, QStackedWidget, QMessageBox, QFileDialog, QLabel
from datetime import datetime


def highlight_row(row):
    """
    Apply conditional formatting to a row based on specific criteria.

    Parameters:
        row (pd.Series): A row from the DataFrame.

    Returns:
        list: List of styles for each row in the dataframe.
    """

    # Convert "Invoice Date" to datetime and calculate the difference in days
    inv_date = pd.to_datetime(row.loc["Invoice Date"], dayfirst=True)
    diff = (pd.to_datetime(datetime.now().strftime("%d/%m/%Y"),dayfirst=True) - inv_date).days
    
    # Set background color based on conditions
    if row.loc["IsCreditMemo"] == True:
        color = "#ff91a4"  # Red color for credit memos
    elif diff > 10:
        color = "#ffffcc"  # Light yellow color for older invoices
    else:
        color = "#FFFFFF"  # White color for other rows
    return [f'background-color: {color}' for r in row]


def excel_to_dict(supp_df):
    """
    Convert Excel data to dictionaries for sorting process.
    The excel file is must be 'Configuration.xlsx' and follow certain format.
    Please find attached excel file.

    Parameters:
        supp_df (pd.DataFrame): Supplier DataFrame.

    Returns:
        tuple: Tuple containing lists of information for sorting.
    """

    supp_list = []
    discriminant_list = []
    special_list = []

    # Iterate through columns in the supplier DataFrame
    for col in supp_df.columns:
        key = col
        values = list(supp_df[col].dropna())
        discriminant_list.append(values)
        supp_list.append(key)
        if len(values[0]) > 1:
            special_list.append(values[0])

    return supp_list, discriminant_list, special_list


def export_pages(sheet, page, entity, ind):
    """
    Export a DataFrame to an Excel file with customised formatting.

    Parameters:
        sheet (pd.DataFrame): Data to be exported.
        page (str): Name of the Excel sheet.
        entity (str): Entity information.
        ind (int): Index for filename differentiation.
    """

    # Generate the file path for the Excel export
    file_path = f"C:/Users/spark2/Desktop/SAP PO Upload/Python for PO/reports/{datetime.now().strftime('%d-%m-%Y')} {page} {entity} - {ind}.xlsx"
    
    # Create Excel writer with xlsxwriter engine
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        # Write DataFrame to the Excel sheet
        sheet.to_excel(writer, sheet_name=page, index=False, float_format = "%.2f")
        
        # Get the workbook and worksheet objects
        wb = writer.book
        ws = writer.sheets[page]

        # Iterate through columns to set column widths
        text_wrap_format = wb.add_format({'text_wrap': True, "num_format":'0.00'})
        num_format = wb.add_format({"num_format":'#.00'})
        for column in sheet:
            column_width = max(sheet[column].astype(str).map(len).max(), len(column))
            col_idx = sheet.columns.get_loc(column)
            ws.set_column(col_idx, col_idx, column_width + 1, num_format)
        
        # Apply text wrapping format to the "Comments" column
        comments_col_idx = sheet.columns.get_loc("Comments")
        ws.set_column(comments_col_idx, comments_col_idx, 35, text_wrap_format)
        
        # Apply conditional formatting using the highlight_row function
        sheet = sheet.style.apply(highlight_row, axis=1)
        sheet.to_excel(writer, sheet_name=page, index=False, float_format = "%.2f")

def main(data_df, supp_df):
    """
    Main function to process and export data.

    Parameters:
        data_df (pd.DataFrame): Master data DataFrame.
        supp_df (pd.DataFrame): Supplier data DataFrame.
    """

    # Get Supplier dictionary
    supp_list, discriminant_list, special_list = excel_to_dict(supp_df)

    # Clean the Master Data
    try: # Incase they do not have following columns
        data_df.drop(columns=["SC_Invoice_UniqueId"], inplace=True)
    except:
        skip
    data_df = data_df.loc[data_df['Status'].isin(["Pending", "Approved"])]
    data_df["Entity"].fillna("BLANK", inplace=True)

    # Save column list for later reference
    column_list = data_df.columns

    # Get unique entities and create a dictionary for each entity's suppliers
    entity_list = data_df['Entity'].unique()
    entity_dict = {entity: {supp: [] for supp in supp_list} for entity in entity_list}

    # Sort and group rows of data into appropriate pages
    for ind,row in data_df.iterrows():
        supp = row["Supplier Name"]
        entity = row["Entity"]

        if row["IsCreditMemo"]:
            try: # Incase they do not have following columns
                row["SubTotal"] *= -1
                row["Tax"] *= -1
                row["Total"] *= -1
            except:
                skip

        if supp in special_list:
            location = discriminant_list.index([supp])
            entity_dict[entity][supp_list[location]].append(row)
        else:
            ind = 0
            skip = 0
            while skip == 0:
                if supp[0].upper() in discriminant_list[ind]:
                    entity_dict[entity][supp_list[ind]].append(row)
                    skip = 1
                ind += 1

    count_dict = {supp: 0 for supp in supp_list}

    # Export data for each entity/supplier
    for entity in entity_list:
        if len(entity_dict[entity]) != 0:
            for page in supp_list:
                count_dict[page] += len(entity_dict[entity][page])
                if len(entity_dict[entity][page]) > 1:
                    ind = 1
                    sheet = pd.DataFrame(entity_dict[entity][page])
                    #sheet.drop(columns=sheet.columns[0], axis=1, inplace=True)
                    sheet.columns = column_list
                    sheet.sort_values(by=["Supplier Name", "Invoice Date", "PO #"], inplace=True)
                    sheet["Invoice Date"] = sheet["Invoice Date"].dt.strftime("%d/%m/%Y")
                    sheet["ReceivedDate"] = sheet["ReceivedDate"].dt.strftime("%d/%m/%Y")
                    if len(sheet) < 50:
                        export_pages(sheet, page, entity, ind)
                    else:
                        while len(sheet) >= 50:
                            sheet1 = sheet.iloc[:50]
                            export_pages(sheet1, page, entity, ind)
                            sheet = sheet.iloc[50:]
                            ind += 1
                        if len(sheet) != 0:
                            export_pages(sheet, page, entity, ind)
    
    count_df = pd.DataFrame.from_dict(count_dict,columns=["NO. PO Lines"], orient='index')
    count_df.T.to_excel("C:/Users/spark2/Desktop/SAP PO Upload/Python for PO/reports\Supplier Statistic.xlsx")


class ReportGeneratorApp(QMainWindow):
    """
    GUI application for importing, processing, and generating reports from Excel data.
    """

    def __init__(self):
        """
        Initialise the
        main application window and user interface.
        """

        super().__init__()
        self.init_ui()


    def init_ui(self):
        """
        Set up the user interface layout and widgets.
        """

        self.setWindowTitle('AP Report Automation')
        self.file_path = 0 # This will tell if user is drag and drop or import data from selection.

        # Create widgets
        self.import_button = QPushButton('Import Data OR Drag and Drop', self)
        self.import_button.clicked.connect(self.import_data)
        self.import_button.setMinimumHeight(60)  # Set the minimum height to 60 (adjust as needed)
        self.import_button.setMinimumWidth(200)

        # Customise the appearance of the import button
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

        self.path_label = QLabel('Reports will be saved in X:\Corporate\FINANCE\SHAREDSERVICES\Data Entry\DE Reports\Daily Reports SpendConsole', self)
        info_font = self.path_label.font()
        info_font.setPointSize(10)  # Set the font size to 12
        self.path_label.setFont(info_font)

        self.default_button = QPushButton('Process', self)
        self.default_button.clicked.connect(lambda: main(self.data, self.supp_df))
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
        main_layout.addWidget(self.path_label)
        main_layout.addWidget(self.default_button)  # Align the button at the bottom

        self.main_widget = QWidget()
        self.main_widget.setLayout(main_layout)

        # Stacked widget to switch between layouts (no change here)
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.addWidget(self.main_widget)
        # Add other layouts as needed

        self.setCentralWidget(self.stacked_widget)


    def import_data(self):
        """
        Handle data import from Excel files and display data previews.
        """

        # Open a file dialog to allow the user to select an Excel file for import
        file_dialog = QFileDialog()

        # If no file has been selected previously, set the file_path to the selected file's path
        if self.file_path == 0:
            self.file_path, _ = file_dialog.getOpenFileName(self, 'Open Excel File', '', 'Excel Files (*.xlsx);;All Files (*)')

        if self.file_path:
            try:
                # Read the Excel data and supplier data from the selected file and "List Example.xlsx"
                self.data = pd.read_excel(self.file_path)
                self.supp_df=pd.read_excel("Configuration.xlsx")

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
        """
        Display a preview of column headers in the header preview table.
        
        Args:
            headers (list of str): List of column headers.
        """
                
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
        """
        Display a preview of imported data in the data preview table.
        
        Args:
            data (pandas.DataFrame): Imported data to be displayed.
        """

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
        """
        Display a message box indicating a successful data import.
        """

        # Display a message box to inform the user about the successful import
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('Import Successful')
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(f'Data imported from:\n{self.file_path}')
        msg_box.exec_()


    def show_done_message(self):
        """
        Display a message box indicating successful report generation.
        """

        # Display a message box to inform the user that the process is complete
        QMessageBox.information(self, 'Successful', 'Reports successfully saved in Daily Reports SpendConsole folder')
        QApplication.quit()


    def show_error_message(self, title, message):
        """
        Display an error message box with a specified title and message.
        
        Args:
            title (str): Title of the error message box.
            message (str): Error message to be displayed.
        """

        # Display an error message box
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText(message)
        msg_box.exec_()


    def dragEnterEvent(self, event):
        """
        Handle drag-and-drop events for file import.
        
        Args:
            event (QDragEnterEvent): Drag-and-drop event object.
        """

        # Accept drag-and-drop events if a valid .xlsx file is being dragged
        if event.mimeData().hasUrls() and event.mimeData().urls()[0].toString().endswith(".xlsx"):
            event.acceptProposedAction()


    def dropEvent(self, event):
        """
        Handle drop events for file import.
        
        Args:
            event (QDropEvent): Drop event object.
        """

        # Handle a drop event when a valid .xlsx file is dropped
        self.file_path = event.mimeData().urls()[0].toLocalFile()
        self.import_data()


if __name__ == '__main__':
    """
    Entry point of the application. Creates the main application window and starts the event loop.
    """
    
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.setGeometry(100, 100, 800, 600)
    window.show()
    sys.exit(app.exec_())
