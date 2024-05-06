from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui as qtg
import sys
import shutil
import qdarktheme
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import pandas as pd
import os
import glob
import re
from PyPDF2 import PdfReader, PdfWriter
import json
# from config2 import FORM_RECOGNIZER_ENDPOINT, FORM_RECOGNIZER_KEY  

if os.path.exists("config.json"):
    # Opening JSON file
    f = open('config.json')
    # returns JSON object as 
    # a dictionary
    data = json.load(f)
    # Closing file
    f.close()
    FORM_RECOGNIZER_ENDPOINT = data["FORM_RECOGNIZER_ENDPOINT"]
    FORM_RECOGNIZER_KEY = data["FORM_RECOGNIZER_KEY"]

#______________GUI_LAYOUT________________________
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        """        Set up the user interface for the main window.

        Args:
            MainWindow (QMainWindow): The main window object.
        """

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        MainWindow.setStyleSheet("")
        self.centralwidget = qtw.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = qtw.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        spacerItem = qtw.QSpacerItem(20, 10, qtw.QSizePolicy.Minimum, qtw.QSizePolicy.Fixed)
        self.verticalLayout_2.addItem(spacerItem)
        self.label = qtw.QLabel(self.centralwidget)
        font = qtg.QFont()
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setAlignment(qtc.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.verticalLayout_2.addWidget(self.label)
        spacerItem1 = qtw.QSpacerItem(20, 35, qtw.QSizePolicy.Minimum, qtw.QSizePolicy.Minimum)
        self.verticalLayout_2.addItem(spacerItem1)
        self.horizontalLayout_2 = qtw.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        spacerItem2 = qtw.QSpacerItem(10, 20, qtw.QSizePolicy.Fixed, qtw.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem2)
        self.verticalLayout = qtw.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = qtw.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.select_file_path = qtw.QLineEdit(self.centralwidget)
        self.select_file_path.setMinimumSize(qtc.QSize(0, 45))
        self.select_file_path.setMaximumSize(qtc.QSize(16777215, 45))
        font = qtg.QFont()
        font.setPointSize(10)
        self.select_file_path.setFont(font)
        self.select_file_path.setText("")
        self.select_file_path.setObjectName("select_file_path")
        self.horizontalLayout.addWidget(self.select_file_path)
        spacerItem3 = qtw.QSpacerItem(10, 20, qtw.QSizePolicy.Fixed, qtw.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem3)
        self.select_file_btn = qtw.QPushButton(self.centralwidget)
        self.select_file_btn.setMinimumSize(qtc.QSize(0, 45))
        font = qtg.QFont()
        font.setPointSize(11)
        self.select_file_btn.setFont(font)
        self.select_file_btn.setObjectName("select_file_btn")
        self.horizontalLayout.addWidget(self.select_file_btn)
        self.verticalLayout.addLayout(self.horizontalLayout)
        spacerItem4 = qtw.QSpacerItem(20, 10, qtw.QSizePolicy.Minimum, qtw.QSizePolicy.Fixed)
        self.verticalLayout.addItem(spacerItem4)
        self.progress_bar = qtw.QProgressBar(self.centralwidget)
        self.progress_bar.setMinimumSize(qtc.QSize(0, 35))
        font = qtg.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.progress_bar.setFont(font)
        self.progress_bar.setProperty("value", 0)
        self.progress_bar.setObjectName("progress_bar")
        self.verticalLayout.addWidget(self.progress_bar)
        spacerItem5 = qtw.QSpacerItem(20, 10, qtw.QSizePolicy.Minimum, qtw.QSizePolicy.Fixed)
        self.verticalLayout.addItem(spacerItem5)
        self.logs = qtw.QTextEdit(self.centralwidget)
        font = qtg.QFont()
        font.setPointSize(10)
        self.logs.setFont(font)
        self.logs.setReadOnly(True)
        self.logs.setObjectName("logs")
        self.verticalLayout.addWidget(self.logs)
        self.horizontalLayout_2.addLayout(self.verticalLayout)
        spacerItem6 = qtw.QSpacerItem(10, 20, qtw.QSizePolicy.Fixed, qtw.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem6)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        spacerItem7 = qtw.QSpacerItem(20, 10, qtw.QSizePolicy.Minimum, qtw.QSizePolicy.Fixed)
        self.verticalLayout_2.addItem(spacerItem7)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        qtc.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        """        Retranslate the user interface elements.

        This function is responsible for retranslating the user interface elements to the specified language.

        Args:
            self: The object instance.
            MainWindow: The main window of the application.
        """

        _translate = qtc.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Invoice Analyzer Alpha"))
        self.label.setText(_translate("MainWindow", "Invoice Analyzer Alpha"))
        self.select_file_path.setPlaceholderText(_translate("MainWindow", "No File Selected"))
        self.select_file_btn.setText(_translate("MainWindow", "Select File"))
        self.logs.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:10pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:11pt;\"><br /></p></body></html>"))

#________________________________________________________________

#___________________GUI_BACKEND__________________________________
class Ui(qtw.QMainWindow,Ui_MainWindow):
    def __init__(self):
        """        Initialize the Ui class.

        It initializes the Ui class and sets up the user interface. It also checks for the existence of the 'config.json' file
        in the current working directory and disables the 'select_file_btn' if the file is missing.

        Args:
            self: Instance of the Ui class.
        """

        super(Ui, self).__init__()
        # uic.loadUi('gui_layout.ui', self)
        self.setupUi(self)

        #select file btton clicked
        self.select_file_btn.clicked.connect(self.run_process)

        self.worker = None
        #check the config.json file exists in current working directory?
        if not os.path.exists("config.json"):
            self.logs.append("missing config.json file!!! \nPlease place the config.json file inside the current directory\nAfter that please rerun the program")
            self.select_file_btn.setEnabled(False)

        self.show()
    
    def run_process(self):
        """        Open a file dialog to select a PDF file and start a worker thread to process the selected file.

        Opens a file dialog to allow the user to select a PDF file. If a file is selected, it normalizes the file path,
        updates the UI elements, and starts a worker thread to process the selected file.
        """

        file_dialog = qtw.QFileDialog(self)
        file_dialog.setFileMode(qtw.QFileDialog.AnyFile)
        file_dialog.setNameFilter("PDF files (*.pdf)")

        if file_dialog.exec_():
            selected_file = file_dialog.selectedFiles()[0]
            selected_file = os.path.normpath(selected_file)
            self.select_file_path.setText(selected_file)
            self.select_file_btn.setEnabled(False)
            # print("Selected file:", selected_file)
            if self.worker is not None:
                self.worker.update_progress_bar.disconnect()
                self.worker.logs_data.disconnect()
                self.worker.work_finished.disconnect()
                self.worker_thread.quit()
                self.worker_thread.wait()
                self.logs.clear()
                
            self.worker= WorkerThread(pdf_file_path=str(self.select_file_path.text()))
            self.worker_thread = qtc.QThread()

            
            

            #connecting signal and slot
            self.worker.update_progress_bar.connect(self.update_progress_bar_func)
            self.worker.logs_data.connect(self.append_logs_data_func)
            self.worker.work_finished.connect(self.enable_widgets)
            #assign worker to thread
            self.worker.moveToThread(self.worker_thread)
            # Connect the thread's started signal to the run method of the WorkerThread
            self.worker_thread.started.connect(self.worker.run)
            
            self.worker_thread.start()

            

    def update_progress_bar_func(self, value):
        """        Update the value of a progress bar.

        This function updates the value of a progress bar with the given value.

        Args:
            self: The instance of the class.
            value (int): The new value for the progress bar.
        """

        self.progress_bar.setValue(value)
    
    def append_logs_data_func(self,text):
        """        Append text to the logs data.

        This function appends the input text to the logs data list.

        Args:
            text (str): The text to be appended to the logs data.
        """

        self.logs.append(text)

    def enable_widgets(self):
        """        Enable the select file button.

        This method enables the select file button to allow user interaction.
        """

        self.select_file_btn.setEnabled(True)   

#________________________________________________________________

#_______________worker_thread_class______________________________
class WorkerThread(qtc.QObject):
    work_finished = qtc.pyqtSignal()
    logs_data = qtc.pyqtSignal(str)
    update_progress_bar = qtc.pyqtSignal(int)
    def __init__(self,pdf_file_path):
        """        Initialize WorkerThread with the provided PDF file path.

        Args:
            pdf_file_path (str): The path to the PDF file.
        """

        super(WorkerThread, self).__init__()
        self.pdf_file_path = pdf_file_path
        self.progress_bar_value = 0
        self.progress_step = 0
    def run(self):
        """        Process the selected file, split the PDF, analyze and rename invoices, and output the results to a CSV file.

        Args:
            self (object): The instance of the class.
        """

        # Process the selected file
        output_folder = os.path.splitext(self.pdf_file_path)[0]
        # output_folder = os.path.dirname(self.pdf_file_path)
        self.split_pdf(self.pdf_file_path, output_folder)

        all_invoices_df = self.analyze_and_rename_invoices_in_directory(output_folder)

        # Output the DataFrame to a CSV file
        folder_name = os.path.basename(os.path.normpath(output_folder))
        csv_filename = f"{folder_name}_invoice_data.csv"
        csv_file_path = os.path.join(output_folder, csv_filename)
        all_invoices_df.to_csv(csv_file_path, index=False)

        # Update the label with the completion status
        # current_file_label.config(text=f"Done processing. Invoice data saved to {csv_file_path}")
        self.logs_data.emit(f"Done processing. Invoice data saved to {csv_file_path}")
        self.progress_bar_value = 100
        self.update_progress_bar.emit(self.progress_bar_value)
        self.work_finished.emit()
    
    def analyze_and_rename_invoice(self,file_path):
        """        Analyze the invoice document and rename it based on extracted fields.

        This function takes the file path of an invoice document, analyzes the document using Azure Form Recognizer,
        extracts relevant data, and renames the file based on the extracted fields.

        Args:
            file_path (str): The file path of the invoice document.

        Returns:
            pandas.DataFrame: A DataFrame containing the extracted invoice data.

        Raises:
            FileNotFoundError: If the input file_path does not exist.
        """

        # Load configuration
        

        # Setup the client
        document_analysis_client = DocumentAnalysisClient(
            endpoint=FORM_RECOGNIZER_ENDPOINT, 
            credential=AzureKeyCredential(FORM_RECOGNIZER_KEY)
        )

        # Analyze the document
        with open(file_path, "rb") as f:
            poller = document_analysis_client.begin_analyze_document(
                "prebuilt-invoice", document=f, locale="en-US"
            )
        invoices = poller.result()

        # Extract data and convert to DataFrame
        invoice_data = []
        for idx, invoice in enumerate(invoices.documents):
            for field_name, field in invoice.fields.items():
                if field.value is not None:
                    invoice_data.append({
                        "Field": field_name,
                        "Value": field.value if isinstance(field.value, str) else str(field.value),
                        "Confidence": field.confidence
                    })
            if "Items" in invoice.fields:
                for idx, item in enumerate(invoice.fields.get("Items").value):
                    for item_field_name, item_field in item.value.items():
                        if item_field.value is not None:
                            invoice_data.append({
                                "Field": f"Item {idx + 1} - {item_field_name}",
                                "Value": item_field.value if isinstance(item_field.value, str) else str(item_field.value),
                                "Confidence": item_field.confidence
                            })
        df = pd.DataFrame(invoice_data)

        # Helper function to get field value and sanitize it for use in a filename
        def get_field_value(field_name, default="Unknown"):
            """            Get the value of a field and sanitize it for use in a filename.

            This function retrieves the value of a specified field from a DataFrame and sanitizes it for use in a filename by removing illegal characters and stripping whitespace.

            Args:
                field_name (str): The name of the field to retrieve.
                default (str?): The default value to return if the field is not found. Defaults to "Unknown".

            Returns:
                str: The sanitized value of the specified field, or the default value if the field is not found.
            """

            value = df.loc[df['Field'] == field_name, 'Value']
            if not value.empty:
                # Remove illegal characters and strip whitespace
                sanitized_value = re.sub(r'[\\/*?:"<>|\n\r\t]', '', value.iloc[0])
                return sanitized_value.strip()
            return default

        # Extract fields for renaming the file
        invoice_id = get_field_value('InvoiceId').replace('/', '-').replace('\\', '-')
        invoice_date = get_field_value('InvoiceDate').replace('/', '-').replace('\\', '-')
        customer_name = get_field_value('CustomerName').replace('/', '-').replace('\\', '-')

        # Construct the filename
        filename = f"{invoice_id}_{invoice_date}_{customer_name}.pdf"
        renamed_invoice_directory = os.path.dirname(file_path)
        new_file_path = os.path.join(renamed_invoice_directory, filename)

        # Rename the file
        try:
            os.rename(file_path, new_file_path)
            # print(f"File saved as {new_file_path}")
            self.logs_data.emit(f"File saved as {new_file_path}")
            self.progress_bar_value = self.progress_bar_value + self.progress_step
            self.update_progress_bar.emit(self.progress_bar_value)
        except FileNotFoundError as e:
            # print(f"Error: {e}")
            self.logs_data.emit(f"Error: {e}")
            # print(f"Could not find the file {file_path}. Please check the path and try again.")
            self.logs_data.emit(f"Could not find the file {file_path}. Please check the path and try again.")
        return df


    def analyze_and_rename_invoices_in_directory(self,directory_path):
        """        Analyze and rename all invoices in the specified directory.

        This function processes each PDF file in the specified directory, analyzes and renames the invoices, and
        returns a DataFrame containing all invoice data.

        Args:
            directory_path (str): The path to the directory containing the PDF files.

        Returns:
            pandas.DataFrame: A DataFrame containing all invoice data.

        Raises:
            Exception: If an error occurs during the processing of any file.
        """

        # Get a list of all PDF files in the directory
        pdf_files = glob.glob(os.path.join(directory_path, '*.pdf'))
        
        # Initialize an empty DataFrame to store all invoice data
        all_invoices_df = pd.DataFrame()
        
        # Process each file
        for file_path in pdf_files:
            try:
                # Analyze and rename the current invoice
                invoice_df = self.analyze_and_rename_invoice(file_path)
                
                # Add a filename column to keep track of which invoice the data came from
                invoice_df['Filename'] = os.path.basename(file_path)
                
                # Concatenate the current invoice data to the all invoices DataFrame
                all_invoices_df = pd.concat([all_invoices_df, invoice_df], ignore_index=True)
            except Exception as e:
                # If there's an error, we print it and continue with the next file
                # print(f"Error processing {file_path}: {e}")
                self.logs_data.emit(f"Error processing {file_path}: {e}")

        return all_invoices_df


	# Function to split PDF into individual pages
    def split_pdf(self,input_pdf_path, output_folder):
        """        Split a PDF into individual pages.

        This function takes an input PDF file and splits it into individual pages, saving each page as a separate PDF file in the specified output folder.

        Args:
            input_pdf_path (str): The file path of the input PDF.
            output_folder (str): The folder where the individual pages will be saved.


        Raises:
            FileNotFoundError: If the input PDF file does not exist.
        """

        # Ensure the output directory exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        else:
            shutil.rmtree(output_folder)
            os.makedirs(output_folder)
            

        # Read the input PDF
        input_pdf = PdfReader(input_pdf_path)
        self.progress_step = int(100/len(input_pdf.pages))
        # Split each page of the PDF
        for i in range(len(input_pdf.pages)):
            output_pdf = PdfWriter()
            output_pdf.add_page(input_pdf.pages[i])
            output_filename = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(input_pdf_path))[0]}({i+1}).pdf")
            with open(output_filename, "wb") as outputStream:
                output_pdf.write(outputStream)
            # print(f"Created: {output_filename}")
            self.logs_data.emit(f"Created: {output_filename}")

#________________________________________________________________


#______________________________Calling_instance__________________
app = qtw.QApplication(sys.argv)
#light theme
qdarktheme.setup_theme("light")
#darktheme
# qdarktheme.setup_theme()

window = Ui()
app.exec_()
#________________________________________________________________