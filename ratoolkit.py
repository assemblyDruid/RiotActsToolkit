# Modules available with default python installations.
import os
import sys
import shutil
import inspect
from datetime import datetime
from enum import Enum

# Secondary python modules.
try:
    import tkinter as tk
    from tkinter import filedialog, scrolledtext
except ImportError:
    print("Riot Acts Toolkit | \033[93mRequired Python Modules\033[0m")
    print("-------------------------------------------")
    print("\033[91mTkinter\033[0m is not available. To continue, please install it via the following command:")
    print("\033[92m`python -m pip install tk`\033[0m")
    print("OR")
    print("\033[92m`<Your Python Version & Path> -m pip install tk`\033[0m")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("Riot Acts Toolkit | \033[93mRequired Python Modules\033[0m")
    print("-------------------------------------------")
    print("\033[91mPandas\033[0m is not available. To continue, please install it via the following command:")
    print("\033[92m`python -m pip install pandas`\033[0m")
    print("OR")
    print("\033[92m`<Your Python Version & Path> -m pip install pandas`\033[0m")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("Riot Acts Toolkit | \033[93mRequired Python Modules\033[0m")
    print("-------------------------------------------")
    print("\033[91mopenpyxl\033[0m is not available. To continue, please install it via the following command:")
    print("\033[92m`python -m pip install openpyxl`\033[0m")
    print("OR")
    print("\033[92m`<Your Python Version & Path> -m pip install openpyxl`\033[0m")
    sys.exit(1)

app = None
html_output_file_path = ""

ANSI_COLOR_RESET = "\033[0m"
ANSI_COLOR_RED = "\033[91m"
ANSI_COLOR_GREEN = "\033[92m"
ANSI_COLOR_YELLOW = "\033[93m"

VALID_INPUT_FILE_EXTENSIONS = [("Excel Files", "*.xlsx *.xls")]
VALID_OUTPUT_FILE_EXTENSIONS = [("HTML", "*.html")]

DEFAULT_INPUT_FILE_LOCATION = '../about/Data.xlsx'
DEFAULT_OUTPUT_FILE_LOCATION = './ratoolkit_output.html'
DEFAULT_BACKUP_FOLDER_LOCATION = './ratoolkit_backups'


class LogType(Enum):
    INFO = 0
    WARNING = 1
    ERROR = 2


def Log(log_string, log_type=LogType.INFO):
    log_time = '[ ' + datetime.now().strftime("%I:%M%p %m/%d/%Y") + ' ]'
    log_color_prefix = None
    log_type_str = None
    if log_type == LogType.INFO:
        log_color_prefix = ANSI_COLOR_GREEN
        log_type_str = "[ info ]"
    elif log_type == LogType.WARNING:
        log_color_prefix = ANSI_COLOR_YELLOW
        log_type_str = "[ warning ]"
    elif log_type == LogType.ERROR:
        caller_frame = inspect.stack()[1]
        caller_function_name = caller_frame.function
        caller_line_number = caller_frame.lineno
        log_color_prefix = ANSI_COLOR_RED
        log_type_str = "[ error ][ Fn: {}::{} ]".format(
            caller_function_name, caller_line_number)

    log_view_output = "{}{}: {}".format(log_type_str, log_time, log_string)
    print("{}{}{}".format(log_color_prefix, log_view_output, ANSI_COLOR_RESET))

    # Update Log View
    if app is not None:
        app.log_view.insert("end", log_view_output + '\n')
        app.log_view.see("end")
        app.log_view.update_idletasks()


def GetFileExtension(file_path):
    _, file_name = os.path.split(file_path)
    file_name_parts = file_name.split('.')
    if len(file_name_parts) > 1:
        return '.' + file_name_parts[-1]
    else:
        return ''


class RiotActsDataFile:
    def __init__(self):
        self.excel_data = None
        self.html_data = None

    def ConvertData(self, data_file_in, data_file_out):
        if not os.path.isfile(data_file_in):
            Log("The file '{}' does not exist. Ignoring...".format(
                data_file_in), LogType.WARNING)
            return False
        elif 'xls' not in GetFileExtension(data_file_in):
            Log("This application only accepts files with Excel extensions. Received: '{}'. Ignoring...".format(
                data_file_in), LogType.WARNING)
            return False
        
        self.excel_data = pd.read_excel(data_file_in)
        self.html_data = self.excel_data.to_html()

        with open(data_file_out, 'w') as f:
            f.write(self.html_data)

        return True


class App:
    def __init__(self, root):
        # Initialize top level class variables.
        self.root = root
        self.root.title("Riot Acts Toolkit")
        self.root.geometry("900x400")
        self.radf = RiotActsDataFile()

        # Construct graphical user interface.
        self.ConstructGUI()

        # Populate graphical user interface.
        self.PopulateGUI()

        # Warn on unexpected environment.
        if os.getcwd().split("/")[-1] != "Riot Acts Toolkit":
            Log("This script is not being run from the default location. Backups and other output will be stored to the current directory!", LogType.WARNING)

    def ConstructGUI(self):
        # Data file label.
        _column = 0
        _row = 0
        self.data_file_label = tk.Label(
            root, text="Input Excel file location:")
        self.data_file_label.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # Excel input file entry.
        _column = 1
        _row = 0
        self.excel_input_file_entry_stringvar = tk.StringVar()
        self.excel_input_file_tkentry = tk.Entry(
            root, textvariable=self.excel_input_file_entry_stringvar)
        self.excel_input_file_tkentry.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # Select Excel input file button.
        _column = 2
        _row = 0
        self.select_excel_input_file_tkbutton = tk.Button(
            root, text="Select Input File...", command=self.HandleSelectInputFile)
        self.select_excel_input_file_tkbutton.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # HTML output file label.
        _column = 0
        _row = 1
        self.html_output_file_tklabel = tk.Label(
            root, text="Output output HTML file location:")
        self.html_output_file_tklabel.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # HTML output file entry.
        _column = 1
        _row = 1
        self.html_output_file_entry_stringvar = tk.StringVar()
        self.output_file_path_entry = tk.Entry(
            root, textvariable=self.html_output_file_entry_stringvar)
        self.output_file_path_entry.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # Select HTML output file button.
        _column = 2
        _row = 1
        self.select_html_output_file_tkbutton = tk.Button(
            root, text="Select Output File...", command=self.HandleSelectOutputFile)
        self.select_html_output_file_tkbutton.grid(
            row=_row, column=_column, columnspan=1, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # Convert data file button.
        _column = 0
        _row = 2
        self.convert_tkbutton = tk.Button(
            root, text="Convert", command=self.ConvertData)
        self.convert_tkbutton.grid(
            row=_row, column=_column, columnspan=3, sticky="nsew")
        root.columnconfigure(_column, weight=1)

        # Log view.
        _column = 0
        _row = 3
        self.log_view = scrolledtext.ScrolledText(root)
        self.log_view.grid(row=_row, column=_column,
                           columnspan=3, sticky="nsew")
        root.columnconfigure(_column, weight=1)

    def PopulateGUI(self):
        # Input Excel File
        if os.path.isfile(DEFAULT_INPUT_FILE_LOCATION) == True:
            self.excel_input_file_entry_stringvar.set(
                DEFAULT_INPUT_FILE_LOCATION)

        # Input Excel File
        if os.path.isfile(DEFAULT_OUTPUT_FILE_LOCATION) == True:
            self.html_output_file_entry_stringvar.set(
                DEFAULT_OUTPUT_FILE_LOCATION)

    def HandleSelectInputFile(self):
        selected_file_path = filedialog.askopenfilename(
            title='Import Riot Acts Excel File', filetypes=VALID_INPUT_FILE_EXTENSIONS)
        if os.path.isfile(selected_file_path):
            self.excel_input_file_entry_stringvar.set(selected_file_path)
            Log("Updated input file path: {}".format(selected_file_path))

    def HandleSelectOutputFile(self):
        selected_file_path = filedialog.askopenfilename(
            title='Export Riot Acts HTML File', filetypes=VALID_OUTPUT_FILE_EXTENSIONS)
        if os.path.isfile(selected_file_path):
            self.excel_input_file_entry_stringvar.set(selected_file_path)
            Log("Updated output file path: {}".format(selected_file_path))

    def BackupFile(self, file):
        if not os.path.isfile(file):
            Log("Cannot back up non existant file: {}. Ignoring...".format(
                file), LogType.WARNING)
            return False

        os.makedirs(DEFAULT_BACKUP_FOLDER_LOCATION, exist_ok=True)
        backup_file_name = "{}/raBACKUP-{}-{}".format(DEFAULT_BACKUP_FOLDER_LOCATION,
                                                      datetime.now().strftime("%B-%d-%Y--%H:%M:%S%p"), os.path.basename(file))
        shutil.copy(file, backup_file_name)
        Log("Backed up {} ---> {}/{}".format(file,
            DEFAULT_BACKUP_FOLDER_LOCATION, backup_file_name))

    def ConvertData(self):
        # Check for existence of conversion files.
        input_excel_file = self.excel_input_file_entry_stringvar.get()
        output_html_file = self.html_output_file_entry_stringvar.get()

        if not os.path.isfile(input_excel_file):
            Log("Select an existing input Excel file; {} does not exist. Ignoring...".format(
                input_excel_file), LogType.WARNING)
            return False
        elif not os.path.isfile(output_html_file):
            Log("Select an existing output HTML file; {} does not exist. Ignoring...".format(
                output_html_file), LogType.WARNING)
            return False

        # Back up files.
        self.BackupFile(input_excel_file)
        self.BackupFile(output_html_file)

        # Convert data.
        Log("Converting data: {} ---> {}...".format(input_excel_file, output_html_file))
        if self.radf.ConvertData(input_excel_file, output_html_file) == True:
            Log("Success!")


# Create the main window
root = tk.Tk()

# Instantiate the App class.
app = App(root)

# Run the application.
root.mainloop()
