import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import subprocess
import json
import logging
from datetime import datetime
import os
import threading

class AppConfig:
    def __init__(self, config_path):
        with open(config_path, 'r') as config_file:
            self.config = json.load(config_file)

    def get_path(self, key_path):
        # key_path is a string like 'ui_settings/window_title'
        keys = key_path.split('/')
        value = self.config
        for key in keys:
            value = value.get(key)
            if value is None:
                return None
        if isinstance(value, str):
            # Convert relative path to absolute path
            return os.path.join(os.path.dirname(__file__), value)
        return value
    
# Configure logging to both file and console
log_file_path = os.path.join(os.path.dirname(__file__), 'application.log')
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(message)s',
                    handlers=[
                        logging.FileHandler(log_file_path),
                        logging.StreamHandler()
                    ])

logging.info('This is a test log message.')

# Determine the path to the Python executable
# For a virtual environment located in the same directory as your script
venv_path = os.path.join(os.path.dirname(__file__), 'venv')
python_executable = os.path.join(venv_path, 'Scripts', 'python.exe') if os.name == 'nt' else os.path.join(venv_path, 'bin', 'python')

# Path to the Unexcused Absences.py script (assuming it's in the same directory as this script)
script_path = os.path.join(os.path.dirname(__file__), 'Unexcused Absences.py')

# Your CSV file path and dates
csv_file_path = 'path/to/your/csv_file.csv'  # Adjust as needed
start_date = '2022-01-01'  # Example start date
end_date = '2022-12-31'  # Example end date

# Construct the command
command = [python_executable, script_path, csv_file_path, start_date, end_date]

# For demonstration, printing the command
print(command)

# This command can then be executed using subprocess.run(command) or a similar method
# Load configuration
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r') as config_file:
    config = json.load(config_file)

# Create the main application window
root = tk.Tk()
root.title(config['ui_settings']['window_title'])

# Set window size and position
window_width = config['ui_settings']['window_width']
window_height = config['ui_settings']['window_height']
root.geometry(f"{window_width}x{window_height}")
position_top = int(root.winfo_screenheight() / 2 - window_height / 2)
position_right = int(root.winfo_screenwidth() / 2 - window_width / 2)
root.geometry("+{}+{}".format(position_right, position_top))

# Global variables for selected file and date range
csv_file_path = None
selected_start_date = None
selected_end_date = None

def run_subprocess(script_path, *args):
    """Runs a subprocess with the given script and arguments, and logs the output."""
    def target():
        command = ["python", script_path] + list(args)
        logging.info(f"Running command: {command}")
        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            logging.info(f"Subprocess output: {result.stdout}")
            if result.stderr:
                logging.error(f"Subprocess error: {result.stderr}")
        except subprocess.CalledProcessError as e:
            logging.error(f"Subprocess failed with error: {e}")
            messagebox.showerror("Subprocess Error", f"Failed to run script: {script_path}\nError: {e}")
            if e.stdout:
                logging.info(f"Subprocess stdout: {e.stdout}")
            if e.stderr:
                logging.error(f"Subprocess stderr: {e.stderr}")

    thread = threading.Thread(target=target)
    thread.start()

def convert_date_format(date_str):
    logging.info(f"Attempting to convert date format for: {date_str}")
    try:
        formatted_date = datetime.strptime(date_str, '%m-%d-%Y').strftime('%Y-%m-%d')
        logging.info(f"Converted date format to: {formatted_date}")
        return formatted_date
    except ValueError as e:
        logging.error(f"Error converting date format: {e}")
        messagebox.showerror("Date Format Error", "Invalid date format. Please use 'MM-DD-YYYY'.")
        return None

def select_file():
    """Function to select a file and update the label."""
    global csv_file_path
    selected_file_path = filedialog.askopenfilename()
    if selected_file_path:
        csv_file_path = selected_file_path
        file_name = os.path.basename(csv_file_path)
        # Split the file name into two lines after 25 characters
        file_name = file_name[:25] + '\n' + file_name[25:]
        file_name_label.config(text=f"{file_name}", foreground="green", anchor='e', justify=tk.LEFT)
        logging.info(f"Selected file: {csv_file_path}")

        # Update the config.json file with the selected file path
        config['entrance_notifications_path']['csv_path'] = csv_file_path
        with open(config_path, 'w') as config_file:
            json.dump(config, config_file, indent=2)
    else:
        messagebox.showwarning("No File Selected", "Please select a file.")

def run_subprocess(script_path, *args):
    """Runs a subprocess with the given script and arguments, and logs the output."""
    try:
        result = subprocess.run(["python", script_path] + list(args), check=True, capture_output=True, text=True)
        logging.info(f"Subprocess output: {result.stdout}")
        if result.stderr:
            logging.error(f"Subprocess error: {result.stderr}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Subprocess failed with error: {e}")
        messagebox.showerror("Subprocess Error", f"Failed to run script: {script_path}\nError: {e}")
        if e.stdout:
            logging.info(f"Subprocess stdout: {e.stdout}")
        if e.stderr:
            logging.error(f"Subprocess stderr: {e.stderr}")

def request_date_range_en():
    global csv_file_path, selected_start_date, selected_end_date
    if not csv_file_path or not selected_start_date or not selected_end_date:
        messagebox.showwarning("Missing Information", "Please select a file and date range.")
        return
    formatted_start_date = convert_date_format(selected_start_date)
    formatted_end_date = convert_date_format(selected_end_date)
    logging.info(f"Running Entrance Notifications.py with start date {formatted_start_date} and end date {formatted_end_date}")
    run_subprocess('./Entrance Notifications.py', csv_file_path, formatted_start_date, formatted_end_date)

def request_date_range_ue():
    """Function to request a date range and call the Unexcused Absences.py script."""
    global csv_file_path
    if csv_file_path is None:
        messagebox.showwarning("No File Selected", "Please select a file.")
        return
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()
    if not start_date or not end_date:
        messagebox.showwarning("Invalid Input", "Please enter a valid start and end date.")
        return
    start_date = datetime.strptime(start_date, "%m-%d-%Y").strftime("%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%m-%d-%Y").strftime("%Y-%m-%d")
    command = ["python", './Unexcused Absences.py', csv_file_path, start_date, end_date]
    subprocess.run(command)
    run_subprocess('./Unexcused Absences.py', csv_file_path, start_date, end_date)

def request_progress_reports():
    """Function to request a date and call the progress reports script."""
    global csv_file_path, selected_start_date, selected_end_date
    if csv_file_path is None:
        messagebox.showwarning("No File Selected", "Please select a file.")
        return
    if selected_start_date != selected_end_date:
        messagebox.showwarning("Invalid Input", "Start and end date must be the same for progress reports.")
        return
    report_date = datetime.strptime(selected_start_date, "%m-%d-%Y").strftime("%Y-%m-%d")
    script_path = os.path.join(os.path.dirname(__file__), 'Progress Reports.py')
    run_subprocess('./Progress Reports.py', csv_file_path, report_date)  # or end_date, if they are the same

def select_date_range():
    global selected_start_date, selected_end_date
    selected_start_date = start_date_entry.get()
    selected_end_date = end_date_entry.get()
    logging.info(f"Selected start date: {selected_start_date}")

    def format_date_range(start_date, end_date):
        # Parse the dates from strings to datetime objects
        start_date_obj = datetime.strptime(start_date, "%m-%d-%Y")
        end_date_obj = datetime.strptime(end_date, "%m-%d-%Y")

        if start_date_obj == end_date_obj:
            # If both dates are the same, format as a single date "Month DD, YYYY"
            return start_date_obj.strftime("%B %d, %Y")
        else:
            # If the dates are different, format as "Month DD, YYYY to Month DD, YYYY"
            # Add a newline character before "to" to make the second date appear on the next line
            return f"{start_date_obj.strftime('%B %d, %Y')} to\n{end_date_obj.strftime('%B %d, %Y')}"
    
    logging.info(f"Selected end date: {selected_end_date}")

    formatted_start_date = convert_date_format(selected_start_date)
    formatted_end_date = convert_date_format(selected_end_date)

    if not formatted_start_date or not formatted_end_date:
        return

    # Call the new function to format the date range
    date_range_str = format_date_range(selected_start_date, selected_end_date)

    # Log before updating the GUI
    logging.info(f"Displaying date range: {date_range_str}")

    selected_date_range_label.config(text=date_range_str)

# Layout and Widgets Configuration
top_frame = ttk.Frame(root)
top_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

file_display_frame = ttk.Frame(root)
file_display_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

button_frame = ttk.Frame(root)
button_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

# Adjust the grid configuration for centering buttons
top_frame.columnconfigure(0, weight=1)
top_frame.columnconfigure(1, weight=0)  # Adjust for button
top_frame.columnconfigure(2, weight=0)  # Adjust for button
top_frame.columnconfigure(3, weight=1)

# Buttons
btn_select_file = ttk.Button(top_frame, text="Select File", command=select_file, width=config['ui_settings']['button_width'])
btn_select_file.grid(row=0, column=1, padx=5, pady=5)

btn_change_date_range = ttk.Button(top_frame, text="Change Date Range", command=select_date_range, width=config['ui_settings']['button_width'])
btn_change_date_range.grid(row=0, column=2, padx=5, pady=5)

# File and Date Range Display
file_label = tk.Label(top_frame, text="Selected File:", foreground="black")
file_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")

file_name_label = tk.Label(top_frame, text="None", foreground="black", wraplength=window_width, height=2, anchor='e', justify=tk.LEFT)
file_name_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")

def select_file():
    """Function to select a file and update the label."""
    global csv_file_path
    selected_file_path = filedialog.askopenfilename()
    if selected_file_path:
        csv_file_path = selected_file_path
        file_name = os.path.basename(csv_file_path)
        # Split the file name into two lines after 25 characters
        file_name = file_name[:25] + '\n' + file_name[25:]
        file_name_label.config(text=f"{file_name}", foreground="green", anchor='e', justify=tk.LEFT)
        logging.info(f"Selected file: {csv_file_path}")

        # Update the config.json file with the selected file path
        config['entrance_notifications_path']['csv_path'] = csv_file_path
        with open(config_path, 'w') as config_file:
            json.dump(config, config_file, indent=2)
    else:
        messagebox.showwarning("No File Selected", "Please select a file.")

if __name__ == "__main__":
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    app_config = AppConfig(config_path)

    window_title = app_config.get_path('ui_settings/window_title')

    print(f"Window title: {window_title}")

# New labels to display "Report Date:" and the selected date range
report_date_label = ttk.Label(top_frame, text="Report Date:")
report_date_label.grid(row=3, column=1, padx=5, pady=5, sticky="w")

selected_date_range_label = ttk.Label(top_frame, text="Not selected", wraplength=window_width)
selected_date_range_label.grid(row=4, column=1, padx=5, pady=5, sticky="w")

# Start Date Label
start_date_label = ttk.Label(top_frame, text="Start Date")
start_date_label.grid(row=1, column=2, padx=5, pady=5, sticky="w")

start_date_entry = DateEntry(top_frame, date_pattern='mm-dd-yyyy')
start_date_entry.grid(row=2, column=2, padx=5, pady=5, sticky="ew")

# End Date Label
end_date_label = ttk.Label(top_frame, text="End Date")
end_date_label.grid(row=3, column=2, padx=5, pady=5, sticky="w")

end_date_entry = DateEntry(top_frame, date_pattern='mm-dd-yyyy')
end_date_entry.grid(row=4, column=2, padx=5, pady=5, sticky="ew")

# New label to display the selected date range
selected_date_range_label = ttk.Label(top_frame, text="Not selected", wraplength=window_width)
selected_date_range_label.grid(row=4, column=1, padx=5, pady=5, sticky="w")

# Operation Buttons
btn_entrance_notifications = ttk.Button(button_frame, text="Entrance Notifications", command=request_date_range_en, width=config['ui_settings']['button_width'])
btn_entrance_notifications.pack(fill=tk.X, expand=True, padx=10, pady=2)

# Create a new button for Progress Reports
btn_progress_reports = ttk.Button(button_frame, text="Progress Reports", command=request_progress_reports, width=config['ui_settings']['button_width'])
btn_progress_reports.pack(fill=tk.X, expand=True, padx=10, pady=2)

btn_unexcused_absences = ttk.Button(button_frame, text="Unexcused Absences", command=request_date_range_ue, width=config['ui_settings']['button_width'])
btn_unexcused_absences.pack(fill=tk.X, expand=True, padx=10, pady=2)
root.mainloop()
