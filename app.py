import sys
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import csv
import re

def load_config(config_file):
    config = {}
    try:
        with open(config_file, 'r') as file:
            for line in file:
                key, value = line.strip().split('=', 1)
                config[key.strip()] = value.strip()
    except FileNotFoundError:
        print(f"Configuration file '{config_file}' not found. Exiting...")
        sys.exit(1)
    except ValueError:
        print(f"Invalid format in configuration file '{config_file}'. Each line must be in 'key=value' format.")
        sys.exit(1)
    return config

config_file = 'config.txt'
config = load_config(config_file)

# Set pytesseract and poppler paths from the configuration file
try:
    pytesseract.pytesseract.tesseract_cmd = config['tesseract_path']
    poppler_path = config['poppler_path']
except KeyError as e:
    print(f"Missing key in configuration file: {e}. Please ensure 'tesseract_path' and 'poppler_path' are defined.")
    sys.exit(1)

# Validate the paths
if not os.path.exists(pytesseract.pytesseract.tesseract_cmd):
    print(f"Tesseract executable not found at '{pytesseract.pytesseract.tesseract_cmd}'. Please check the configuration file.")
    sys.exit(1)

if not os.path.exists(poppler_path):
    print(f"Poppler path not found at '{poppler_path}'. Please check the configuration file.")
    sys.exit(1)


os.environ["PATH"] = poppler_path + ";" + os.environ["PATH"]

headers = ['filename', 'policynum', 'name', 'street3', 'street2', 'street1', 'csz']

def get_version_from_folder_name(this_folder_name):
    # returns the version name, upper crop area, and lower crop area to divide by
    folder_name = this_folder_name.lower()
    match folder_name:
        case folder_name if "bills" in folder_name:
            return "bills", 550, 3.5 #cropped area defined
        case folder_name if "grace" in folder_name:
            return "grace", 250, 3.5 #cropped area defined
        case folder_name if "claims" in folder_name:
            return "claims", 400, 3 #cropped area defined
        case folder_name if "v2n" in folder_name:
            return "v2n", 550, 3.45       
        case folder_name if "std_rtn" in folder_name:
            return "std_rtn", 450, 3.78
        case _:
            return "default_version", 200, 3



def extract_policy_num_and_save_to_line(file_name, line_to_csv):
    parts = file_name.split(' ')
    policy_num_substring = parts[0] if parts else file_name
    line_to_csv.append(policy_num_substring)



def clean_and_process_text(text, line_to_csv):
    text = text.strip()
    lines = text.split('\n')
    if not lines:  # all lines empty, write error messaging to csv
        line_to_csv.extend(["ERROR: Unable to read PDF", "", "", "ERROR: Manual entry required", ""])
        return

    substrings_to_remove = ["MassMutual", "Mass Mutual", "MIP W", "MIP B", "*", "Massachusetts Mutual", "insurance", "Billing Detail", "Premium", "Plan:", "Task ID:"]

    lines = [line for line in lines if line and not any(sub.lower() in line.lower() for sub in substrings_to_remove)]

    if not lines:  # Handle case where all lines are removed
        line_to_csv.extend(["ERROR: Unable to read PDF", "", "", "ERROR: Manual entry required", ""])
        return

    
    line_to_csv.append(lines[0])

    last_line = str(lines[-1])

    zip_pattern_plus4 = r'^\d{5}(-\d{4})?$'

    if re.match(zip_pattern_plus4, last_line):
        city_state = str(lines[-2])
        lines[-2] = city_state + " " + last_line
        lines.pop(-1)


    if len(lines) == 3:
        line_to_csv.extend(["", "", lines[1], lines[2]])
    elif len(lines) == 4:
        line_to_csv.extend(["", lines[1], lines[2], lines[3]])
    else:
        line_to_csv.extend(lines[1:5])


   


def process_pdf(pdf_file, pdf_short_name, writer, version, upper, lower_to_div):
    try:
        
        pages = convert_from_path(pdf_file, 300, first_page=1, last_page=1)
        page = pages[0]

        # Define crop area
        left = 0
        right = page.width // 2
        if version == "claims" or version == "std_rtn":
            right = 1625

        lower = page.height // lower_to_div
        cropped_image = page.crop((left, upper, right, lower))

        # code to save cropped image for further adjustments if need  vvv
        # cropped_image_path = f"{pdf_short_name}_cropped.png"
        # cropped_image.save(cropped_image_path)
        
        text = pytesseract.image_to_string(cropped_image)
        
        line_to_csv = [pdf_short_name]
        extract_policy_num_and_save_to_line(pdf_short_name, line_to_csv)

        clean_and_process_text(text, line_to_csv)
        
        print(f"line_to_csv: {line_to_csv}")
        # matches header length in case there is an exception
        while len(line_to_csv) < len(headers):
            line_to_csv.append("")

        writer.writerow(line_to_csv)
    except Exception as e:
        error_line = [pdf_short_name]
        error_line.extend(["ERROR", "ERROR: " + str(e)])
        writer.writerow(error_line)
        print(f"Error processing {pdf_file}: {e}")




root = Tk()
root.withdraw()
folder_to_process = askdirectory(title="Select Folder to Process PDFs")


if not folder_to_process:
    print("No folder selected, exiting...")
    sys.exit(1)

pdf_files = [f for f in os.listdir(folder_to_process) if f.endswith('.pdf')]

version, upper, lower_to_div = get_version_from_folder_name(folder_to_process)

if not pdf_files:
    print("No PDF files found in the selected folder. Exiting...")
    sys.exit(1)

data_file_name = os.path.basename(folder_to_process)+"-data.csv"
csv_file_path = os.path.join(folder_to_process, data_file_name)

with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(headers)


    for pdf_file in pdf_files:
        full_pdf_path = os.path.join(folder_to_process, pdf_file)
        if os.path.getsize(full_pdf_path) > 0:
            process_pdf(full_pdf_path, pdf_file, writer, version, upper, lower_to_div)
        else:
            continue



print(f"Processing complete. Data saved to {csv_file_path}.")
