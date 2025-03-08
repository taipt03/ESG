import os
import re
from collections import defaultdict
import openpyxl

# Paths
input_folder = "C:/Users/tuant/Downloads/data/data_txt"  # Folder containing the .txt files


# Step 4: Process all .txt files in the specified folder
file_paths = [
    os.path.join(input_folder, file)
    for file in os.listdir(input_folder)
    if file.endswith('.txt')
]

all_files = [
    os.path.join(input_folder, file)
    for file in os.listdir(input_folder)
    if file.endswith('.txt')
]
processed_files = []


for file_path in file_paths:
    filename = os.path.basename(file_path)
    match = re.match(r"(\d+)_.*_(\d{4})(?:_.*)?\.txt", filename)
    if match:
        processed_files.append(filename)  # Track successfully processed files




# Step 5: Write results to an Excel file
unprocessed_files = [
    os.path.basename(file) for file in all_files if os.path.basename(file) not in processed_files
]

if unprocessed_files:
    print("Files not processed:")
    for file in unprocessed_files:
        print(file)