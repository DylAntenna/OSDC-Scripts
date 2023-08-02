import os
import subprocess

root_folder = r"C:\Users\dylan\OneDrive\Desktop\test\watchfolder"
output_folder = r"C:\Users\dylan\OneDrive\Desktop\test\pdfout"
extension = ".docx"

# Recursively traverse the directory structure
for root, dirs, files in os.walk(root_folder):
    for file in files:
        if file.endswith(extension):
            file_path = os.path.join(root, file)
            output_file = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.pdf")
            subprocess.run(['SBCCmd', '-d', file_path, '-o', output_file])
