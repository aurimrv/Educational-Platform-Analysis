import os
import re
import pandas as pd

def extract_numbers_auto_subfolders():
    # Detects the directory where the script is running
    root_dir = os.getcwd()
    
    print(f"Starting scan in the current directory and subdirectories:\n{root_dir}\n")
    
    pattern = re.compile(r"S_(\d+)_", re.IGNORECASE)
    data = []

    # Traverses the current directory and all subfolders
    for root, _dirs, files in os.walk(root_dir):
        relative_subfolder = os.path.relpath(root, root_dir)
        if relative_subfolder == ".":
            relative_subfolder = "(Root Directory)"

        for file in files:
            match = pattern.search(file)
            if match:
                number = int(match.group(1))
                data.append({
                    "Number": number,
                    "Subfolder": relative_subfolder,
                    "File Name": file,
                    "Full Path": os.path.join(root, file)
                })

    if not data:
        print("No files matching the pattern were found.")
        return

    # Structure data and sort numerically
    df = pd.DataFrame(data)
    df = df.sort_values(by="Number")

    excel_output = "beecrowd_numbers.xlsx"
    csv_output = "beecrowd_numbers.csv"

    df.to_excel(excel_output, index=False)
    df.to_csv(csv_output, index=False, encoding="utf-8-sig")

    print("Scan completed successfully!")
    print(f"Total files found: {len(df)}")
    print(f"Spreadsheet generated: {os.path.abspath(excel_output)}")

if __name__ == "__main__":
    extract_numbers_auto_subfolders()