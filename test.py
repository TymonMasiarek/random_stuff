import os
import pandas as pd
# pip install openpyxl

path = "C:\\Users\\tymon\Desktop\\" # wheret to look for files
save_path = "C:\\Users\\tymon\Desktop\\" # where to save the files
EXT = ".xlsx"

def every_excel_file(directory_path, extension=""):
    for root, subFolder, files in os.walk(directory_path):
        for item in files:
            if item.endswith(extension) : # ignore if not needed; xlsx
                fileNamePath = str(os.path.join(root,item))
                print(f"yielding: {fileNamePath}")
                yield fileNamePath

def process_file(file):
    file_name, _ = os.path.splitext(os.path.basename(file))
    print(f"Current file name: {file_name}.")
    df = pd.read_excel(file, index_col=None)

    df = df.drop(columns=df.columns[0:2]) # removes first two columns

    output_path = os.path.join(save_path, f"{file_name}1{EXT}")
    df.to_excel(output_path, index=False)
    print(f"File: {file_name} saved to new location with the new name.\n")


for file in every_excel_file(path, "xlsx"):
    try:
        process_file(file)
    except Exception as e:
        print(e)


    