import glob
import os
from pdf2docx import Converter


def check_and_create_folders():
    current_path = os.getcwd()
    folder_name_created = 'input'
    print(current_path)
    if not os.path.exists(current_path + '/' + folder_name_created):
        os.makedirs(folder_name_created)
        print(f"Folder '{folder_name_created}' created.")
    else:
        print(f"Folder '{folder_name_created}' already exists.")

    folder_name_created = 'output'
    if not os.path.exists(current_path + '/' + folder_name_created):
        os.makedirs(folder_name_created)
        print(f"Folder '{folder_name_created}' created.")
    else:
        print(f"Folder '{folder_name_created}' already exists.")


# Create input and output folders if they don't exist in the current directory
check_and_create_folders()


pdfs_path = "./input/"  # folder where the .pdf files are stored

# Checking for multiple files in the input folder
for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")):
    print(doc)
    filename = doc.split('\\')[-1]
    print(filename)
    in_file = os.path.abspath(doc)
    print(in_file)
    reqs_path = "./output/"
    out_file = os.path.abspath(reqs_path +
                               filename[0:-4] + ".docx".format(i))
    print(out_file)

    # Conversion part
    cv = Converter(in_file)
    cv.convert(out_file)
    cv.close()
