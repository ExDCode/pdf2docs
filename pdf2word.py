import glob
import win32com.client
import os


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


word = win32com.client.Dispatch("Word.Application")
word.visible = 0

pdfs_path = "./input/"  # folder where the .pdf files are stored
for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")):
    print(doc)
    filename = doc.split('\\')[-1]
    in_file = os.path.abspath(doc)
    print(in_file)
    wb = word.Documents.Open(in_file)
    reqs_path = "./output/"
    out_file = os.path.abspath(reqs_path +
                               filename[0:-4] + ".docx".format(i))
    print("outfile\n", out_file)
    wb.SaveAs2(out_file, FileFormat=16)  # file format for docx
    print("success...")
    wb.Close()

word.Quit()
