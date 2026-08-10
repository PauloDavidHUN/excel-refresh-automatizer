import os

#C:\Users\UserID\Document\Project\Folder\Path

BASE_PATH = os.path.join(os.path.expanduser("~"),  r"Document", r"Project", r"Folder", r"Path")



EXCEL_FAJLOK = {
    "EXCEL_FILE1"  : os.path.join(BASE_PATH, "your_excel_file1.xlsx"),
    "EXCEL_FILE2"  : os.path.join(BASE_PATH, "your_excel_file2.xlsx"),
}