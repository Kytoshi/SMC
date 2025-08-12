import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog
import aomoSAP
import logging

logging.basicConfig(
    filename="error_log.txt",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path_var.set(folder_selected)

def submit():
    try:
        SAPusername = SAPusername_var.get()
        SAPpassword = SAPpassword_var.get()
        PDBSusername = PDBSusername_var.get()
        PDBSpassword = PDBSpassword_var.get()
        folder_path = folder_path_var.get()
        aomoSAP.Report(SAPusername, SAPpassword, str(folder_path))
    except Exception as e:
        logging.error("An error occurred", exc_info=True)
    
    # You can now use these variables as needed
    print("SAP Username:", SAPusername)
    print("SAP Password:", SAPpassword)
    print("PDBS Username:", PDBSusername)
    print("PDBS Password:", PDBSpassword)
    print("Folder Path:", folder_path)


app = tb.Window(themename="flatly")  # Modern theme

app.title("AO MO SO Report Downloader")

# Variables
SAPusername_var = tb.StringVar()
SAPpassword_var = tb.StringVar()
PDBSusername_var = tb.StringVar()
PDBSpassword_var = tb.StringVar()
folder_path_var = tb.StringVar()

# UI Layout
frame = tb.Frame(app, padding=20)
frame.pack(fill=BOTH, expand=YES)

tb.Label(frame, text="SAP Username:", font=("Segoe UI", 12)).grid(row=0, column=0, sticky=W, pady=5)
username_entry = tb.Entry(frame, textvariable=SAPusername_var, font=("Segoe UI", 12))
username_entry.grid(row=0, column=1, pady=5, sticky=EW)

tb.Label(frame, text="SAP Password:", font=("Segoe UI", 12)).grid(row=1, column=0, sticky=W, pady=5)
password_entry = tb.Entry(frame, textvariable=SAPpassword_var, font=("Segoe UI", 12), show="*")
password_entry.grid(row=1, column=1, pady=5, sticky=EW)

tb.Label(frame, text="PDBS Username:", font=("Segoe UI", 12)).grid(row=3, column=0, sticky=W, pady=5)
username_entry = tb.Entry(frame, textvariable=PDBSusername_var, font=("Segoe UI", 12))
username_entry.grid(row=3, column=1, pady=5, sticky=EW)

tb.Label(frame, text="PDBS Password:", font=("Segoe UI", 12)).grid(row=4, column=0, sticky=W, pady=5)
password_entry = tb.Entry(frame, textvariable=PDBSpassword_var, font=("Segoe UI", 12), show="*")
password_entry.grid(row=4, column=1, pady=5, sticky=EW)

tb.Label(frame, text="Folder Path:", font=("Segoe UI", 12)).grid(row=5, column=0, sticky=W, pady=5)
folder_entry = tb.Entry(frame, textvariable=folder_path_var, font=("Segoe UI", 12))
folder_entry.grid(row=5, column=1, pady=5, sticky=EW)

browse_button = tb.Button(frame, text="Browse...", command=browse_folder)
browse_button.grid(row=5, column=2, padx=10, pady=5)

submit_button = tb.Button(frame, text="Submit", bootstyle=SUCCESS, command=submit)
submit_button.grid(row=6, column=1, pady=20)

# Make columns expand nicely
frame.columnconfigure(1, weight=1)

if __name__ == "__main__":
    app.mainloop()