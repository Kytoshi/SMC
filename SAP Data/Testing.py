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

class FormPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        # Variables
        self.SAPusername_var = tb.StringVar()
        self.SAPpassword_var = tb.StringVar()
        self.PDBSusername_var = tb.StringVar()
        self.PDBSpassword_var = tb.StringVar()
        self.folder_path_var = tb.StringVar()

        # Widgets
        self.sapUserText = tb.Label(self, 
            text="SAP Username:", 
            font=("Segoe UI", 12)
            )
        self.sapUserText.grid(row=0, column=0, sticky=W, pady=5)

        self.sapUserForm = tb.Entry(self, 
            textvariable=self.SAPusername_var, 
            font=("Segoe UI", 12)
            )
        self.sapUserForm.grid(row=0, column=1, pady=5, sticky=EW)

        self.sapPassText = tb.Label(self, 
            text="SAP Password:", 
            font=("Segoe UI", 12)
            )
        self.sapPassText.grid(row=1, column=0, sticky=W, pady=5)

        self.sapPassForm = tb.Entry(self, textvariable=self.SAPpassword_var, font=("Segoe UI", 12), show="*").grid(row=1, column=1, pady=5, sticky=EW)

        self.pdbsUserText = tb.Label(self, text="PDBS Username:", font=("Segoe UI", 12)).grid(row=3, column=0, sticky=W, pady=5)
        self.pdbsUserForm = tb.Entry(self, textvariable=self.PDBSusername_var, font=("Segoe UI", 12)).grid(row=3, column=1, pady=5, sticky=EW)

        self.pdbsPassText = tb.Label(self, text="PDBS Password:", font=("Segoe UI", 12)).grid(row=4, column=0, sticky=W, pady=5)
        self.pdbsPassForm = tb.Entry(self, textvariable=self.PDBSpassword_var, font=("Segoe UI", 12), show="*").grid(row=4, column=1, pady=5, sticky=EW)

        self.folderText = tb.Label(self, text="Folder Path:", font=("Segoe UI", 12)).grid(row=5, column=0, sticky=W, pady=5)
        self.folderForm = tb.Entry(self, textvariable=self.folder_path_var, font=("Segoe UI", 12)).grid(row=5, column=1, pady=5, sticky=EW)

        browse_btn = tb.Button(self, text="Browse...", command=self.browse_folder)
        browse_btn.grid(row=5, column=2, padx=10, pady=5)

        submit_btn = tb.Button(self, text="Submit", bootstyle=SUCCESS, command=self.submit)
        submit_btn.grid(row=6, column=1, pady=20)

        self.columnconfigure(1, weight=1)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)

    def submit(self):
        try:
            SAPusername = self.SAPusername_var.get()
            SAPpassword = self.SAPpassword_var.get()
            PDBSusername = self.PDBSusername_var.get()
            PDBSpassword = self.PDBSpassword_var.get()
            folder_path = self.folder_path_var.get()

            # aomoSAP.Report(SAPusername, SAPpassword, folder_path)

            # Go to next page
            self.controller.show_frame("SecondPage")

        except Exception as e:
            logging.error("An error occurred", exc_info=True)

class SecondPage(tb.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, padding=20)
        self.controller = controller

        tb.Label(self, text="Choose an option:", font=("Segoe UI", 14))
        .pack(pady=10)

        tb.Button(self, text="Option 1", bootstyle=INFO, width=20, command=self.option1).pack(pady=5)
        tb.Button(self, text="Option 2", bootstyle=INFO, width=20, command=self.option2).pack(pady=5)
        tb.Button(self, text="Back", bootstyle=SECONDARY, width=20, command=lambda: controller.show_frame("FormPage")).pack(pady=20)

    def option1(self):
        print("Option 1 clicked!")

    def option2(self):
        print("Option 2 clicked!")

class App(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("AO MO SO Report Downloader")

        container = tb.Frame(self)
        container.pack(fill=BOTH, expand=YES)

        self.frames = {}

        for F in (FormPage, SecondPage):
            page_name = F.__name__
            frame = F(container, self)
            frame.grid(row=0, column=0, sticky="nsew")
            self.frames[page_name] = frame

        self.show_frame("FormPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

if __name__ == "__main__":
    app = App()
    app.mainloop()
