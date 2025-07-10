# SMC PROJECTS

All projetcts listed have been configured to only work with an additional file that includes work sensitive information.

## Automated Reports

### B39Delta

**Description:** Downloads filtered data from internal company site and replaces the pre existing report location while archiving the previous version.

<u> Files: </u>

- 39DRExtract.py

- config.json (not included)

---

### GI GR PENDING REPORT

**Description:** Downloads transaction history from SAP and then formats the excel engine file into legible summarized format for teams to utilize. 2 Versions: 1 Power Automate Script to structure before creating the python file.

**<u> Files: </u>**

- gigrXL.py

- Power Automate Scripts.txt

---

### AOMO Report

**Description:** Uses Power Automate to download transaction history from SAP, power automate also activates a script which downloads data from company internal website (Rename.py) and then finally modifies the Excel Engine file to a digestible summarized format for teams to read and utilize.

**<u> Files: </u>**

#### Excel

- aomoXL.py

- config.json (not included)

#### SAP Data

- Power Automate Functions.txt

#### Website Data Downloader

- Rename.py

- config.json (not included)

### B5B7 Report

**Description:** Variety of python scripts to aid with downloading large amounts of data quickly with some being broken into multiple parts in case of an error and not needing to restart the entire process again.

**<u> Files: </u>**

- B5B7Download.py
- CS07CSEDownload.py
- CSEReport.py
- ItemsDistrPDBS.py
- UsageXL.py
- Weekly Rpt Func.py
  
### 131819

**Description:** Used to adjust power queries for certain excel report

**<u> Files: </u>**

- folderChange.py
  
## Tools

### AutoKeyer

**Description:** tool to automate physical count keying in based off excel sheet to increase efficency. Updated to include visual UI instead of terminal prompts.

**<u> Files: </u>**

- AutoKey.py

  - Asks for an input excel file, sheet name, and starting cell of the table you are inputing

- minus1.py

  - If a mistake was made in the key program, to reset the list, replaces all boxes with "-1" as it is a number that can be replaced without popping up a error message.
  - Once minus1.py has run, you can use the AutoKey.py script again.

- PICountKeyer_v2.py (2025)

  - Version 2 of the program which combines both functionality of "AutoKey.py" and "minus1.py" into a singular program which can be navigated with operational UI elements.

---

### Excel Converter

**Description:** Helper script which if a file is downloaded as a xls file instead of a xlsx file, will convert multiple files at the same time from a xls to a xlsx file.

**<u> Files: </u>**

- excelConverter.py
