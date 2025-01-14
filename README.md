# SMC PROJECTS

## Automated Reports

### B39Delta

**Description:** Downloads filtered data from internal company site and replaces the pre existing report location while archiving the previous version.

**<u>Files:</u>**

- 39DRExtract.py

---

### GI GR PENDING REPORT

**Description:** Downloads transaction history from SAP utilizing Power Automate integration with the SAP program and then formats the excel engine file into legible summarized format for teams to utilize.

**<u>Files:</u>**

- gigrXL.py

- Power Automate Scripts.txt

---

### AOMO Report

**Description:** Uses Power Automate to download transaction history from SAP, power automate also activates a script which downloads data from company internal website (Rename.py) and then finally modifies the Excel Engine file to a digestible summarized format for teams to read and utilize.

**<u>Files:</u>**

#### Excel

- aomoXL.py

- config.json (not included)

#### SAP Data

- Power Automate Functions.txt

#### Website Data Downloader

- Rename.py

- config.json (not included)
