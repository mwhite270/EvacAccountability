Project Description
======================================
This python script was originally made for use by a large chemical company. It was built to scrape the badge logs from a SaaS website the company used. It then used Excel to generate an accountability report.

The script was frozen and distributed as a .exe that could be used by any employee (no Python knowledge required).

The result is that what once took a team of people 30+ minutes, now takes 1 person ~ 5 minutes.

Jupyter/Colab Notebook (.ipynb) Readme
======================================
The notebook (built in Google Colab) in this repository was created to provide a visualization of how the code works.

The .py file in the repository can be run with the provided Excel workbook.

Python Script (.py) Readme
======================================

Summary
-------
The `Evac_SIPnologinvector.py` script process roll-call / badge-reader exports and generates an Accountability / Shelter-in-Place (SIP) report in Excel using xlwings and pandas. The script reads data pasted into an accompanying Excel workbook (`Evac_Report_Toolvec.xlsm`), classifies personnel (e.g., "Never Mustered", "Mustered", "Badged After Incident", "No Badge Data"), and can enrich records from a security badge log (visitor badges -> phone number / name).

This copy of the script is sanitized for demo purposes and contains commented-out sections that originally performed web scraping to pull badge logs. The provided Excel workbook has been populated with dummy data, since the web scraping and network file portions of the code are not usable without Company access.

Features
--------
- Loads Facility and Shelter In Place roll-call tables from an Excel workbook.
- Vectorized classification of each row into status categories based on badge access time and configured incident window.
- Matches visitor badge logs from a security export to fill `Phone Number` and `Last Name` where applicable.
- Writes filtered lists back to sheets in the Excel workbook (Never Mustered, Badged After Incident, Mustered, etc.).

Prerequisites
-------------
- Script uses xlwings which requires Microsoft Excel for Windows and macOS if running interactively.
- Python 3.8+ recommended.
- Microsoft Excel installed and accessible by xlwings.
- The Excel workbook `Evac_Report_Toolvec.xlsm` is used by the demo script (the script calls `xw.Book.set_mock_caller()` and then `xw.Book.caller()` so the workbook is expected to be available in the same path or opened by Excel). Edit the workbook name in the script if it has been renamed.

Installation
------------
1. Create and activate a virtual environment (optional):

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

Files created/required
----------------------
- `Evac_SIPnologinvector.py` — main script (already present).
- `Evac_Report_Toolvec.xlsm` — Excel workbook the script reads/writes (provided in repository).
- `requirements.txt` — Python dependencies for the script (this file).

How to run
----------
This script is designed to be run with the Excel workbook open. Typical workflows:

  1. Open `Evac_Report_Toolvec.xlsm` in Excel.
    - It asks if you want to open in Read-Only. It is recommended to do so.
  2. On the StartHere tab, select the Plant from the dropdown.
    - Use Plant1 for the demo.
  4. On the StartHere tab, enter the incident time and enter the incident time.
    - Use the default time for the demo.
  5. Run: `Evac_SIPnologinvector.py` 

- A note on the Excel file:
    - There are no passwords for the sheets. They have just been locked to prevent accidental overwriting. 
    - The names and phone numbers are randomly generated/fake.
    - The highlighted name/badge cells are there to flag cases of some of the sorting that is done.

Notes about interactive features and commented code
--------------------------------------------------
- The script contains commented-out Selenium-based browser automation used originally to download roll-call reports. That is intentionally disabled in this version, as there is no possible connection to the site that it was coded to access. All links shown in the script are invalid/fake.
- Network calls (requests + negotiate auth) are also commented out. Dummy data was added to the demo file to simulate what they would pull (ex: `SecurityEntry` sheet).

Troubleshooting
---------------
- xlwings errors: Ensure Excel is open.
- Date parsing: the script expects date/time strings in the format `m/d/yy H:MM` (see `pd.to_datetime(..., format='%-m/-%d/%y %-I:%M')`). If your locale differs, adjust the format or remove the `format=` parameter to allow flexible parsing.

Development notes
-----------------
- The script was refactored to use vectorized `numpy.select` / boolean masks instead of row-wise `apply` for better performance on large tables. Hince the use of vec(tor) in the naming convention.
- Helper logic maps visitor badge IDs to `ContactNumber`/`Name` using a dictionary lookup created from the security table.

Extending the script
--------------------
- Add unit tests for the classification and mapping functions.
- Improve Excel integration: add a small CLI to support a `--workbook` argument instead of hard-coding the workbook name.
