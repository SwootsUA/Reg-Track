# Registry Software Audit Tool

This Python script is a Windows registry auditing tool that extracts software information from various registry keys and tracks installed applications over time. It saves the extracted data into an Excel file named with the current date, moves older Excel files into an `old` folder, and compares registry snapshots to identify changes.

## Features

-   **Registry Extraction:**  
    Reads multiple registry keys (e.g., Uninstall keys and MSI keys) from both `HKEY_CURRENT_USER` and `HKEY_LOCAL_MACHINE`.

-   **Candidate Value Handling:**  
    Supports reading multiple possible value names such as `DisplayName`, `ProductName`, and nested keys like `InstallProperties\DisplayName`.

-   **Excel Reporting:**  
    Stores the registry data in a date-named Excel file and adjusts the column widths for readability.

-   **Historical Data Management:**  
    Automatically moves older Excel files to an `old` folder and retains the most recent data for comparison.

-   **Snapshot Comparison:**  
    Compares two Excel files (expected to be in `YYYY-MM-DD.xlsx` format) to report added or removed software entries and column changes.

## Requirements

-   **Operating System:**  
    Windows (for registry access).

-   **Python Version:**  
    Python 3.x

-   **Dependencies:**
    -   [openpyxl](https://openpyxl.readthedocs.io/)
    -   [pandas](https://pandas.pydata.org/)

## Usage

1. Configure Paths and Registry Keys:
   Modify the file paths and registry keys in the script as necessary.

2. Run the Script:
    Open `start.bat` file.

3. Interact:
    The script will perform its operations, give you analisys on what have changed from the last time and prompt you to press Enter before exiting.


