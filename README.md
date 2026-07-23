# VBA-Automations
# Excel VBA Code Modules

This repository contains standalone VBA code modules for Excel, designed to automate data processing and reporting tasks.

## Overview

The repo includes two VBA code files:

1. **SQL Stored Procedure Executor (`VBA Automation.bas`)**  
   A VBA module that executes a SQL stored procedure and writes the returned data into specific worksheet cells.

2. **Report Splitter (`VBA Split Data into multiple sheets.bas`)**  
   A VBA module that filters or splits a report into multiple sheets within the same workbook.

---

## Usage Instructions

### Importing VBA Modules into Your Workbook

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, go to `File` > `Import File...`.
4. Select the desired `.bas` file from this repository to import the code.
5. Once imported, you can run the macros or call the functions as per your needs.

### Prerequisites

- Enable macros in your Excel environment.
- For the SQL Stored Procedure Executor:
  - You need access to a SQL database.
  - Update the connection string and stored procedure parameters inside the VBA code before running.
- For the Report Splitter:
  - Ensure your report data is structured properly on a worksheet.

---

## Description of Modules

### `VBA Automation.bas`

- Connects to a SQL database.
- Runs a stored procedure.
- Populates designated cells with the fetched data.

*Note:* Modify the database connection details and cell references within the module.

### `VBA Split Data into multiple sheets.bas`

- Splits or filters a report by specified criteria.
- Creates new worksheets with filtered subsets of the data.

---

## Notes

- Always backup your workbook before importing or running VBA code.
- Adjust the code as necessary to fit your workbook structure and data.
- These modules are designed as starting points and may require customization.

---


## Support

Feel free to open issues for bug reports or feature requests.

![VBA](https://img.shields.io/badge/Language-VBA-blue)
