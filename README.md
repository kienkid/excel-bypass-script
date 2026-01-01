# Excel Bypass Script
Script to remove VBA project protect password, unprotect worksheet and workbook. For Excel file encryption, the approach is brute force which not include in this script

## Prerequisite
You need to install Python before using the script

## Getting Started
1. **Unprotect workbook/worksheet**:
  - Locate your destination file path
```bash
python unprotect-excel.py <file_path>
```
  - Unprotected file will be created within the folder contain the source file 

2. **Remove VBA project protect password**:
  - Locate your destination file path
```bash
python crack-VBA-protect.py <file_path>
```
  - Cracked file will be created within the folder contain the source file
  - Open the cracked excel file > Alt + F11 to open the Visual Basic Editor > Right-click the VBA project > VBAProject properties > Update the password to something you know
  - Some error message dialog will pop up along the process > Just click ok
  - Now you the VBA project protect password will be the one you set
