# Excel Bypass Script
Script to remove VBA project protect password, unprotect worksheet and workbook. For Excel file encryption, the approach is brute force

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
