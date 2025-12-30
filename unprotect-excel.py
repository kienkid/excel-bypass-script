import zipfile
import shutil
import os
import re

WORKBOOK_PATTERN = re.compile(r"<workbookProtection\b[^>]*?\s*/>", flags=re.IGNORECASE)
WORKSHEET_PATTERN = re.compile(r"<sheetProtection\b[^>]*?\s*/>", flags=re.IGNORECASE)

input_excel = input("Please enter your file path for the excel file: ").replace('"', '')
file_root, file_ext = os.path.splitext(input_excel)
output_excel = file_root + " (unprotect)" + file_ext

def main():    
    if input_excel.lower().endswith((".xlsm", ".xlsx")):
        TEMP_ZIP = "temp.zip"
        TEMP_DIR = "temp_excel"
        
        shutil.copy(input_excel, TEMP_ZIP)
        
        with zipfile.ZipFile(TEMP_ZIP, "r") as z:
            z.extractall(TEMP_DIR)
        
        workbook_path = os.path.join(TEMP_DIR, "xl", "workbook.xml")
        patched_path = os.path.join(TEMP_DIR, "xl", "workbook_patched.xml")
        sheets_path = []
        for root, dirs, files in os.walk(os.path.join(TEMP_DIR, "xl", "worksheets")):
            for file in files:
                full_path = os.path.join(root, file)
                sheets_path.append(full_path)
        
        if not os.path.exists(workbook_path):
            raise FileNotFoundError("workbook.xml not found (is this really an .xlsx?)")
        
        shutil.copy(workbook_path, patched_path)
        
        workbookPatching(patched_path)

        shutil.move(patched_path, workbook_path)

        for sheet in sheets_path:
            dir_name = os.path.dirname(sheet)
            sheet_pathed_path = os.path.join(dir_name, f"cache_{os.path.basename(sheet)}")

            shutil.copy(sheet, sheet_pathed_path)
            sheetPatching(sheet_pathed_path)
            shutil.move(sheet_pathed_path, sheet)
            print(f"[+] Unprotect worksheet {sheet} successful, progressing...")
        
        with zipfile.ZipFile(output_excel, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(TEMP_DIR):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, TEMP_DIR)
                    z.write(full_path, arcname)
        
        shutil.rmtree(TEMP_DIR)
        os.remove(TEMP_ZIP)
        
        print("[+] Done. Output saved as:", output_excel)

    else:
        print("[-] Your file path is not recognize as excel file (.xlsm, .xlsx,...)")
        print("[-] Exiting the program...")

def workbookPatching(workbook_xml):
    
    with open(workbook_xml, "r", encoding="utf-8") as f:
        content = f.read()  

    new_content, count = WORKBOOK_PATTERN.subn("", content)

    if count > 0:
        with open(workbook_xml, "w", encoding="utf-8", newline="") as f:
            f.write(new_content)
    
    print(f"[{count}] Unprotect workbook successful, progressing...")

def sheetPatching(sheet_xml):
    
    with open(sheet_xml, "r", encoding="utf-8") as f:
        content = f.read()

    new_content, count = WORKSHEET_PATTERN.subn("", content)

    if count > 0:
        with open(sheet_xml, "w", encoding="utf-8", newline="") as f:
            f.write(new_content)



if __name__ == "__main__":
    main()