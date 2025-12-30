import zipfile
import shutil
import os

OLD_BYTES = b"DPB"
NEW_BYTES = b"DPx"

input_excel = input("Please enter your file path for the excel file: ").replace('"', '')
file_root, file_ext = os.path.splitext(input_excel)
output_excel = file_root + " (cracked)" + file_ext

def main():    
    if ".xlsm" in input_excel:
        TEMP_ZIP = "temp.zip"
        TEMP_DIR = "temp_excel"
        
        VBA_ORIGINAL = "vbaProject.bin"
        VBA_PATCHED = "vbaProject_patched.bin"
        
        shutil.copy(input_excel, TEMP_ZIP)
        
        with zipfile.ZipFile(TEMP_ZIP, "r") as z:
            z.extractall(TEMP_DIR)
        
        vba_path = os.path.join(TEMP_DIR, "xl", VBA_ORIGINAL)
        patched_path = os.path.join(TEMP_DIR, "xl", VBA_PATCHED)
        
        if not os.path.exists(vba_path):
            raise FileNotFoundError("vbaProject.bin not found (is this really an .xlsm?)")
        
        shutil.copy(vba_path, patched_path)
        
        hexPatching(patched_path)

        shutil.move(patched_path, vba_path)
        
        with zipfile.ZipFile(output_excel, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(TEMP_DIR):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, TEMP_DIR)
                    z.write(full_path, arcname)
        
        shutil.rmtree(TEMP_DIR)
        os.remove(TEMP_ZIP)
        
        print("[+] Done. Output saved as:", output_excel)
        print("[+] Open the cracked file and change the VBA password to your password. Your cracked file should usable right away")

    elif ".xls" in input_excel:
        shutil.copy(input_excel, output_excel)
        hexPatching(output_excel)
        print("[+] Done. Output saved as:", output_excel)
        print("[+] Open the cracked file and change the VBA password to your password. Your cracked file should usable right away")
    
    else:
        print("[-] Your file path is not recognize as excel file (.xls, .xlsm, .xlsx,...)")
        print("[-] Exiting the program...")

def hexPatching(binary_file_path):
    with open(binary_file_path, "rb") as f:
        data = f.read()

    count = data.count(OLD_BYTES)
    if count == 0:
        raise ValueError("Password protect pattern not found in vbaProject.bin")

    data = data.replace(OLD_BYTES, NEW_BYTES)
    
    with open(binary_file_path, "wb") as f:
        f.write(data)
    
    print(f"[{count}] Patched successful, creating the cracked file...")



if __name__ == "__main__":
    main()