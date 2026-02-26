import os


def fix_encoding(path):
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    # Explicitly write as utf-8 WITHOUT BOM
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(content)
    print(f"Fixed encoding: {path}")


files = [
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORFunctions\vba-files\Informational.bas",
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\vba-files\FrmFinestraInformativa.frm",
    r"D:\Cloud\OneDrive\MC\Investimenti\PMS\PMS 3.1\vba-files\ThisWorkbook.cls",
]

for f in files:
    if os.path.exists(f):
        fix_encoding(f)
    else:
        print(f"File not found: {f}")
