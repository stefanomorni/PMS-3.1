import os


def strip_bom_and_fix(folder):
    if not os.path.exists(folder):
        print(f"Directory not found: {folder}")
        return

    for root, dirs, files in os.walk(folder):
        for name in files:
            if name.endswith((".bas", ".cls", ".frm")):
                path = os.path.join(root, name)
                content = None

                # Try UTF-8 (with or without BOM)
                try:
                    with open(path, "r", encoding="utf-8-sig") as f:
                        content = f.read()
                except UnicodeDecodeError:
                    # Try CP1252 (Western European)
                    try:
                        with open(path, "r", encoding="cp1252") as f:
                            content = f.read()
                        print(f"Read as CP1252: {path}")
                    except Exception as e:
                        print(f"Fatal error reading {path}: {e}")
                        continue

                if content is not None:
                    try:
                        # Write back as utf-8 (no BOM) with CRLF
                        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
                            f.write(content)
                        print(f"Cleaned (UTF-8 No BOM): {path}")
                    except Exception as e:
                        print(f"Error writing {path}: {e}")


projects = [
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORFunctions\vba-files",
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\vba-files",
    r"D:\Cloud\OneDrive\MC\Investimenti\PMS\PMS 3.1\vba-files",
]

for p in projects:
    strip_bom_and_fix(p)
