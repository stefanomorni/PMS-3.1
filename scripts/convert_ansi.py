import os


def convert_to_ansi(folder):
    if not os.path.exists(folder):
        return

    for root, dirs, files in os.walk(folder):
        for name in files:
            if name.endswith((".bas", ".cls", ".frm")):
                path = os.path.join(root, name)
                # Read as UTF-8 (stripping BOM if present)
                try:
                    with open(path, "r", encoding="utf-8-sig") as f:
                        content = f.read()

                    # Write as Windows-1252 (ANSI)
                    with open(
                        path,
                        "w",
                        encoding="windows-1252",
                        errors="replace",
                        newline="\r\n",
                    ) as f:
                        f.write(content)
                    print(f"Converted to ANSI: {path}")
                except Exception as e:
                    print(f"Error converting {path}: {e}")


projects = [
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORFunctions\vba-files",
    r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\vba-files",
    r"D:\Cloud\OneDrive\MC\Investimenti\PMS\PMS 3.1\vba-files",
]

for p in projects:
    convert_to_ansi(p)
