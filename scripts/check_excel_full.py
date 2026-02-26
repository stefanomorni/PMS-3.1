import win32com.client
import os


def check_excel_full():
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        print(f"Excel Instance: {xl.Caption}")

        print("\n--- Workbooks ---")
        for wb in xl.Workbooks:
            print(f"Name: {wb.Name}")
            print(f"  Path: {wb.FullName}")
            print(
                f"  Visible: {wb.Windows(1).Visible if wb.Windows.Count > 0 else 'No Window'}"
            )

        print("\n--- Add-Ins (Installed/Loaded) ---")
        for ai in xl.AddIns:
            if ai.Installed or "MOR" in ai.Name:
                print(
                    f"Name: {ai.Name} (Installed: {ai.Installed}, Path: {ai.FullName})"
                )

        # Specific paths for our projects
        MOR_FUNCTIONS = r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORFunctions\MORFunctions.xlam"
        MOR_PROCEDURES = r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\MORProcedures.xlam"

        # Check if we can see them by name
        names = [wb.Name.lower() for wb in xl.Workbooks]
        for label, path in [
            ("MORFunctions", MOR_FUNCTIONS),
            ("MORProcedures", MOR_PROCEDURES),
        ]:
            basename = os.path.basename(path).lower()
            if basename not in names:
                print(
                    f"\n[ACTION] {label} not found in Workbooks. Attempting to open from: {path}"
                )
                try:
                    xl.Workbooks.Open(path)
                    print(f"  Successfully opened {label}")
                except Exception as open_err:
                    print(f"  Failed to open {label}: {open_err}")
            else:
                print(f"\n{label} already present in Workbooks.")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    check_excel_full()
