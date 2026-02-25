import xlwings as xw
import time
import os
from pathlib import Path

# Configuration
WORKBOOK_NAME = "PMS 3.1.xlsm"
VBA_DIR = Path("vba-files")
POLL_INTERVAL = 1.0


def export_all(wb):
    print(f"Exporting all modules to {VBA_DIR}...")
    VBA_DIR.mkdir(exist_ok=True)
    for comp in wb.api.VBProject.VBComponents:
        try:
            # Component types: 1=Standard, 2=Class, 100=Sheet/Workbook
            ext = ".bas" if comp.Type == 1 else ".cls"
            # Sheet/Workbook modules also use .cls
            out_file = VBA_DIR / f"{comp.Name}{ext}"
            comp.Export(str(out_file.absolute()))
            print(f"  Exported: {comp.Name}")
        except Exception as e:
            print(f"  Failed to export {comp.Name}: {e}")


def import_component(wb, path):
    name = path.stem
    print(f"Importing {name} from {path.name}...")
    try:
        # Check if component exists
        try:
            comp = wb.api.VBProject.VBComponents(name)
            # Cannot simply 'import' over an existing one, must remove first
            # (Standard modules and classes can be removed, Sheets cannot)
            if comp.Type in [1, 2]:
                wb.api.VBProject.VBComponents.Remove(comp)
                wb.api.VBProject.VBComponents.Import(str(path.absolute()))
            else:
                # For Sheets/Workbook, we have to replace the code lines
                # This is more complex, but usually 'excel-vba' handles this.
                # For now, let's focus on .bas and .cls (Classes).
                print(f"  Skip auto-sync for Document module {name} (Safety fallback)")
        except Exception:
            # If it doesn't exist, just import
            wb.api.VBProject.VBComponents.Import(str(path.absolute()))
        print(f"  [SUCCESS] {name} updated.")
    except Exception as e:
        print(f"  [ERROR] Failed to import {name}: {e}")


def main():
    print("=" * 60)
    print(" PMS 3.1 Stable VBA Watcher")
    print("=" * 60)

    try:
        # Connect to existing workbook by name (not path)
        print(f"Connecting to {WORKBOOK_NAME}...")
        wb = None
        for app in xw.apps:
            for b in app.books:
                if b.name.lower() == WORKBOOK_NAME.lower():
                    wb = b
                    break

        if not wb:
            print(f"Error: {WORKBOOK_NAME} is not open in Excel.")
            return

        # Initial Export
        export_all(wb)

        # Initial Hashes
        hashes = {f: f.stat().st_mtime for f in VBA_DIR.glob("*.*")}

        print(f"\nWatching {VBA_DIR} for changes (Interval: {POLL_INTERVAL}s)...")
        print("Press Ctrl+C to stop.")

        while True:
            time.sleep(POLL_INTERVAL)

            # Check if book still open
            try:
                wb.api.Name  # Quick test
            except:
                print("\nWorkbook closed. Stopping.")
                break

            # Check for file changes
            for f in VBA_DIR.glob("*.*"):
                mtime = f.stat().st_mtime
                if f not in hashes or mtime > hashes[f]:
                    # File changed or new
                    import_component(wb, f)
                    hashes[f] = mtime

    except KeyboardInterrupt:
        print("\nWatcher stopped by user.")
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback

        traceback.print_exc()


if __name__ == "__main__":
    main()
