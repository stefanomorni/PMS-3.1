import win32com.client
import sys
import time


def debug_excel():
    print("Attempting to connect to Excel via COM...")
    try:
        # Connect to existing Excel instance
        xl = win32com.client.GetActiveObject("Excel.Application")
        print("Excel COM is responding.")

        print(f"Excel Version: {xl.Version}")
        print(f"Calculation State: {xl.Calculation}")
        print(f"EnableEvents: {xl.EnableEvents}")
        print(f"ScreenUpdating: {xl.ScreenUpdating}")

        print("\nOpen Workbooks:")
        for wb in xl.Workbooks:
            print(
                f"- {wb.Name} (VBA Protected: {wb.VBProject.Protection == 1 if hasattr(wb, 'VBProject') else 'Unknown'})"
            )

        # Try to see if it's busy
        try:
            print("\nAttempting to call a simple Excel function...")
            res = xl.WorksheetFunction.Sum(1, 2)
            print(f"WorksheetFunction.Sum(1, 2) = {res}")
        except Exception as e:
            print(f"Excel Busy/Error: {e}")

    except Exception as e:
        print(f"Failed to connect or Excel is completely hung: {e}")


if __name__ == "__main__":
    debug_excel()
