import win32com.client
import sys


def check_excel():
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        print(f"Excel is running: {xl.Caption}")
        print(f"Interactive mode: {xl.Interactive}")
        print(f"Calculation state: {xl.Calculation}")
        print(f"Screen updating: {xl.ScreenUpdating}")

        print("\nOpen Workbooks:")
        for wb in xl.Workbooks:
            print(f"- {wb.Name} (Path: {wb.FullName})")

    except Exception as e:
        print(f"Error accessing Excel via COM: {e}")
        print(
            "This usually means Excel is busy, in edit mode (cursor in cell), or a modal dialog is open."
        )


if __name__ == "__main__":
    check_excel()
