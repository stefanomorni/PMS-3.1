import win32com.client


def verify_document_module():
    print("Verifying Commissioni document module in MORFunctions.xlam...")
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        wb = xl.Workbooks("MORFunctions.xlam")
        comp = wb.VBProject.VBComponents("Commissioni")
        cm = comp.CodeModule

        count = cm.CountOfLines
        print(f"Total lines in VBE: {count}")

        if count > 0:
            print("FIRST 10 LINES IN VBE:")
            print("---")
            # cm.Lines(Start, Count)
            to_read = min(10, count)
            print(cm.Lines(1, to_read))
            print("---")
        else:
            print("Module is empty!")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    verify_document_module()
