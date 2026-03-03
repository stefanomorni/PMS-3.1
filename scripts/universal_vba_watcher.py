import time
import sys
import os
from pathlib import Path
import win32com.client

# ================================================================== #
#  Universal Office VBA Watcher (Win32COM - Word/PPT/Excel/Access/Outlook)
# ================================================================== #
# Purpose: Robust VBA synchronization for OneDrive-synced projects.
#
# Usage:
#   python scripts/universal_vba_watcher.py [filename] [app_type] [flag]
#
# app_type: excel (default), word, powerpoint, access, outlook
# flag:
#   --export-only: Export code from Office app to VS Code, then exit.
#   --import-all:  Force overwrite code in Office app with VS Code files.
#   (none):        Continuous real-time sync (save in VS Code -> sync to Office).
# ================================================================== #


class UniversalVBAWatcher:
    def __init__(self, target_name=None, app_type="excel", poll_interval=1.0):
        self.target_name = target_name
        self.app_type = app_type.lower()
        self.poll_interval = poll_interval
        self.vba_dir = Path("vba-files")
        self.hashes = {}
        self.doc = None

    def connect(self):
        print(
            f"Connecting to {self.app_type.upper()} for: {self.target_name or 'Active Document'}..."
        )
        try:
            if self.app_type == "excel":
                app = win32com.client.GetActiveObject("Excel.Application")
                docs = app.Workbooks
            elif self.app_type == "word":
                app = win32com.client.GetActiveObject("Word.Application")
                docs = app.Documents
            elif self.app_type == "powerpoint":
                app = win32com.client.GetActiveObject("PowerPoint.Application")
                docs = app.Presentations
            elif self.app_type == "access":
                app = win32com.client.GetActiveObject("Access.Application")
                # Access projects are singletons per app instance usually
                self.doc = app.CurrentProject
                print(f"  [CONNECTED] Access Project: {app.CurrentProject.Name}")
                return True
            elif self.app_type == "outlook":
                app = win32com.client.GetActiveObject("Outlook.Application")
                # Outlook has one VBAProject.OTM
                self.doc = app.Session.Application  # Accessing through session
                print(f"  [CONNECTED] Outlook Session")
                return True
            else:
                raise ValueError(f"Unsupported app type: {self.app_type}")

            if self.app_type in ["excel", "word", "powerpoint"]:
                for d in docs:
                    if (
                        not self.target_name
                        or d.Name.lower() == self.target_name.lower()
                    ):
                        self.doc = d
                        self.target_name = d.Name
                        print(f"  [CONNECTED] {d.Name}")
                        return True
        except Exception as e:
            print(f"  [CONNECTION ERROR] {e}")
        return False

    def export_all(self):
        print(f"Exporting modules to {self.vba_dir}...")
        self.vba_dir.mkdir(exist_ok=True)
        # Accessing VBProject (requires 'Trust Access to VBA Project Object Model' enabled)
        for comp in self.doc.VBProject.VBComponents:
            try:
                # 1=Standard, 2=Class, 3=Form, 100=Document
                ext = ".bas" if comp.Type == 1 else ".cls"
                if comp.Type == 3:
                    ext = ".frm"

                out_file = self.vba_dir / f"{comp.Name}{ext}"
                comp.Export(str(out_file.absolute()))
                print(f"    Exported: {comp.Name}")
            except Exception as e:
                print(f"    Failed {comp.Name}: {e}")

    def sync_file(self, path):
        name = path.stem
        print(f"Syncing: {path.name} -> {self.app_type.upper()}...")
        try:
            try:
                comp = self.doc.VBProject.VBComponents(name)
                if comp.Type in [1, 2, 3]:  # Standard, Class, Form
                    self.doc.VBProject.VBComponents.Remove(comp)
                    self.doc.VBProject.VBComponents.Import(str(path.absolute()))
                else:
                    print(
                        f"    [NOTE] {name} is a Document module (Sheet/ThisWorkbook). Auto-sync limited."
                    )
            except Exception:
                self.doc.VBProject.VBComponents.Import(str(path.absolute()))
            print(f"    [SUCCESS] {name} updated.")
        except Exception as e:
            print(f"    [ERROR] {e}")

    def run(self):
        if not self.connect():
            print(
                "Error: Could not find target document. Ensure it is open in the application."
            )
            return

        self.export_all()
        self.hashes = {
            f: f.stat().st_mtime
            for f in self.vba_dir.glob("*.*")
            if f.suffix in [".bas", ".cls", ".frm"]
        }

        print(f"\nWatching {self.vba_dir} (Heartbeat: {self.poll_interval}s)...")
        try:
            while True:
                time.sleep(self.poll_interval)
                # Heartbeat check
                try:
                    _ = self.doc.Name
                except:
                    print("\nConnection lost. Document closed?")
                    break

                for f in self.vba_dir.glob("*.*"):
                    if f.suffix not in [".bas", ".cls", ".frm"]:
                        continue
                    mtime = f.stat().st_mtime
                    if f not in self.hashes or mtime > self.hashes[f]:
                        self.sync_file(f)
                        self.hashes[f] = mtime
        except KeyboardInterrupt:
            print("\nStopped.")


if __name__ == "__main__":
    # 0. Load Configuration
    try:
        import project_config

        config = project_config.load_config()
    except ImportError:
        config = {}

    t_name = config.get("file") or None
    if t_name:
        t_name = Path(t_name).name

    t_type = config.get("app_type", "excel").lower()
    export_only = False
    import_all = False

    # 1. Override with Command Line Arguments
    for arg in sys.argv[1:]:
        if arg == "--export-only":
            export_only = True
        elif arg == "--import-all":
            import_all = True
        elif arg.lower() in ["excel", "word", "powerpoint", "access", "outlook"]:
            t_type = arg.lower()
        elif not arg.startswith("-"):
            t_name = arg

    watcher = UniversalVBAWatcher(target_name=t_name, app_type=t_type)

    if export_only:
        if watcher.connect():
            watcher.export_all()
        else:
            print("Error: Could not find document for export.")
            sys.exit(1)
    elif import_all:
        if watcher.connect():
            print(f"Force-importing all components from {watcher.vba_dir}...")
            for f in watcher.vba_dir.glob("*.*"):
                if f.suffix in [".bas", ".cls", ".frm"]:
                    watcher.sync_file(f)
            print("Bulk import complete.")
        else:
            print("Error: Could not find document for import.")
            sys.exit(1)
    else:
        watcher.run()
