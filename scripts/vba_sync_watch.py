import xlwings as xw
import time
import sys
import os
import argparse
from pathlib import Path

# ================================================================== #
#  Universal Office VBA Watcher (OneDrive Resistant)                  #
# ================================================================== #
# Version: 1.1 (2026-02-28)
# Strategy: Name-based correlation (immune to FullPath URL flip)
# ================================================================== #


class OfficeVBAWatcher:
    def __init__(
        self, target_name=None, vba_dir="vba-files", app_type="excel", poll_interval=1.0
    ):
        self.target_name = target_name
        self.app_type = app_type.lower()
        self.poll_interval = poll_interval
        self.vba_dir = Path(vba_dir)
        self.hashes = {}
        self.app = None
        self.wb = None

    def connect(self):
        """Connect to the open document/workbook by Name."""
        print(
            f"Connecting to {self.app_type.capitalize()} for: {self.target_name or 'Active Book'}..."
        )

        # xlwings.apps works only for Excel.
        for app in xw.apps:
            for b in app.books:
                if not self.target_name or b.name.lower() == self.target_name.lower():
                    self.wb = b
                    self.target_name = b.name
                    print(f"  [CONNECTED] {b.fullname}")
                    return True
        return False

    def export_all(self):
        print(f"Exporting modules to {self.vba_dir}...")
        self.vba_dir.mkdir(exist_ok=True)
        for comp in self.wb.api.VBProject.VBComponents:
            try:
                # 1=Standard, 2=Class, 3=Form, 100=Document/Sheet
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
        print(f"Syncing: {path.name} -> {self.app_type.capitalize()}...")
        try:
            try:
                comp = self.wb.api.VBProject.VBComponents(name)
                if comp.Type in [1, 2, 3]:  # Std, Class, Form
                    self.wb.api.VBProject.VBComponents.Remove(comp)
                    self.wb.api.VBProject.VBComponents.Import(str(path.absolute()))
                else:
                    print(
                        f"    [SKIP] {name} is a Document module (requires manual sync)."
                    )
            except Exception:
                self.wb.api.VBProject.VBComponents.Import(str(path.absolute()))
            print(f"    [SUCCESS] {name} updated.")
        except Exception as e:
            print(f"    [ERROR] {e}")

    def run(self, force_import=False):
        if not self.connect():
            print(
                f"Error: Could not find {self.target_name if self.target_name else 'any open book'}."
            )
            return

        if force_import:
            print(f"Force-importing all components from {self.vba_dir}...")
            for f in self.vba_dir.glob("*.*"):
                if f.suffix in [".bas", ".cls", ".frm"]:
                    self.sync_file(f)
            print("Bulk import complete.")
        else:
            self.export_all()

        self.hashes = {f: f.stat().st_mtime for f in self.vba_dir.glob("*.*")}

        print(f"\nWatching {self.vba_dir} (Heartbeat: {self.poll_interval}s)...")
        try:
            while True:
                time.sleep(self.poll_interval)
                # Heartbeat check
                try:
                    _ = self.wb.api.Name
                except:
                    print("\nConnection lost (Document closed?). Stopping.")
                    break

                for f in self.vba_dir.glob("*.*"):
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

    default_project = config.get("file") or None
    if default_project and "\\" in default_project:
        default_project = Path(default_project).name

    parser = argparse.ArgumentParser(description="Office VBA Sync Watcher")
    parser.add_argument(
        "project",
        nargs="?",
        default=default_project,
        help="Project name (workbook name)",
    )
    parser.add_argument(
        "--dir",
        default=config.get("vba_directory", "vba-files"),
        help="VBA source directory",
    )
    parser.add_argument(
        "--import-all", action="store_true", help="Import all files on startup"
    )
    args = parser.parse_args()

    watcher = OfficeVBAWatcher(target_name=args.project, vba_dir=args.dir)
    watcher.run(force_import=args.import_all)
