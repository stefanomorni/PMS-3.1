import xlwings as xw
import time
import sys
import os
import argparse
from pathlib import Path

# ================================================================== #
#  Universal Office VBA Watcher (v2.0 - AddFromString + Meta-Strip)   #
# ================================================================== #
# Strategy: AddFromString for all code-only modules to prevent
#           encoding mismatches while stripping illegal Attributes.
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
        self.wb = None

    def connect(self):
        """Connect to the open document/workbook by Name."""
        print(f"Connecting to {self.app_type.capitalize()} for: {self.target_name}...")

        # Try exact name first
        for app in xw.apps:
            try:
                xl = app.api
                possible_names = [self.target_name]
                if "." not in self.target_name:
                    possible_names.extend(
                        [
                            f"{self.target_name}.xlsm",
                            f"{self.target_name}.xlam",
                        ]
                    )

                for name in possible_names:
                    try:
                        wb_api = xl.Workbooks(name)
                        self.wb = app.books[wb_api.Name]
                        print(f"  [CONNECTED] {wb_api.FullName}")
                        return True
                    except:
                        continue
            except:
                continue

        # Fallback: substring match
        for app in xw.apps:
            for b in app.books:
                if not self.target_name or self.target_name.lower() in b.name.lower():
                    self.wb = b
                    self.target_name = b.name
                    print(f"  [CONNECTED] {b.fullname}")
                    return True
        return False

    def reset_vbe(self):
        """Try to reset the VBE (Stop debugging)."""
        try:
            self.wb.api.VBProject.VBE.CommandBars.FindControl(Id=228).Execute()
        except:
            pass

    def strip_metadata(self, lines):
        """Strip VBE headers and embedded Attribute lines."""
        clean = []
        in_begin = False
        header_done = False

        for line in lines:
            stripped = line.strip()

            # Always skip Attribute lines (prevents syntax errors in AddFromString)
            if stripped.startswith("Attribute "):
                continue

            if not header_done:
                if stripped.startswith("VERSION ") or stripped.startswith("BEGIN"):
                    if stripped.startswith("BEGIN"):
                        in_begin = True
                    continue
                if in_begin:
                    if stripped == "END":
                        in_begin = False
                    continue
                if stripped:
                    header_done = True
                else:
                    continue

            if header_done:
                clean.append(line)

        return "".join(clean).strip()

    def sync_file(self, path):
        name = path.stem
        ext = path.suffix.lower()
        print(f"Syncing: {path.name} -> {self.wb.name}...")

        try:
            self.reset_vbe()
            project = self.wb.api.VBProject

            # 1. Identify Existing Component
            try:
                comp = project.VBComponents(name)
                comp_type = comp.Type
            except:
                comp = None
                comp_type = None

            # 2. Forms stick to .Import() as they have .frx binary dependencies
            if ext == ".frm":
                if comp:
                    project.VBComponents.Remove(comp)
                project.VBComponents.Import(str(path.absolute()))
                print(f"    [SUCCESS] {name} imported (via Import).")
                return

            # 3. Code Modules (.bas, .cls, and Document modules)
            # Read source
            try:
                with open(path, "r", encoding="utf-8-sig") as f:
                    lines = f.readlines()
            except UnicodeDecodeError:
                with open(path, "r", encoding="windows-1252") as f:
                    lines = f.readlines()

            code = self.strip_metadata(lines)

            if not comp or comp_type in [1, 2]:
                # Non-document modules: Remove and Re-add fresh
                if comp:
                    project.VBComponents.Remove(comp)

                vba_type = 1 if ext == ".bas" else 2
                new_comp = project.VBComponents.Add(vba_type)
                new_comp.Name = name
                comp = new_comp

            # Inject code
            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            if code:
                cm.AddFromString(code)

            print(f"    [SUCCESS] {name} updated via AddFromString.")

        except Exception as e:
            print(f"    [ERROR] {e}")

    def run(self, force_import=False):
        if not self.connect():
            print(f"Could not connect to {self.target_name}")
            return

        if not self.vba_dir.exists():
            print(f"Error: VBA directory not found at {self.vba_dir}")
            return

        files = list(self.vba_dir.glob("*.*"))
        self.hashes = {f: f.stat().st_mtime for f in files}

        if force_import:
            print(f"Force importing all files from {self.vba_dir}...")
            for f in files:
                if f.suffix.lower() in [".bas", ".cls", ".frm"]:
                    self.sync_file(f)

        print(f"\nWatching {self.vba_dir}...")
        try:
            while True:
                time.sleep(self.poll_interval)
                try:
                    _ = self.wb.api.Name
                except:
                    break
                for f in self.vba_dir.glob("*.*"):
                    if f.suffix.lower() not in [".bas", ".cls", ".frm"]:
                        continue
                    mtime = f.stat().st_mtime
                    if f not in self.hashes or mtime > self.hashes[f]:
                        self.sync_file(f)
                        self.hashes[f] = mtime
        except KeyboardInterrupt:
            print("\nStopped.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Office VBA Sync Watcher")
    parser.add_argument(
        "project", nargs="?", default="PMS 3.1", help="Project name (workbook name)"
    )
    parser.add_argument("--dir", default="vba-files", help="VBA source directory")
    parser.add_argument(
        "--import-all", action="store_true", help="Import all files on startup"
    )
    args = parser.parse_args()

    watcher = OfficeVBAWatcher(target_name=args.project, vba_dir=args.dir)
    watcher.run(force_import=args.import_all)
