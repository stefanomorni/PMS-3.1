import win32com.client
import os
import sys
from pathlib import Path

# VBE Component Types
VBA_MODULE_STANDARD = 1  # .bas
VBA_MODULE_CLASS = 2  # .cls
VBA_MODULE_FORM = 3  # .frm
VBA_MODULE_DOCUMENT = 100  # Sheet, ThisWorkbook, etc.

# Map file extension to new component type
EXT_TO_TYPE = {
    ".bas": VBA_MODULE_STANDARD,
    ".cls": VBA_MODULE_CLASS,
    ".frm": VBA_MODULE_FORM,
}


def read_source(f: Path) -> list[str]:
    """Read a VBA source file, trying UTF-8-sig first then windows-1252."""
    try:
        return f.read_text(encoding="utf-8-sig").splitlines(keepends=True)
    except UnicodeDecodeError:
        return f.read_text(encoding="windows-1252").splitlines(keepends=True)


def strip_vbe_headers(lines: list[str]) -> str:
    """
    Strip VBE file headers (VERSION, Attribute, BEGIN/END blocks) and return
    the remaining code as a single string. Used for both Document and
    code-module injection so that AddFromString receives only clean code.
    """
    clean = []
    in_begin = False
    header_done = False

    for line in lines:
        stripped = line.strip()

        if not header_done:
            if (
                stripped.startswith("VERSION ")
                or stripped.startswith("Attribute ")
                or stripped.startswith("BEGIN")
            ):
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


def raw_import_v4(project_target_name: str, source_dir: str):
    """
    Import all VBA modules from *source_dir* into the matching VBProject.

    All injection is done via CodeModule.AddFromString so that Python's
    Unicode strings flow through COM directly into VBA — no intermediate
    ANSI file is involved, which permanently solves the encoding issue with
    accented characters on Italian/non-English Windows systems.
    """
    print(f"--- Starting Raw Import v4 for {project_target_name} ---")
    print(f"Source: {source_dir}")

    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        print(f"Connected to Excel: {xl.Caption}")
    except Exception as e:
        print(f"FAILED to connect to Excel: {e}")
        return

    # --- Locate the target VBProject ---
    project = None
    print(f"Searching for VBProject matching: {project_target_name}")
    try:
        for p in xl.VBE.VBProjects:
            name, filename, base = "", "", ""
            try:
                name = p.Name
            except:
                pass
            try:
                filename = os.path.basename(p.FileName)
            except:
                pass
            base = os.path.splitext(filename)[0]

            print(f"  Checking: '{name}' (File: '{filename}', Base: '{base}')")

            if project_target_name.lower() in [
                name.lower(),
                filename.lower(),
                base.lower(),
            ] or filename.lower().startswith(project_target_name.lower()):
                project = p
                print(f"  [MATCH FOUND] {name}")
                break
    except Exception as e:
        print(f"  Error accessing VBE: {e}")
        return

    if not project:
        print(f"FAILED: Could not find VBProject '{project_target_name}'")
        return

    vba_dir = Path(source_dir)
    if not vba_dir.exists():
        print(f"FAILED: Source directory does not exist: {source_dir}")
        return

    # --- Process each source file ---
    for f in sorted(vba_dir.glob("*.*")):
        if f.suffix.lower() not in EXT_TO_TYPE:
            continue

        module_name = f.stem
        desired_type = EXT_TO_TYPE[f.suffix.lower()]
        print(f"Processing: {module_name}...")

        try:
            lines = read_source(f)
            code = strip_vbe_headers(lines)

            # Check for an existing component
            comp = None
            try:
                comp = project.VBComponents(module_name)
            except:
                pass

            if comp is not None and comp.Type == VBA_MODULE_DOCUMENT:
                # --- Document module (Sheet, ThisWorkbook): inject only ---
                print(f"    Updating Document module {module_name} (AddFromString)...")
                cm = comp.CodeModule
                if cm.CountOfLines > 0:
                    cm.DeleteLines(1, cm.CountOfLines)
                if code:
                    cm.AddFromString(code)
                print(f"    [SUCCESS] {module_name} (Document) updated.")

            else:
                # --- Standard / Class / Form module: remove then re-add ---
                if comp is not None:
                    print(f"    Removing existing module...")
                    project.VBComponents.Remove(comp)

                print(
                    f"    Adding module {module_name} (type {desired_type}) via AddFromString..."
                )
                new_comp = project.VBComponents.Add(desired_type)
                new_comp.Name = module_name
                cm = new_comp.CodeModule
                if cm.CountOfLines > 0:
                    cm.DeleteLines(1, cm.CountOfLines)
                if code:
                    cm.AddFromString(code)
                print(f"    [SUCCESS] {module_name} (type {desired_type}) injected.")

        except Exception as e:
            print(f"    [ERROR] Failed to process {module_name}: {e}")

    print(f"--- Completed Raw Import v4 for {project_target_name} ---")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python raw_vba_import.py <ProjectName> <SourceDir>")
        sys.exit(1)
    raw_import_v4(sys.argv[1], sys.argv[2])
