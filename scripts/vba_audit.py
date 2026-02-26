import os
import re
from pathlib import Path

# Common VBA/Excel/Office keywords, functions, constants, and objects
VBA_BUILTINS = {
    "true",
    "false",
    "nothing",
    "vbnullstring",
    "vbcritical",
    "vbtab",
    "vbnewline",
    "vbinformation",
    "vbexclamation",
    "vbquestion",
    "vbyesno",
    "vbyes",
    "vbno",
    "vbokonly",
    "vbokcancel",
    "vbretrycancel",
    "vbabortretryignore",
    "vbok",
    "activeworkbook",
    "activesheet",
    "thisworkbook",
    "application",
    "range",
    "instr",
    "mid",
    "left",
    "right",
    "lcase",
    "ucase",
    "replace",
    "msgbox",
    "exit",
    "sub",
    "function",
    "option",
    "explicit",
    "next",
    "for",
    "each",
    "in",
    "if",
    "then",
    "else",
    "elseif",
    "end",
    "select",
    "case",
    "dim",
    "static",
    "public",
    "private",
    "set",
    "let",
    "with",
    "while",
    "wend",
    "do",
    "loop",
    "until",
    "as",
    "variant",
    "string",
    "long",
    "integer",
    "boolean",
    "object",
    "double",
    "single",
    "date",
    "currency",
    "byte",
    "new",
    "collection",
    "dictionary",
    "now",
    "environ",
    "createobject",
    "vba",
    "debug",
    "print",
    "join",
    "split",
    "typename",
    "vartype",
    "ismissing",
    "isempty",
    "isnull",
    "cstr",
    "clng",
    "cint",
    "cbool",
    "cdbl",
    "csng",
    "cdate",
    "ccur",
    "cbyte",
    "cvar",
    "fix",
    "datediff",
    "dateadd",
    "int",
    "err",
    "number",
    "clear",
    "description",
    "iif",
    "timer",
    "timevalue",
    "datevalue",
    "year",
    "month",
    "day",
    "hour",
    "minute",
    "second",
    "weekday",
    "weekdayname",
    "monthname",
    "dateserial",
    "timeserial",
    "time",
    "ubound",
    "lbound",
    "array",
    "filter",
    "strconv",
    "instrrev",
    "space",
    "len",
    "str",
    "val",
    "abs",
    "sqr",
    "log",
    "exp",
    "atn",
    "cos",
    "sin",
    "tan",
    "rnd",
    "randomize",
    "sgn",
    "hex",
    "oct",
    "chr",
    "chrw",
    "asc",
    "trim",
    "ltrim",
    "rtrim",
    "format",
    "formatcurrency",
    "formatdatetime",
    "formatnumber",
    "formatpercent",
    "doevents",
    "inputbox",
    "dir",
    "getattr",
    "setattr",
    "filelen",
    "filedatetime",
    "fileattr",
    "freefile",
    "open",
    "close",
    "get",
    "put",
    "input",
    "print",
    "write",
    "seek",
    "loc",
    "lof",
    "environ",
    "shell",
    "sendkeys",
    "beep",
    "call",
    "resume",
    "on",
    "error",
    "goto",
    "stop",
    "with",
    "byval",
    "byref",
    "optional",
    "paramarray",
    "attribute",
    "empty",
    "null",
    "me",
    "activewindow",
    "activecell",
    "selection",
    "cells",
    "columns",
    "rows",
    "worksheets",
    "sheets",
    "workbooks",
    "intersect",
    "union",
    "on",
    "error",
    "resume",
    "vbscript",
    "regex",
    "matches",
    "objresult",
    "vbmsgboxresult",
    "vbscript",
    "excel",
    "office",
    "msforms",
    "listobject",
    "listrow",
    "querytable",
    "query",
    "adodb",
    "connection",
    "recordset",
    "filesystemobject",
    "textstream",
    "file",
    "scripting",
    "environ",
    "is",
    "not",
    "to",
    "and",
    "or",
    "xor",
    "shmain",
}

# Excel Constants (prefixes)
EXCEL_CONST_PREFIXES = ["xl", "vb", "ms"]

RE_PROC_START = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Static)?\s*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([a-zA-Z0-9_]+)\s*\((.*?)\)",
    re.IGNORECASE,
)
RE_DIM = re.compile(
    r"\b(?:Dim|Static|Public|Private|Const)\s+([a-zA-Z0-9_]+)(?:\s+As\s+[a-zA-Z0-9_]+)?",
    re.IGNORECASE,
)
RE_FOR_EACH = re.compile(r"\bFor\s+Each\s+([a-zA-Z0-9_]+)\s+In\b", re.IGNORECASE)
RE_WORDS = re.compile(r"\b([a-zA-Z][a-zA-Z0-9_]*)\b")

PROJECTS = {
    "MORFunctions": r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORFunctions\vba-files",
    "MORProcedures": r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\vba-files",
    "PMS 3.1": r"D:\Cloud\OneDrive\MC\Investimenti\PMS\PMS 3.1\vba-files",
}


def is_valid_vba_identifier(word, module_names, local_decls):
    wl = word.lower()
    if wl in VBA_BUILTINS:
        return True
    for prefix in EXCEL_CONST_PREFIXES:
        if wl.startswith(prefix):
            return True
    if word in module_names:
        return True
    if word in local_decls:
        return True
    return False


def audit_vba_file(filepath, module_names):
    try:
        with open(filepath, "r", encoding="latin-1") as f:
            lines = f.readlines()
    except Exception:
        return []

    module_vars = set()
    procedures = []
    current_proc = None

    # First pass: map structure
    for i, line in enumerate(lines):
        code_part = line.split("'")[0].strip()
        if not code_part:
            continue

        proc_match = RE_PROC_START.match(code_part)
        if proc_match:
            name = proc_match.group(1)
            params_str = proc_match.group(2)
            params = set()
            for p in params_str.split(","):
                p = p.strip()
                if not p:
                    continue
                p_parts = p.split()
                for part in p_parts:
                    clean_part = re.sub(r"[^a-zA-Z0-9_]", "", part)
                    if clean_part and clean_part.lower() not in {
                        "byval",
                        "byref",
                        "optional",
                        "paramarray",
                        "as",
                        "string",
                        "long",
                        "integer",
                        "boolean",
                        "object",
                        "double",
                        "single",
                        "date",
                        "currency",
                        "byte",
                        "variant",
                    }:
                        params.add(clean_part)
                        break
            current_proc = {
                "name": name,
                "start": i,
                "params": params,
                "declarations": set(),
                "suspects": set(),
            }
            procedures.append(current_proc)
            continue

        if any(
            keyword in code_part
            for keyword in ["End Sub", "End Function", "End Property"]
        ):
            current_proc = None
            continue

        for dim_match in RE_DIM.finditer(code_part):
            var_name = dim_match.group(1)
            if current_proc:
                current_proc["declarations"].add(var_name)
            else:
                module_vars.add(var_name)

    # Second pass: usages
    current_proc = None
    for i, line in enumerate(lines):
        code_part = line.split("'")[0].strip()
        if not code_part:
            continue

        proc_match = RE_PROC_START.match(code_part)
        if proc_match:
            current_proc = next(
                (
                    p
                    for p in procedures
                    if p["name"] == proc_match.group(1) and p["start"] == i
                ),
                None,
            )
            continue

        if any(
            keyword in code_part
            for keyword in ["End Sub", "End Function", "End Property"]
        ):
            current_proc = None
            continue

        if current_proc:
            for word_match in RE_WORDS.finditer(code_part):
                word = word_match.group(1)

                # Check preceding context (e.g. not a member access like .Name)
                idx = word_match.start()
                if idx > 0 and code_part[idx - 1] == ".":
                    continue

                # Heuristics for non-variables
                if word == current_proc["name"]:
                    continue
                if word.startswith("sh") or word.startswith("Foglio"):
                    continue
                if any(p["name"] == word for p in procedures):
                    continue

                if not is_valid_vba_identifier(
                    word,
                    module_names,
                    current_proc["declarations"]
                    .union(current_proc["params"])
                    .union(module_vars),
                ):
                    current_proc["suspects"].add(word)

    results = []
    for p in procedures:
        if p["suspects"]:
            results.append(
                {"procedure": p["name"], "missing": sorted(list(p["suspects"]))}
            )
    return results


def get_all_module_names():
    names = set()
    for vba_dir in PROJECTS.values():
        path_obj = Path(vba_dir)
        if path_obj.exists():
            for f in path_obj.glob("*.*"):
                names.add(f.stem)
    return names


def run_audit():
    module_names = get_all_module_names()
    all_results = {}
    for project, vba_dir in PROJECTS.items():
        results = []
        path_obj = Path(vba_dir)
        if not path_obj.exists():
            continue
        for file_path in path_obj.glob("*.*"):
            if file_path.suffix.lower() in {".bas", ".cls", ".frm"}:
                file_results = audit_vba_file(str(file_path), module_names)
                if file_results:
                    results.append({"file": file_path.name, "issues": file_results})
        all_results[project] = results
    return all_results


if __name__ == "__main__":
    report = run_audit()
    print("\n--- REFINED VBA OPTION EXPLICIT AUDIT REPORT ---\n")
    for project, files in report.items():
        if not files:
            print(f"## {project}: Clean!\n")
            continue
        print(f"## {project}: {len(files)} files with potential missing declarations\n")
        for f in files:
            issues_found = False
            for issue in f["issues"]:
                if issue["missing"]:
                    if not issues_found:
                        print(f"  [{f['file']}]")
                        issues_found = True
                    print(f"    - {issue['procedure']}: {', '.join(issue['missing'])}")
            if issues_found:
                print("")
