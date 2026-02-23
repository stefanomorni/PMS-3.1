"""
split_power_queries.py
======================
Purpose : Split the single combined .m file produced by ewc3labs "Extract Power Query
          from Excel" into individual per-query .m files compatible with VS Code editing.

Key format rule
---------------
  Combined file  : section document  → each query is  `shared <name> = <expr>;`
  Individual files: plain M expression → just  `<expr>`  (no `shared`, no trailing `;`)

Folder structure produced:
    power-queries/
        tables/          result-producing queries (not prefixed fn_ / lukb_ / get_)
        functions/       callable functions (fn_*, get_*, lukb_* prefix)
        lookups/         reference tables (*_tbl)
        _internal/       test / debug helpers (*test*)

Companion script:
    scripts/combine_power_queries.py   → merges per-file back → combined section doc

Usage:
    python scripts/split_power_queries.py [input_file]

    If input_file is omitted it defaults to the most recent *_PowerQuery.m in
    the project root.
"""

import re
import sys
import json
from pathlib import Path
from datetime import datetime, timezone

# ------------------------------------------------------------------ #
#  Paths                                                               #
# ------------------------------------------------------------------ #
SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
PQ_DIR = PROJECT_DIR / "power-queries"

# ------------------------------------------------------------------ #
#  Categorisation rules (first match wins)                            #
# ------------------------------------------------------------------ #
CATEGORY_RULES = [
    ("_internal", lambda n: "test" in n.lower()),
    ("lookups", lambda n: n.endswith("_tbl")),
    (
        "functions",
        lambda n: n.startswith("fn_") or n.startswith("lukb_") or n.startswith("get_"),
    ),
    ("tables", lambda n: True),  # catch-all
]

# ------------------------------------------------------------------ #
#  Regex helpers                                                       #
# ------------------------------------------------------------------ #
SECTION_HEADER_RE = re.compile(r"^section\s+\w+\s*;", re.MULTILINE | re.IGNORECASE)

# Matches an optional [ Description = "..." ] annotation, then the shared declaration.
# Capture groups: 1=annotation (or None), 2=query name
QUERY_DECL_RE = re.compile(
    r"^(\[.*?\]\s*\n)?"  # group 1: optional annotation
    r"shared\s+"
    r'(#"[^"]+"|[\w\.#\'!]+)',  # group 2: query name
    re.MULTILINE,
)


# ------------------------------------------------------------------ #
#  Helpers                                                             #
# ------------------------------------------------------------------ #
def sanitise_filename(name: str) -> str:
    """Produce a safe filename from an M query name (strips #"...") ."""
    if name.startswith('#"') and name.endswith('"'):
        name = name[2:-1]
    return re.sub(r'[\\/:*?"<>|!\'@#\[\]]+', "_", name)


def categorise(name: str) -> str:
    plain = name.strip('#"').strip("'")
    for subfolder, pred in CATEGORY_RULES:
        if pred(plain):
            return subfolder
    return "tables"


def find_input_file() -> Path:
    candidates = sorted(
        PROJECT_DIR.glob("*_PowerQuery.m"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(
            f"No *_PowerQuery.m file found in {PROJECT_DIR}.\n"
            "Use the ewc3labs extension → 'Extract Power Query from Excel' first."
        )
    return candidates[0]


def parse_queries(text: str) -> list[dict]:
    """
    Return list of dicts with keys:
        name        : raw M name (e.g. #"'CHF M'!col_ID" or fn_expand_all_tables)
        annotation  : optional [ Description = "..." ] string (may be None)
        expr        : pure M expression – NO `shared`, NO trailing `;`
    """
    body = SECTION_HEADER_RE.sub("", text).strip()
    matches = list(QUERY_DECL_RE.finditer(body))
    if not matches:
        raise ValueError("No 'shared' queries found in the file.")

    queries = []
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(body)
        block = body[start:end]

        annotation = m.group(1)  # may be None
        name = m.group(2)

        # Remove the leading annotation + "shared <name> =" to get the raw expression
        inner_start = m.end()  # position right after name match within `block`
        relative_start = inner_start - start

        expr_raw = block[relative_start:]

        # Strip leading " = " or "=" (the equals sign after the name)
        expr_raw = re.sub(r"^\s*=\s*", "", expr_raw)

        # Strip the trailing ";" (section-document terminator)
        expr_raw = expr_raw.rstrip()
        if expr_raw.endswith(";"):
            expr_raw = expr_raw[:-1].rstrip()

        queries.append(
            {
                "name": name,
                "annotation": (annotation or "").strip() or None,
                "expr": expr_raw,
            }
        )

    return queries


def write_query_files(queries: list[dict], source_path: Path) -> list[dict]:
    now_iso = datetime.now(timezone.utc).isoformat(timespec="seconds")
    written = []

    for q in queries:
        subfolder = categorise(q["name"])
        target_dir = PQ_DIR / subfolder
        target_dir.mkdir(parents=True, exist_ok=True)

        filename = sanitise_filename(q["name"]) + ".m"
        target_file = target_dir / filename

        lines = [
            f"// Query   : {q['name']}",
            f"// Category: {subfolder}",
            f"// Source  : {source_path.name}",
            f"// Split   : {now_iso}",
        ]
        if q["annotation"]:
            lines.append(f"// Note    : {q['annotation']}")
        lines.append("")  # blank line before code
        lines.append(q["expr"].strip())  # pure M expression
        lines.append("")  # trailing newline

        target_file.write_text("\n".join(lines), encoding="utf-8")
        print(f"  [{subfolder:9s}] {filename}")

        written.append(
            {
                "name": q["name"],
                "annotation": q["annotation"],
                "file": str(target_file.relative_to(PROJECT_DIR)),
                "category": subfolder,
            }
        )

    return written


def write_manifest(written: list[dict], source_path: Path):
    manifest_path = PQ_DIR / "manifest.json"
    data = {
        "generated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "note": "Auto-generated by split_power_queries.py. Do not edit manually.",
        "source_file": source_path.name,
        "queries": written,
    }
    manifest_path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"\n  Manifest  → {manifest_path.relative_to(PROJECT_DIR)}")


def write_combine_script():
    combine_path = SCRIPT_DIR / "combine_power_queries.py"
    combine_path.write_text(COMBINE_SCRIPT_CONTENT, encoding="utf-8")
    print(f"  Combiner  → {combine_path.relative_to(PROJECT_DIR)}")


# ------------------------------------------------------------------ #
#  The companion combine script (embedded as a string)                #
# ------------------------------------------------------------------ #
COMBINE_SCRIPT_CONTENT = '''\
"""
combine_power_queries.py
========================
Merges individual per-query .m files back into a single Section document
suitable for import via the ewc3labs extension.

Output is always: <ExcelFileName>_PowerQuery.m  (canonical ewc3labs name)
This file is picked up automatically by the ewc3labs watcher.

Usage:
    python scripts/combine_power_queries.py [output_file]
"""

import json
import sys
from pathlib import Path
from datetime import datetime, timezone

SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
PQ_DIR      = PROJECT_DIR / "power-queries"
MANIFEST    = PQ_DIR / "manifest.json"


def find_canonical_output() -> Path:
    """
    Return the canonical *_PowerQuery.m path.
    Priority:
      1. manifest["source_file"] if present (set by split_power_queries.py)
      2. Any existing *_PowerQuery.m in project root
      3. Fallback: <ProjectDir.name>_PowerQuery.m
    """
    try:
        data = json.loads(MANIFEST.read_text(encoding="utf-8"))
        source_file = data.get("source_file")
        if source_file:
            return PROJECT_DIR / source_file
    except Exception:
        pass

    # Legacy manifests: look for existing canonical file
    candidates = sorted(
        PROJECT_DIR.glob("*_PowerQuery.m"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]

    # Last resort
    return PROJECT_DIR / f"{PROJECT_DIR.name}_PowerQuery.m"


def main():
    if not MANIFEST.exists():
        raise FileNotFoundError(
            f"Manifest not found: {MANIFEST}\\nRun split_power_queries.py first."
        )

    data    = json.loads(MANIFEST.read_text(encoding="utf-8"))
    queries = data["queries"]

    # Determine output path
    if len(sys.argv) > 1:
        out_path = PROJECT_DIR / sys.argv[1]
    else:
        out_path = find_canonical_output()

    now_iso = datetime.now(timezone.utc).isoformat(timespec="seconds")

    blocks = [
        f"// Power Query — {PROJECT_DIR.name}",
        f"// Generated : {now_iso}",
        f"// Queries   : {len(queries)}",
        "",
        "section Section1;",
        "",
    ]

    for q in queries:
        fpath = PROJECT_DIR / q["file"]
        if not fpath.exists():
            print(f"  WARNING: {fpath} missing — skipping {q[\'name\']}")
            continue

        raw = fpath.read_text(encoding="utf-8")

        # Strip per-file header comment lines (lines starting with //)
        lines = [l for l in raw.splitlines() if not l.startswith("//")]
        expr  = "\\n".join(lines).strip()

        # Prefix annotation if present
        if q.get("annotation"):
            blocks.append(q["annotation"])

        # Re-wrap in section document format: shared <name> = <expr>;
        blocks.append(f"shared {q[\'name\']} = {expr};")
        blocks.append("")   # blank line between entries

    out_path.write_text("\\n".join(blocks), encoding="utf-8")
    print(f"Combined file → {out_path}")


if __name__ == "__main__":
    main()
'''


# ------------------------------------------------------------------ #
#  Main                                                                #
# ------------------------------------------------------------------ #
def main():
    input_path = Path(sys.argv[1]) if len(sys.argv) > 1 else find_input_file()

    print(f"\nInput   : {input_path}")
    print(f"Output  : {PQ_DIR}\n")

    text = input_path.read_text(encoding="utf-8")
    queries = parse_queries(text)
    print(f"Found {len(queries)} queries:\n")

    written = write_query_files(queries, input_path)
    write_manifest(written, input_path)
    write_combine_script()

    print(f"\nDone — {len(written)} queries → {PQ_DIR.relative_to(PROJECT_DIR)}/")
    print(
        "Edit .m files freely, then run  scripts/combine_power_queries.py  to merge for import."
    )


if __name__ == "__main__":
    main()
