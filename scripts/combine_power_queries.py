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
            f"Manifest not found: {MANIFEST}\nRun split_power_queries.py first."
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
            print(f"  WARNING: {fpath} missing — skipping {q['name']}")
            continue

        raw = fpath.read_text(encoding="utf-8")

        # Strip per-file header comment lines (lines starting with //)
        lines = [l for l in raw.splitlines() if not l.startswith("//")]
        expr  = "\n".join(lines).strip()

        # Prefix annotation if present
        if q.get("annotation"):
            blocks.append(q["annotation"])

        # Re-wrap in section document format: shared <name> = <expr>;
        blocks.append(f"shared {q['name']} = {expr};")
        blocks.append("")   # blank line between entries

    out_path.write_text("\n".join(blocks), encoding="utf-8")
    print(f"Combined file → {out_path}")


if __name__ == "__main__":
    main()
