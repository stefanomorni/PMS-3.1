"""
watch_power_queries.py
======================
Unified Power Query watcher for PMS 3.1.

Startup logic:
  1. If the canonical *_PowerQuery.m file is NEWER than the sentinel
     (.last_combine), a fresh extraction was performed since the last
     watcher session → auto-split it into power-queries/ subfolders.
  2. Start watching power-queries/**/*.m for changes.

On change:
  3. Wait for debounce window (default 600ms) so rapid successive saves
     collapse into a single combine.
  4. Re-combine all split files → overwrite the canonical _PowerQuery.m.
  5. Touch the sentinel .last_combine so the next startup does NOT re-split.
  6. ewc3labs chokidar watcher detects canonical file change → auto-syncs
     to Excel and creates a backup.

Usage (VS Code task or terminal):
    python scripts\\watch_power_queries.py

Dependencies (stdlib-only, no pip required):
    Uses os, time, hashlib, threading, subprocess — all stdlib.
    Falls back to polling if os-level events are unavailable.
"""

import os
import sys
import time
import hashlib
import threading
import subprocess
from pathlib import Path

# ------------------------------------------------------------------ #
#  Paths                                                               #
# ------------------------------------------------------------------ #
SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
PQ_DIR = PROJECT_DIR / "power-queries"
SENTINEL = PQ_DIR / ".last_combine"  # gitignored; only mtime matters

SPLIT_SCRIPT = SCRIPT_DIR / "split_power_queries.py"
COMBINE_SCRIPT = SCRIPT_DIR / "combine_power_queries.py"

# ------------------------------------------------------------------ #
#  Settings                                                            #
# ------------------------------------------------------------------ #
DEBOUNCE_SECONDS = 0.6  # collapse rapid saves within this window
POLL_INTERVAL = 0.5  # filesystem poll interval (seconds)


# ------------------------------------------------------------------ #
#  Helpers                                                             #
# ------------------------------------------------------------------ #
def find_canonical_m() -> Path | None:
    """Return the canonical *_PowerQuery.m file in the project root, if any."""
    candidates = sorted(
        PROJECT_DIR.glob("*_PowerQuery.m"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def sentinel_mtime() -> float:
    """Return mtime of the sentinel file, or 0.0 if it does not exist."""
    try:
        return SENTINEL.stat().st_mtime
    except FileNotFoundError:
        return 0.0


def touch_sentinel():
    """Create / update the sentinel file to record the combine timestamp."""
    SENTINEL.parent.mkdir(parents=True, exist_ok=True)
    SENTINEL.touch()
    print(f"  [sentinel] updated: {SENTINEL.name}")


def file_hash(path: Path) -> str:
    """Return MD5 of a file's content (used to detect meaningful changes)."""
    h = hashlib.md5()
    h.update(path.read_bytes())
    return h.hexdigest()


def run_script(script: Path, label: str) -> bool:
    """Run a companion Python script. Return True on success."""
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=False,
        text=True,
    )
    if result.returncode != 0:
        print(f"  [ERROR] {label} failed (exit {result.returncode})")
        return False
    return True


def collect_split_files() -> list[Path]:
    """Return all *.m files inside power-queries/ subfolders."""
    return [p for p in PQ_DIR.rglob("*.m") if p.name != ".last_combine"]


# ------------------------------------------------------------------ #
#  Startup: auto-split if canonical is newer than sentinel            #
# ------------------------------------------------------------------ #
def check_and_auto_split():
    canonical = find_canonical_m()

    if canonical is None:
        print("[startup] No *_PowerQuery.m found in project root.")
        print("          Extract from Excel first (ewc3labs → Extract Power Query).")
        return

    canon_mtime = canonical.stat().st_mtime
    sentl_mtime = sentinel_mtime()

    if canon_mtime > sentl_mtime:
        age = canon_mtime - sentl_mtime
        print(f"[startup] Canonical file is {age:.0f}s newer than sentinel.")
        print(
            f"          Detected fresh extraction → auto-splitting {canonical.name} ..."
        )
        if run_script(SPLIT_SCRIPT, "split"):
            # After split, touch sentinel so next startup does NOT re-split.
            # Note: we do NOT touch sentinel on combine — only here.
            touch_sentinel()
            print("[startup] Auto-split complete.")
        else:
            print("[startup] Auto-split FAILED. Fix errors before continuing.")
    else:
        split_files = collect_split_files()
        print(f"[startup] Sentinel is current. {len(split_files)} split files ready.")


# ------------------------------------------------------------------ #
#  Watcher: poll power-queries/**/*.m for changes                     #
# ------------------------------------------------------------------ #
class PollWatcher:
    """
    Simple periodic-poll watcher.
    Detects new/modified/deleted .m files inside PQ_DIR.
    On any change (confirmed by file hash), fires on_change() after debounce.
    """

    def __init__(self):
        self._hashes: dict[str, str] = {}
        self._debounce_timer: threading.Timer | None = None
        self._lock = threading.Lock()
        self._running = False

    def _snapshot(self) -> dict[str, str]:
        snapshot = {}
        for p in collect_split_files():
            try:
                snapshot[str(p)] = file_hash(p)
            except OSError:
                pass
        return snapshot

    def _schedule_combine(self, changed_files: list[str]):
        with self._lock:
            if self._debounce_timer:
                self._debounce_timer.cancel()
            self._debounce_timer = threading.Timer(
                DEBOUNCE_SECONDS,
                self._do_combine,
                args=(changed_files,),
            )
            self._debounce_timer.start()

    def _do_combine(self, changed_files: list[str]):
        print(f"\n[change] {len(changed_files)} file(s) modified:")
        for f in changed_files:
            print(f"         • {Path(f).relative_to(PROJECT_DIR)}")
        print("[combine] Recombining → canonical _PowerQuery.m ...")
        if run_script(COMBINE_SCRIPT, "combine"):
            touch_sentinel()
            print("[combine] Done. ewc3labs will sync to Excel automatically.\n")
        else:
            print("[combine] FAILED — canonical file NOT updated.\n")

    def run(self):
        self._running = True
        self._hashes = self._snapshot()
        print(
            f"[watcher] Polling {PQ_DIR.relative_to(PROJECT_DIR)}/**/*.m every {POLL_INTERVAL}s"
        )
        print("[watcher] Press Ctrl+C to stop.\n")

        try:
            while self._running:
                time.sleep(POLL_INTERVAL)
                new_hashes = self._snapshot()
                changed = [k for k, v in new_hashes.items() if self._hashes.get(k) != v]
                deleted = [k for k in self._hashes if k not in new_hashes]
                all_changes = changed + deleted
                if all_changes:
                    self._hashes = new_hashes
                    self._schedule_combine(all_changes)
        except KeyboardInterrupt:
            print("\n[watcher] Stopped by user.")
            if self._debounce_timer:
                self._debounce_timer.cancel()


# ------------------------------------------------------------------ #
#  Entry point                                                         #
# ------------------------------------------------------------------ #
def main():
    print("=" * 60)
    print(" Power Query Watcher — PMS 3.1")
    print("=" * 60)
    print(f" Project : {PROJECT_DIR}")
    print(f" PQ dir  : {PQ_DIR.relative_to(PROJECT_DIR)}")
    print()

    # Step 1: auto-split if needed
    check_and_auto_split()
    print()

    # Step 2: start polling watcher
    watcher = PollWatcher()
    watcher.run()


if __name__ == "__main__":
    main()
