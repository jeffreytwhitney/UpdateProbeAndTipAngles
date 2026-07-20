"""
probe_tip_import.py

Python conversion of the original Microsoft Access VBA module that enumerates
PC-DMIS part programs and stores probe / tip-angle information into the SQL
Server database via the existing DB.py helpers.

PC-DMIS COM library: C:\Program Files\Hexagon\PC-DMIS 2025.2 64-bit\pcdlrn.tlb
COM ProgID          : PCDLRN.Application

VBA command-type constants used:
    61 → LOADPROBE  – carries the probe name  (gettext(152, 0))
    60 → TIP/ANGLE  – carries the tip id      (gettext(3,   0))
"""

import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import win32com.client

import DB

# ---------------------------------------------------------------------------
# Module-level COM objects (mirrors VBA module-level Dims)
# ---------------------------------------------------------------------------
_dmis_app: Optional[Any] = None
_dmis_parts: Optional[Any] = None

TEMP_DIR = Path(r"C:\pcdmis-temp")

# ---------------------------------------------------------------------------
# Status helper  (mirrors SetStatusBarText)
# ---------------------------------------------------------------------------

def set_status(message: str) -> None:
    print(f"[STATUS] {message}", flush=True)


# ---------------------------------------------------------------------------
# StoreTipAngles
# ---------------------------------------------------------------------------

def store_tip_angles(file_name: str, file_path: str, department_id: int) -> None:
    """
    Open a PC-DMIS part program (offline), walk its commands, and insert
    every probe→tip pairing found into tblTipAngles.

    Mirrors VBA: StoreTipAngles(file_name, file_path, department_id)
    """
    global _dmis_parts

    copy_to_path = TEMP_DIR / file_name
    shutil.copy2(file_path, copy_to_path)

    dmis_part = _dmis_parts.Open(str(copy_to_path), "OFFLINE")
    dmis_commands = dmis_part.Commands

    # Set insertion point to end of program (required by the COM API before
    # iterating – mirrors the original VBA pattern)
    command_count = dmis_commands.Count
    last_command = dmis_commands.Item(command_count)
    dmis_commands.InsertionPointAfter(last_command)

    program_name: str = dmis_part.Name
    set_status(program_name)

    probe_name: str = ""

    for dmis_command in dmis_commands:
        cmd_type = dmis_command.Type

        if cmd_type == 61:  # LOADPROBE
            probe_name = dmis_command.GetText(152, 0)

        elif cmd_type == 60:  # TIP / ANGLE
            tip_id = dmis_command.GetText(3, 0)
            if probe_name:
                sql = (
                    "INSERT INTO tblTipAngles "
                    "  (DepartmentID, ProgramName, ProbeName, TipName, IsStillThere) "
                    "VALUES ("
                    f"  {department_id}"
                    f", '{_esc(program_name)}'"
                    f", '{_esc(probe_name)}'"
                    f", '{_esc(tip_id)}'"
                    ",  1"          # IsStillThere = true  (VBA used -1 / Access Boolean)
                    ")"
                )
                DB.execute_sql_statement(sql)

    dmis_part.Quit()
    copy_to_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# EnumeratePCDMISPrograms  (recursive)
# ---------------------------------------------------------------------------

def enumerate_pcdmis_programs(directory_path: str, department_id: int) -> None:
    """
    Recursively walk *directory_path*, process every *.PRG file found, and
    either insert new records or mark existing ones as still-present.

    Mirrors VBA: EnumeratePCDMISPrograms(directory_path, department_id)
    """
    root = Path(directory_path)

    for entry in root.iterdir():
        if entry.is_file() and entry.suffix.upper() == ".PRG":
            file_name = entry.name
            file_path = str(entry)

            count_sql = (
                "SELECT COUNT(*) FROM tblTipAngles "
                f"WHERE ProgramName = '{_esc(file_name)}' "
                f"  AND DepartmentID = {department_id}"
            )
            record_count = DB.get_sql_scalar(count_sql)

            if record_count == 0:
                store_tip_angles(file_name, file_path, department_id)
            else:
                update_sql = (
                    "UPDATE tblTipAngles SET IsStillThere = 1 "
                    f"WHERE DepartmentID = {department_id} "
                    f"  AND ProgramName  = '{_esc(file_name)}'"
                )
                DB.execute_sql_statement(update_sql)

    for sub in root.iterdir():
        if sub.is_dir():
            enumerate_pcdmis_programs(str(sub), department_id)


# ---------------------------------------------------------------------------
# RunProbeAndTipImport  (main entry-point, mirrors VBA sub of same name)
# ---------------------------------------------------------------------------

def run_probe_and_tip_import(
    directory_path: str,
    department_id: int,
    delete_unused: bool = False,
    full_refresh: bool = False,
    partial_refresh: bool = False,
    probe_name: str = "",
) -> None:

    global _dmis_app, _dmis_parts

    start_time = datetime.now()

    # ── Temp directory ──────────────────────────────────────────────────────
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir(parents=True)

    # ── PC-DMIS COM ─────────────────────────────────────────────────────────
    set_status("Opening PC-DMIS...")
    _dmis_app = win32com.client.Dispatch("PCDLRN.Application")
    _dmis_parts = _dmis_app.PartPrograms

    try:
        # ── Optional pre-import DB cleanup ──────────────────────────────────

        if partial_refresh and probe_name:
            sql = (
                "DELETE FROM tblTipAngles "
                f"WHERE DepartmentID = {department_id} "
                f"  AND ProgramName IN ("
                f"      SELECT DISTINCT ProgramName FROM tblTipAngles "
                f"      WHERE DepartmentID = {department_id} "
                f"        AND ProgramName = '{_esc(probe_name)}'"
                f"  )"
            )
            DB.execute_sql_statement(sql)

        if full_refresh:
            sql = (
                "DELETE FROM tblTipAngles "
                f"WHERE DepartmentID = {department_id}"
            )
            DB.execute_sql_statement(sql)

        if delete_unused:
            sql = (
                "UPDATE tblTipAngles SET IsStillThere = 0 "
                f"WHERE DepartmentID = {department_id}"
            )
            DB.execute_sql_statement(sql)

        # ── Main enumeration ─────────────────────────────────────────────────
        set_status("Processing...")
        enumerate_pcdmis_programs(directory_path, department_id)

        # ── Remove stale records ─────────────────────────────────────────────
        if delete_unused:
            sql = (
                "DELETE FROM tblTipAngles "
                f"WHERE IsStillThere = 0 AND DepartmentID = {department_id}"
            )
            DB.execute_sql_statement(sql)

        # ── Log the import run ───────────────────────────────────────────────
        end_time = datetime.now()
        log_sql = (
            "INSERT INTO tblImportRun (DepartmentID, StartTime, EndTime) VALUES ("
            f"  {department_id}"
            f", '{start_time.strftime('%Y-%m-%dT%H:%M:%S')}'"
            f", '{end_time.strftime('%Y-%m-%dT%H:%M:%S')}'"
            ")"
        )
        DB.execute_sql_statement(log_sql)

    finally:
        # ── Always clean up COM and temp dir ─────────────────────────────────
        _dmis_app.Quit()
        _dmis_parts = None
        _dmis_app = None
        if TEMP_DIR.exists():
            shutil.rmtree(TEMP_DIR, ignore_errors=True)

    set_status("Done!")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _esc(value: str) -> str:
    """Escape single quotes for inline SQL strings."""
    return value.replace("'", "''")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Import PC-DMIS probe/tip-angle data into the database."
    )
    parser.add_argument("directory_path",  help="Root directory containing .PRG files")
    parser.add_argument("department_id",   type=int, help="Department ID")
    parser.add_argument("--delete-unused",    action="store_true",
                        help="Mark then delete records for programs no longer on disk")
    parser.add_argument("--full-refresh",     action="store_true",
                        help="Delete all existing records for this department before import")
    parser.add_argument("--partial-refresh",  action="store_true",
                        help="Delete records for a specific program before re-importing it")
    parser.add_argument("--probe-name",       default="",
                        help="Program name to target when using --partial-refresh")

    args = parser.parse_args()

    run_probe_and_tip_import(
        directory_path=args.directory_path,
        department_id=args.department_id,
        delete_unused=args.delete_unused,
        full_refresh=args.full_refresh,
        partial_refresh=args.partial_refresh,
        probe_name=args.probe_name,
    )

