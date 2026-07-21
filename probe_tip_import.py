import shutil
from pathlib import Path
from typing import Any, Optional

import win32com.client

import DB

_dmis_app: Optional[Any] = None
_dmis_parts: Optional[Any] = None


def set_status(message: str) -> None:
    print(f"[STATUS] {message}", flush=True)


def store_tip_angles(file_name: str, file_path: str, department_id: int) -> None:
    global _dmis_parts

    dmis_part = None
    processed_ok = False
    try:
        dmis_part = _dmis_parts.Open(str(file_path), "OFFLINE")
        dmis_commands = dmis_part.Commands

        try:
            command_count = dmis_commands.Count
            last_command = dmis_commands.Item(command_count)
            if last_command is not None:
                dmis_commands.InsertionPointAfter(last_command)
        except TypeError:
            pass  # Item is not callable in this PC-DMIS version – safe to skip

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
                        ",  1"  # IsStillThere = true  (VBA used -1 / Access Boolean)
                        ")"
                    )
                    DB.execute_sql_statement(sql)
        processed_ok = True
    finally:
        _close_part_no_save(dmis_part)
        # Delete the source file only after successful processing so remaining
        # files represent pending work for restart/retry workflows.
        if processed_ok:
            Path(file_path).unlink(missing_ok=True)


def enumerate_pcdmis_programs(directory_path: str, department_id: int) -> None:
    root = Path(directory_path)

    # Walk recursively so queued programs in nested folders are included.
    for entry in root.rglob("*"):
        if not entry.is_file():
            continue

        if entry.suffix.upper() != ".PRG":
            entry.unlink(missing_ok=True)
            continue

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
            entry.unlink(missing_ok=True)


def run_probe_and_tip_import(
        directory_path: str,
        department_id: int,
        delete_unused: bool = False,
        full_refresh: bool = False,
        partial_refresh: bool = False,
        probe_name: str = "",
) -> None:
    global _dmis_app, _dmis_parts
    import_run_id: Optional[int] = None
    run_completed = False

    open_run_sql = (
        "SELECT TOP 1 ID "
        "FROM tblTipAngle_ImportRun "
        f"WHERE DepartmentID = {department_id} "
        "  AND EndTime IS NULL "
        "ORDER BY StartTime DESC"
    )
    open_runs = DB.get_sql_recordset(open_run_sql)
    if open_runs:
        import_run_id = int(open_runs[0]["ID"])
    else:
        start_run_sql = (
            "INSERT INTO tblTipAngle_ImportRun (DepartmentID, StartTime, EndTime) "
            f"VALUES ({department_id}, GETDATE(), NULL); "
            "SELECT CAST(SCOPE_IDENTITY() AS INT)"
        )
        import_run_id = int(DB.execute_sql_scalar_statement(start_run_sql))

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

        run_completed = True

    finally:
        if run_completed and import_run_id is not None:
            close_run_sql = (
                "UPDATE tblTipAngle_ImportRun "
                "SET EndTime = GETDATE() "
                f"WHERE ID = {import_run_id} AND EndTime IS NULL"
            )
            DB.execute_sql_statement(close_run_sql)

        # ── Always clean up COM ───────────────────────────────────────────────
        try:
            if _dmis_app is not None:
                _dmis_app.Quit()
        finally:
            _dmis_parts = None
            _dmis_app = None

    set_status("Done!")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _close_part_no_save(dmis_part: Optional[Any]) -> None:
    """Close an open part without save prompts across COM signature variants."""
    if dmis_part is None:
        return

    close_member = getattr(dmis_part, "Close", None)
    if callable(close_member):
        for args in ((False,), (0,), tuple()):
            try:
                close_member(*args)
                return
            except TypeError:
                continue
            except Exception:
                break

    quit_member = getattr(dmis_part, "Quit", None)
    if callable(quit_member):
        for args in ((False,), (0,), tuple()):
            try:
                quit_member(*args)
                return
            except TypeError:
                continue
            except Exception:
                break


def _esc(value: str) -> str:
    """Escape single quotes for inline SQL strings."""
    return value.replace("'", "''")


if __name__ == "__main__":
    # 15 =     Ortho Turning
    # 12 =     Ortho Mills
    # 6  =     Additive
    # 10 =     Pacing
    # 5 =      Cardio




    directory_path = r"C:\pcdmis-temp"
    department_id = 12
    delete_unused = True
    full_refresh = False
    partial_refresh = False
    probe_name = ""

    run_probe_and_tip_import(
        directory_path=directory_path,
        department_id=department_id,
        delete_unused=delete_unused,
        full_refresh=full_refresh,
        partial_refresh=partial_refresh,
        probe_name=probe_name,
    )
