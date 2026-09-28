import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import win32com.client

import DB

_dmis_app: Optional[Any] = None
_dmis_parts: Optional[Any] = None

LOG_FILE_PATH = Path(__file__).resolve().parent / "crash_log.txt"


def _clear_log() -> None:
    """Wipe the crash log at the start of a run so it only reflects this run."""
    LOG_FILE_PATH.write_text(
        f"Run started: {datetime.now().isoformat(timespec='seconds')}\n"
        f"Loaded module: {Path(__file__).resolve()}\n"
        f"Python executable: {Path(sys.executable).resolve()}\n"
        f"Working directory: {Path.cwd()}\n"
        f"Log file: {LOG_FILE_PATH.resolve()}\n",
        encoding="utf-8",
    )


def _log_opening_file(file_path: str) -> None:
    """Record the file about to be opened, so a crash can be traced to it."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE_PATH, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] Opening: {file_path}\n")


@dataclass(frozen=True)
class DepartmentImport:
    department_id: int
    department_name: str
    dirpath: str


def set_status(message: str) -> None:
    print(f"[STATUS] {message}", flush=True)


def store_tip_angles(file_name: str, file_path: str, department_id: int) -> None:
    global _dmis_parts

    dmis_part = None
    processed_ok = False
    try:
        _log_opening_file(file_path)
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


def enumerate_pcdmis_programs(directory_path: str, department_id: int) -> dict[str, int]:
    """Enumerate and process .PRG files, returning counts of newly added vs. updated."""
    root = Path(directory_path)
    newly_added = 0
    updated = 0

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
            newly_added += 1
        else:
            update_sql = (
                "UPDATE tblTipAngles SET IsStillThere = 1 "
                f"WHERE DepartmentID = {department_id} "
                f"  AND ProgramName  = '{_esc(file_name)}'"
            )
            DB.execute_sql_statement(update_sql)
            entry.unlink(missing_ok=True)
            updated += 1

    return {"newly_added": newly_added, "updated": updated}


def run_probe_and_tip_import(
        directory_path: str,
        department_id: int,
        delete_unused: bool = False,
        full_refresh: bool = False,
        partial_refresh: bool = False,
        probe_name: str = "",
        manage_session: bool = True,
) -> dict[str, int]:
    global _dmis_app, _dmis_parts
    import_run_id: Optional[int] = None
    run_completed = False
    owns_session = False
    counts = {"newly_added": 0, "updated": 0}

    if manage_session:
        _clear_log()

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
    if manage_session:
        set_status("Opening PC-DMIS...")
        _ensure_pcdmis_session()
        owns_session = True

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
        counts = enumerate_pcdmis_programs(directory_path, department_id)

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

        # ── Clean up COM only when this call owns the lifecycle ──────────────
        if owns_session:
            _shutdown_pcdmis_session()

    set_status("Done!")
    return counts


def run_multi_department_import(
        departments: list[DepartmentImport],
        temp_directory: str,
        delete_unused: bool = False,
        full_refresh: bool = False,
        partial_refresh: bool = False,
        probe_name: str = "",
) -> None:
    temp_root = Path(temp_directory)
    department_stats: list[tuple[str, int, int]] = []

    _clear_log()
    set_status("Opening PC-DMIS...")
    _ensure_pcdmis_session()
    try:
        for department in departments:
            set_status(
                f"Staging DepartmentID={department.department_id} from {department.dirpath}"
            )
            _clear_directory_contents(temp_root)
            copied_count = _copy_prg_files_to_temp(Path(department.dirpath), temp_root)

            if copied_count == 0:
                set_status(
                    f"No .PRG files found for DepartmentID={department.department_name}; skipping"
                )
                continue
            else:
                set_status(f"Copying PRG files to {department.department_name}")

            counts = run_probe_and_tip_import(
                directory_path=str(temp_root),
                department_id=department.department_id,
                delete_unused=delete_unused,
                full_refresh=full_refresh,
                partial_refresh=partial_refresh,
                probe_name=probe_name,
                manage_session=False,
            )
            newly_added = counts.get("newly_added", 0)
            updated = counts.get("updated", 0)
            department_stats.append((department.department_name, newly_added, updated))
    finally:
        _shutdown_pcdmis_session()

    # ── Summary report ──────────────────────────────────────────────────────
    set_status("=" * 70)
    set_status("IMPORT SUMMARY")
    set_status("=" * 70)
    total_newly_added = 0
    total_updated = 0
    for dept_name, newly_added, updated in department_stats:
        set_status(f"{dept_name:30} | Added: {newly_added:4} | Updated: {updated:4}")
        total_newly_added += newly_added
        total_updated += updated
    set_status("-" * 70)
    set_status(f"{'TOTAL':30} | Added: {total_newly_added:4} | Updated: {total_updated:4}")
    set_status("=" * 70)


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


def _ensure_pcdmis_session() -> None:
    """Create PC-DMIS COM app/session if not already available."""
    global _dmis_app, _dmis_parts
    if _dmis_app is None or _dmis_parts is None:
        _dmis_app = win32com.client.Dispatch("PCDLRN.Application.20.2")
        _dmis_parts = _dmis_app.PartPrograms


def _shutdown_pcdmis_session() -> None:
    """Dispose PC-DMIS COM app/session safely."""
    global _dmis_app, _dmis_parts
    try:
        if _dmis_app is not None:
            _dmis_app.Quit()
    finally:
        _dmis_parts = None
        _dmis_app = None


def _clear_directory_contents(directory: Path) -> None:
    """Delete all files and folders inside a directory while keeping the root."""
    directory.mkdir(parents=True, exist_ok=True)
    for entry in directory.iterdir():
        if entry.is_dir():
            shutil.rmtree(entry, ignore_errors=True)
        else:
            entry.unlink(missing_ok=True)


def _copy_prg_files_to_temp(source_root: Path, temp_root: Path) -> int:
    """Copy .PRG files recursively from source_root into temp_root."""
    temp_root.mkdir(parents=True, exist_ok=True)
    copied_count: int = 0

    for source_file in source_root.rglob("*"):
        if not source_file.is_file() or source_file.suffix.upper() != ".PRG":
            continue

        relative_path = source_file.relative_to(source_root)
        destination_path = temp_root / relative_path
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_file, destination_path)
        copied_count += 1

    return copied_count


if __name__ == "__main__":
    # 15 =     Ortho Turning
    # 12 =     Ortho Mills
    # 6  =     Additive
    # 10 =     Pacing
    # 5 =      Cardio
    # 20 =     Anoka

    departments = [
        DepartmentImport(department_id=15, department_name="Ortho Turning", dirpath=r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Ortho Turning Approved"),
        DepartmentImport(department_id=12, department_name="Ortho Mills", dirpath=r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Ortho Mill Approved"),
        DepartmentImport(department_id=6, department_name="Additive", dirpath=r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Additive Approved"),
        DepartmentImport(department_id=10, department_name="Pacing", dirpath=r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Pacing Approved"),
        DepartmentImport(department_id=5, department_name="Cardio", dirpath=r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Cardio Approved"),
        DepartmentImport(department_id=20, department_name="Anoka", dirpath=r"\\vrmss-fs1\DNC\CMM Programs\LEVEL 2 Approved Programs"),
    ]




    temp_directory = r"C:\pcdmis-temp"
    delete_unused = True
    full_refresh = False
    partial_refresh = False
    probe_name = ""

    run_multi_department_import(
        departments=departments,
        temp_directory=temp_directory,
        delete_unused=delete_unused,
        full_refresh=full_refresh,
        partial_refresh=partial_refresh,
        probe_name=probe_name,
    )
