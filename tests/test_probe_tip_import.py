"""
tests/test_probe_tip_import.py

Run with:  pytest tests/ -v
Live test: pytest tests/ -v -m live   (requires PC-DMIS installed and a real .PRG file)
"""

import shutil
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

import probe_tip_import as pti


# ---------------------------------------------------------------------------
# Helpers / fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def reset_module_globals():
    """Ensure module-level COM handles are None before and after every test."""
    pti._dmis_app = None
    pti._dmis_parts = None
    yield
    pti._dmis_app = None
    pti._dmis_parts = None


@pytest.fixture()
def tmp_prg_dir(tmp_path: Path):
    """Create a small directory tree with fake .PRG files."""
    (tmp_path / "PART_A.PRG").write_text("fake prg content")
    (tmp_path / "PART_B.PRG").write_text("fake prg content")
    (tmp_path / "README.txt").write_text("not a prg")
    sub = tmp_path / "subdir"
    sub.mkdir()
    (sub / "PART_C.PRG").write_text("fake prg content")
    return tmp_path


def _make_command(cmd_type: int, text_map: dict[tuple, str]) -> MagicMock:
    """Return a mock DmisCommand whose .Type and .GetText behave correctly."""
    cmd = MagicMock()
    cmd.Type = cmd_type
    cmd.GetText.side_effect = lambda prop, idx: text_map.get((prop, idx), "")
    return cmd


def _make_dmis_part(program_name: str, commands: list) -> MagicMock:
    """Return a mock DmisPart."""
    part = MagicMock()
    part.Name = program_name
    dmis_cmds = MagicMock()
    dmis_cmds.Count = len(commands)
    dmis_cmds.Item.return_value = commands[-1] if commands else MagicMock()
    dmis_cmds.__iter__ = MagicMock(return_value=iter(commands))
    part.Commands = dmis_cmds
    return part


# ---------------------------------------------------------------------------
# _esc
# ---------------------------------------------------------------------------

class TestEsc:
    def test_no_quotes(self):
        assert pti._esc("hello") == "hello"

    def test_single_quote_escaped(self):
        assert pti._esc("O'Brien") == "O''Brien"

    def test_multiple_quotes(self):
        assert pti._esc("it's a test's") == "it''s a test''s"


# ---------------------------------------------------------------------------
# store_tip_angles
# ---------------------------------------------------------------------------

class TestStoreTipAngles:
    def test_log_identifies_loaded_module_and_file_before_open(self, tmp_path, monkeypatch):
        prg_file = tmp_path / "PART_LOGGED.PRG"
        prg_file.write_text("fake")
        log_file = tmp_path / "crash_log.txt"
        log_file.write_text("stale run")
        monkeypatch.setattr(pti, "LOG_FILE_PATH", log_file)
        pti._clear_log()

        mock_part = _make_dmis_part("PART_LOGGED", [])
        pti._dmis_parts = MagicMock()

        def open_program(path, mode):
            log_contents = log_file.read_text(encoding="utf-8")
            assert f"Loaded module: {Path(pti.__file__).resolve()}" in log_contents
            assert f"Opening: {prg_file}" in log_contents
            return mock_part

        pti._dmis_parts.Open.side_effect = open_program
        pti.store_tip_angles("PART_LOGGED.PRG", str(prg_file), department_id=1)

        assert "stale run" not in log_file.read_text(encoding="utf-8")

    def test_inserts_probe_tip_pairs(self, tmp_path):
        """Each LOADPROBE(61) + TIP(60) pair should produce one INSERT."""
        # Arrange
        prg_file = tmp_path / "PART_A.PRG"
        prg_file.write_text("fake")

        commands = [
            _make_command(61, {(152, 0): "PROBE_1"}),   # LOADPROBE
            _make_command(60, {(3,   0): "T1A0B0"}),    # TIP
            _make_command(60, {(3,   0): "T1A90B0"}),   # TIP (same probe)
            _make_command(61, {(152, 0): "PROBE_2"}),   # LOADPROBE
            _make_command(60, {(3,   0): "T2A0B0"}),    # TIP
        ]
        mock_part = _make_dmis_part("PART_A", commands)
        pti._dmis_parts = MagicMock()
        pti._dmis_parts.Open.return_value = mock_part

        with patch("DB.execute_sql_statement") as mock_exec:
            pti.store_tip_angles("PART_A.PRG", str(prg_file), department_id=1)

        assert mock_exec.call_count == 3
        calls_sql = [c.args[0] for c in mock_exec.call_args_list]
        assert any("PROBE_1" in s and "T1A0B0"  in s for s in calls_sql)
        assert any("PROBE_1" in s and "T1A90B0" in s for s in calls_sql)
        assert any("PROBE_2" in s and "T2A0B0"  in s for s in calls_sql)

    def test_skips_tip_before_any_probe(self, tmp_path):
        """TIP commands that appear before the first LOADPROBE should be ignored."""
        prg_file = tmp_path / "PART_X.PRG"
        prg_file.write_text("fake")

        commands = [
            _make_command(60, {(3, 0): "T1A0B0"}),  # TIP with no probe yet
            _make_command(61, {(152, 0): "PROBE_1"}),
            _make_command(60, {(3, 0): "T1A0B0"}),
        ]
        mock_part = _make_dmis_part("PART_X", commands)
        pti._dmis_parts = MagicMock()
        pti._dmis_parts.Open.return_value = mock_part

        with patch("DB.execute_sql_statement") as mock_exec:
            pti.store_tip_angles("PART_X.PRG", str(prg_file), department_id=2)

        assert mock_exec.call_count == 1  # only the tip after the probe

    def test_source_file_deleted_after_successful_processing(self, tmp_path):
        """The source .PRG should be deleted after successful DB updates."""
        prg_file = tmp_path / "PART_CLEAN.PRG"
        prg_file.write_text("fake")

        mock_part = _make_dmis_part("PART_CLEAN", [_make_command(0, {})])
        pti._dmis_parts = MagicMock()
        pti._dmis_parts.Open.return_value = mock_part

        with patch("DB.execute_sql_statement"):
            pti.store_tip_angles("PART_CLEAN.PRG", str(prg_file), department_id=1)

        assert not prg_file.exists()


# ---------------------------------------------------------------------------
# enumerate_pcdmis_programs
# ---------------------------------------------------------------------------

class TestEnumeratePCDMISPrograms:
    def test_new_files_are_stored(self, tmp_prg_dir, tmp_path):
        """Files not yet in the DB (count=0) should trigger store_tip_angles."""
        pti._dmis_parts = MagicMock()

        with (
            patch("DB.get_sql_scalar", return_value=0),
            patch("probe_tip_import.store_tip_angles") as mock_store,
        ):
            pti.enumerate_pcdmis_programs(str(tmp_prg_dir), department_id=1)

        # PART_A, PART_B, PART_C (in subdir) → 3 calls
        assert mock_store.call_count == 3
        stored_names = {c.args[0] for c in mock_store.call_args_list}
        assert stored_names == {"PART_A.PRG", "PART_B.PRG", "PART_C.PRG"}

    def test_existing_files_are_marked_still_there(self, tmp_prg_dir):
        """Files already in DB (count>0) should be UPDATE'd, not re-stored."""
        with (
            patch("DB.get_sql_scalar", return_value=5),
            patch("DB.execute_sql_statement") as mock_exec,
            patch("probe_tip_import.store_tip_angles") as mock_store,
        ):
            pti.enumerate_pcdmis_programs(str(tmp_prg_dir), department_id=1)

        mock_store.assert_not_called()
        assert mock_exec.call_count == 3
        for c in mock_exec.call_args_list:
            assert "IsStillThere = 1" in c.args[0]

    def test_non_prg_files_are_deleted(self, tmp_prg_dir):
        """Non-.PRG files should be deleted from the queue directory."""
        with (
            patch("DB.get_sql_scalar", return_value=0),
            patch("probe_tip_import.store_tip_angles") as mock_store,
        ):
            pti.enumerate_pcdmis_programs(str(tmp_prg_dir), department_id=1)

        stored_names = {c.args[0] for c in mock_store.call_args_list}
        assert "README.txt" not in stored_names
        assert not (tmp_prg_dir / "README.txt").exists()

    def test_non_prg_files_in_subdirectories_are_deleted(self, tmp_prg_dir):
        """Nested non-.PRG files should be deleted, but directories should remain."""
        nested_noise = tmp_prg_dir / "subdir" / "notes.tmp"
        nested_noise.write_text("delete me")

        with (
            patch("DB.get_sql_scalar", return_value=0),
            patch("probe_tip_import.store_tip_angles"),
        ):
            pti.enumerate_pcdmis_programs(str(tmp_prg_dir), department_id=1)

        assert not nested_noise.exists()
        assert (tmp_prg_dir / "subdir").is_dir()


# ---------------------------------------------------------------------------
# run_probe_and_tip_import  (orchestration)
# ---------------------------------------------------------------------------

class TestRunProbeAndTipImport:
    def _run(self, tmp_path, open_run_rows=None, **kwargs):
        mock_app = MagicMock()
        with (
            patch("win32com.client.Dispatch", return_value=mock_app),
            patch("probe_tip_import.enumerate_pcdmis_programs"),
            patch("DB.get_sql_recordset", return_value=open_run_rows or []),
            patch("DB.execute_sql_scalar_statement", return_value=123) as mock_insert_run,
            patch("DB.execute_sql_statement") as mock_exec,
        ):
            pti.run_probe_and_tip_import(
                directory_path=str(tmp_path),
                department_id=1,
                **kwargs,
            )
            return mock_exec, mock_app, mock_insert_run

    def test_full_refresh_deletes_department_rows(self, tmp_path):
        mock_exec, _, _ = self._run(tmp_path, full_refresh=True)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("DELETE FROM tblTipAngles" in s and "DepartmentID = 1" in s
                   for s in sqls)

    def test_delete_unused_updates_then_deletes(self, tmp_path):
        mock_exec, _, _ = self._run(tmp_path, delete_unused=True)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("IsStillThere = 0" in s for s in sqls)
        assert any("IsStillThere = 0" in s and "DELETE" in s for s in sqls)

    def test_partial_refresh_deletes_named_program(self, tmp_path):
        mock_exec, _, _ = self._run(tmp_path, partial_refresh=True,
                                    probe_name="PART_A.PRG")
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("PART_A.PRG" in s and "DELETE" in s for s in sqls)

    def test_import_run_closed_with_end_time(self, tmp_path):
        mock_exec, _, _ = self._run(tmp_path)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("UPDATE tblTipAngle_ImportRun" in s and "EndTime = GETDATE()" in s for s in sqls)

    def test_import_run_row_created_when_no_open_run_exists(self, tmp_path):
        _, _, mock_insert_run = self._run(tmp_path, open_run_rows=[])
        insert_sql = mock_insert_run.call_args.args[0]
        assert "INSERT INTO tblTipAngle_ImportRun" in insert_sql

    def test_import_run_row_not_created_when_open_run_exists(self, tmp_path):
        _, _, mock_insert_run = self._run(tmp_path, open_run_rows=[{"ID": 77}])
        mock_insert_run.assert_not_called()

    def test_pcdmis_quit_called_on_success(self, tmp_path):
        _, mock_app, _ = self._run(tmp_path)
        mock_app.Quit.assert_called_once()

    def test_pcdmis_quit_called_on_exception(self, tmp_path):
        """PC-DMIS must be shut down even when an exception is raised mid-import."""
        mock_app = MagicMock()
        with (
            patch("win32com.client.Dispatch", return_value=mock_app),
            patch("probe_tip_import.enumerate_pcdmis_programs",
                  side_effect=RuntimeError("boom")),
            patch("DB.get_sql_recordset", return_value=[]),
            patch("DB.execute_sql_scalar_statement", return_value=123),
            patch("DB.execute_sql_statement"),
        ):
            with pytest.raises(RuntimeError, match="boom"):
                pti.run_probe_and_tip_import(str(tmp_path), department_id=1)

        mock_app.Quit.assert_called_once()


# ---------------------------------------------------------------------------
# Multi-department staging/orchestration
# ---------------------------------------------------------------------------

class TestMultiDepartmentImport:
    def test_clear_directory_contents_keeps_root_and_removes_children(self, tmp_path):
        root = tmp_path / "pcdmis-temp"
        root.mkdir()
        (root / "a.txt").write_text("x")
        nested = root / "nested"
        nested.mkdir()
        (nested / "b.PRG").write_text("y")

        pti._clear_directory_contents(root)

        assert root.exists()
        assert list(root.iterdir()) == []

    def test_copy_prg_files_to_temp_copies_only_prg_files(self, tmp_path):
        source = tmp_path / "source"
        source.mkdir()
        (source / "A.PRG").write_text("a")
        (source / "B.txt").write_text("b")
        sub = source / "sub"
        sub.mkdir()
        (sub / "C.prg").write_text("c")

        temp = tmp_path / "temp"
        copied = pti._copy_prg_files_to_temp(source, temp)

        assert copied == 2
        assert (temp / "A.PRG").is_file()
        assert (temp / "sub" / "C.prg").is_file()
        assert not (temp / "B.txt").exists()

    def test_run_multi_department_import_stages_and_runs_each_department(self, tmp_path):
        departments = [
            pti.DepartmentImport(department_id=15, department_name="Turning", dirpath=r"C:\src\turning"),
            pti.DepartmentImport(department_id=12, department_name="Mills", dirpath=r"C:\src\mills"),
        ]

        with (
            patch("probe_tip_import._ensure_pcdmis_session"),
            patch("probe_tip_import._shutdown_pcdmis_session"),
            patch("probe_tip_import._clear_directory_contents") as mock_clear,
            patch("probe_tip_import._copy_prg_files_to_temp", return_value=1) as mock_copy,
            patch("probe_tip_import.run_probe_and_tip_import", return_value={"newly_added": 2, "updated": 3}) as mock_run,
        ):
            pti.run_multi_department_import(
                departments=departments,
                temp_directory=str(tmp_path / "pcdmis-temp"),
                delete_unused=True,
                full_refresh=False,
                partial_refresh=False,
                probe_name="",
            )

        assert mock_clear.call_count == 2
        assert mock_copy.call_count == 2
        assert mock_run.call_args_list == [
            call(
                directory_path=str(tmp_path / "pcdmis-temp"),
                department_id=15,
                delete_unused=True,
                full_refresh=False,
                partial_refresh=False,
                probe_name="",
                manage_session=False,
            ),
            call(
                directory_path=str(tmp_path / "pcdmis-temp"),
                department_id=12,
                delete_unused=True,
                full_refresh=False,
                partial_refresh=False,
                probe_name="",
                manage_session=False,
            ),
        ]

    def test_run_multi_department_import_skips_when_no_prg_files(self, tmp_path):
        departments = [pti.DepartmentImport(department_id=20, department_name="Anoka", dirpath=r"C:\src\anoka")]

        with (
            patch("probe_tip_import._ensure_pcdmis_session"),
            patch("probe_tip_import._shutdown_pcdmis_session"),
            patch("probe_tip_import._clear_directory_contents"),
            patch("probe_tip_import._copy_prg_files_to_temp", return_value=0),
            patch("probe_tip_import.run_probe_and_tip_import") as mock_run,
        ):
            pti.run_multi_department_import(
                departments=departments,
                temp_directory=str(tmp_path / "pcdmis-temp"),
            )

        mock_run.assert_not_called()

    def test_run_multi_department_import_opens_and_quits_com_once(self, tmp_path):
        departments = [
            pti.DepartmentImport(department_id=15, department_name="Turning", dirpath=r"C:\src\turning"),
            pti.DepartmentImport(department_id=12, department_name="Mills", dirpath=r"C:\src\mills"),
        ]

        with (
            patch("probe_tip_import._clear_directory_contents"),
            patch("probe_tip_import._copy_prg_files_to_temp", return_value=1),
            patch("probe_tip_import.run_probe_and_tip_import", return_value={"newly_added": 1, "updated": 0}),
            patch("probe_tip_import.win32com.client.Dispatch") as mock_dispatch,
        ):
            mock_app = MagicMock()
            mock_dispatch.return_value = mock_app
            pti.run_multi_department_import(
                departments=departments,
                temp_directory=str(tmp_path / "pcdmis-temp"),
            )

        mock_dispatch.assert_called_once_with("PCDLRN.Application")
        mock_app.Quit.assert_called_once()


# ---------------------------------------------------------------------------
# LIVE integration test  – requires PC-DMIS 2025.2 installed
# ---------------------------------------------------------------------------

@pytest.mark.live
class TestLivePCDMIS:
    """
    Spins up a real PC-DMIS instance via COM, opens a user-supplied .PRG file
    offline, and verifies that at least one LOADPROBE (type 61) or TIP (type 60)
    command is discovered.

    Configure the path to a real .PRG file either:
      • via the environment variable  LIVE_PRG_FILE
      • or by editing PRG_FILE below.

    Run with:  pytest tests/ -v -m live
    """

    PRG_FILE = r""   # ← set to a real .PRG path, or use env var LIVE_PRG_FILE

    @pytest.fixture(autouse=True)
    def require_prg_file(self):
        import os
        raw = os.environ.get("LIVE_PRG_FILE", self.PRG_FILE)
        if not raw:
            pytest.skip(
                "No .PRG file available for live test. "
                "Set LIVE_PRG_FILE env var or edit TestLivePCDMIS.PRG_FILE."
            )
        # Resolve relative to the project root (parent of the tests/ folder)
        prg = Path(raw)
        if not prg.is_absolute():
            prg = (Path(__file__).parent.parent / prg).resolve()
        if not prg.is_file():
            pytest.skip(f"LIVE_PRG_FILE path does not exist: {prg}")
        self.prg_path = prg

    def test_open_program_and_read_commands(self, tmp_path):
        """Open a real .PRG, iterate commands, confirm type 61 or 60 is present."""
        import win32com.client

        temp_dir = tmp_path / "pcdmis-temp"
        temp_dir.mkdir()
        copy_path = temp_dir / self.prg_path.name
        shutil.copy2(self.prg_path, copy_path)

        dmis_app = win32com.client.Dispatch("PCDLRN.Application")
        dmis_part = None
        try:
            dmis_parts = dmis_app.PartPrograms
            dmis_part = dmis_parts.Open(str(copy_path), "OFFLINE")
            dmis_commands = dmis_part.Commands

            command_count = dmis_commands.Count
            assert command_count > 0, "Program has no commands"

            found_types = {cmd.Type for cmd in dmis_commands}
            print(f"\n[LIVE] Program: {dmis_part.Name}")
            print(f"[LIVE] Command count: {command_count}")
            print(f"[LIVE] Command types present: {sorted(found_types)}")

            probe_tip_types = found_types & {60, 61}
            assert probe_tip_types, (
                f"Expected at least one command of type 60 or 61; "
                f"found types: {sorted(found_types)}"
            )

            # Verify GetText works for each relevant command
            for cmd in dmis_commands:
                if cmd.Type == 61:
                    name = cmd.GetText(152, 0)
                    print(f"[LIVE]   LOADPROBE → {name!r}")
                    assert isinstance(name, str)
                elif cmd.Type == 60:
                    tip = cmd.GetText(3, 0)
                    print(f"[LIVE]   TIP       → {tip!r}")
                    assert isinstance(tip, str)


        finally:
            # Use the same shutdown behavior as production code: close part
            # without prompting to save, then quit the application.
            pti._close_part_no_save(dmis_part)
            dmis_app.Quit()
            shutil.rmtree(temp_dir, ignore_errors=True)
