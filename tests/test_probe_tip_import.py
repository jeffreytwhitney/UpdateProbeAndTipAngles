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
    def test_inserts_probe_tip_pairs(self, tmp_path):
        """Each LOADPROBE(61) + TIP(60) pair should produce one INSERT."""
        # Arrange
        prg_file = tmp_path / "PART_A.PRG"
        prg_file.write_text("fake")
        pti.TEMP_DIR = tmp_path / "pcdmis-temp"
        pti.TEMP_DIR.mkdir()

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
        pti.TEMP_DIR = tmp_path / "pcdmis-temp"
        pti.TEMP_DIR.mkdir()

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

    def test_temp_file_cleaned_up(self, tmp_path):
        """The copied .PRG file should be removed from the temp dir after processing."""
        prg_file = tmp_path / "PART_CLEAN.PRG"
        prg_file.write_text("fake")
        temp_dir = tmp_path / "pcdmis-temp"
        temp_dir.mkdir()
        pti.TEMP_DIR = temp_dir

        mock_part = _make_dmis_part("PART_CLEAN", [_make_command(0, {})])
        pti._dmis_parts = MagicMock()
        pti._dmis_parts.Open.return_value = mock_part

        with patch("DB.execute_sql_statement"):
            pti.store_tip_angles("PART_CLEAN.PRG", str(prg_file), department_id=1)

        assert not (temp_dir / "PART_CLEAN.PRG").exists()


# ---------------------------------------------------------------------------
# enumerate_pcdmis_programs
# ---------------------------------------------------------------------------

class TestEnumeratePCDMISPrograms:
    def test_new_files_are_stored(self, tmp_prg_dir, tmp_path):
        """Files not yet in the DB (count=0) should trigger store_tip_angles."""
        pti.TEMP_DIR = tmp_path / "pcdmis-temp"
        pti.TEMP_DIR.mkdir()
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

    def test_non_prg_files_ignored(self, tmp_prg_dir):
        """README.txt must never be passed to store_tip_angles."""
        with (
            patch("DB.get_sql_scalar", return_value=0),
            patch("probe_tip_import.store_tip_angles") as mock_store,
        ):
            pti.enumerate_pcdmis_programs(str(tmp_prg_dir), department_id=1)

        stored_names = {c.args[0] for c in mock_store.call_args_list}
        assert "README.txt" not in stored_names


# ---------------------------------------------------------------------------
# run_probe_and_tip_import  (orchestration)
# ---------------------------------------------------------------------------

class TestRunProbeAndTipImport:
    def _run(self, tmp_path, **kwargs):
        pti.TEMP_DIR = tmp_path / "pcdmis-temp"
        mock_app = MagicMock()
        with (
            patch("win32com.client.Dispatch", return_value=mock_app),
            patch("probe_tip_import.enumerate_pcdmis_programs"),
            patch("DB.execute_sql_statement") as mock_exec,
        ):
            pti.run_probe_and_tip_import(
                directory_path=str(tmp_path),
                department_id=1,
                **kwargs,
            )
            return mock_exec, mock_app

    def test_full_refresh_deletes_department_rows(self, tmp_path):
        mock_exec, _ = self._run(tmp_path, full_refresh=True)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("DELETE FROM tblTipAngles" in s and "DepartmentID = 1" in s
                   for s in sqls)

    def test_delete_unused_updates_then_deletes(self, tmp_path):
        mock_exec, _ = self._run(tmp_path, delete_unused=True)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("IsStillThere = 0" in s for s in sqls)
        assert any("IsStillThere = 0" in s and "DELETE" in s for s in sqls)

    def test_partial_refresh_deletes_named_program(self, tmp_path):
        mock_exec, _ = self._run(tmp_path, partial_refresh=True,
                                 probe_name="PART_A.PRG")
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("PART_A.PRG" in s and "DELETE" in s for s in sqls)

    def test_import_run_logged(self, tmp_path):
        mock_exec, _ = self._run(tmp_path)
        sqls = [c.args[0] for c in mock_exec.call_args_list]
        assert any("INSERT INTO tblImportRun" in s for s in sqls)

    def test_pcdmis_quit_called_on_success(self, tmp_path):
        _, mock_app = self._run(tmp_path)
        mock_app.Quit.assert_called_once()

    def test_pcdmis_quit_called_on_exception(self, tmp_path):
        """PC-DMIS must be shut down even when an exception is raised mid-import."""
        pti.TEMP_DIR = tmp_path / "pcdmis-temp"
        mock_app = MagicMock()
        with (
            patch("win32com.client.Dispatch", return_value=mock_app),
            patch("probe_tip_import.enumerate_pcdmis_programs",
                  side_effect=RuntimeError("boom")),
            patch("DB.execute_sql_statement"),
        ):
            with pytest.raises(RuntimeError, match="boom"):
                pti.run_probe_and_tip_import(str(tmp_path), department_id=1)

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
        prg = os.environ.get("LIVE_PRG_FILE", self.PRG_FILE)
        if not prg or not Path(prg).is_file():
            pytest.skip(
                "No .PRG file available for live test. "
                "Set LIVE_PRG_FILE env var or edit TestLivePCDMIS.PRG_FILE."
            )
        self.prg_path = Path(prg)

    def test_open_program_and_read_commands(self, tmp_path):
        """Open a real .PRG, iterate commands, confirm type 61 or 60 is present."""
        import win32com.client

        temp_dir = tmp_path / "pcdmis-temp"
        temp_dir.mkdir()
        copy_path = temp_dir / self.prg_path.name
        shutil.copy2(self.prg_path, copy_path)

        dmis_app = win32com.client.Dispatch("PCDLRN.Application")
        try:
            dmis_parts = dmis_app.PartPrograms
            dmis_part = dmis_parts.Open(str(copy_path), "OFFLINE")
            dmis_commands = dmis_part.Commands

            command_count = dmis_commands.Count
            assert command_count > 0, "Program has no commands"

            last = dmis_commands.Item(command_count)
            dmis_commands.InsertionPointAfter(last)

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

            dmis_part.Quit()
        finally:
            dmis_app.Quit()
            shutil.rmtree(temp_dir, ignore_errors=True)

