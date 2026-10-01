# Update Probe and Tip Angles

Windows utility for reading probe and tip definitions from PC-DMIS `.PRG`
programs and synchronizing them with the `tblTipAngles` SQL Server table.
The project can process one directory or stage programs from several
departments through a shared temporary directory.

## What it does

For each `.PRG` file, the importer:

1. Opens the program offline through the PC-DMIS COM API.
2. Tracks `LOADPROBE` commands (PC-DMIS command type `61`).
3. Associates subsequent `TIP / ANGLE` commands (type `60`) with the active
   probe.
4. Inserts each department, program, probe, and tip combination into
   `tblTipAngles`.
5. Removes the source file after successful processing so files left in the
   queue can be retried.

Files that already have tip-angle records are marked as present instead of
being opened again. Non-`.PRG` files in an import queue are removed.

`FindOldCMMPrograms.py` is a separate reporting utility that recursively
finds `.PRG` files older than its configured cutoff date and writes
`old_cmm_programs.csv`.

## Requirements

- Windows
- Python 3.10 or newer
- PC-DMIS installed and available through COM
- SQL Server reachable from the machine running the importer
- Database tables used by the importer:
  - `tblTipAngles`
  - `tblTipAngle_ImportRun`
- Python packages listed in `requirements.txt`

Install the dependencies in a virtual environment:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

The `pywin32` package provides the PC-DMIS COM integration. The `pymssql`
package provides the SQL Server connection.

## Database setup

`Setup\SetupDb.sql` creates the tables required by the importer:

- `dbo.tblTipAngles` stores the department, program, probe, and tip
  relationships. `IsStillThere` defaults to `0` and is set to `1` when a
  program is found during an import.
- `dbo.tblTipAngle_ImportRun` records the start and completion time for each
  department import.

Before running the script, open `Setup\SetupDb.sql` and replace
`YOURDATABASE` in the `USE` statement with the name of the target SQL Server
database:

```sql
USE [YourDatabaseName]
```

Run the script in SQL Server Management Studio, Azure Data Studio, or another
SQL Server query tool using an account that can create tables in the target
database. The script is intended for an empty database setup; if either table
already exists, the `CREATE TABLE` statements will fail rather than alter or
delete existing data.

After the tables are created, configure the same server and database in `.env`
and verify the connection before running the importer:

```dotenv
FBCR_DB_SERVER=sql-server-host-or-instance
FBCR_DB_NAME=YourDatabaseName
```

## Configuration

Create a local `.env` file in the project directory. It is ignored by Git.
Set the following values:

```dotenv
FBCR_DB_SERVER=sql-server-host-or-instance
FBCR_DB_USER=database-user
FBCR_DB_PASSWORD=database-password
FBCR_DB_NAME=database-name
```

The importer reads these values when `DB.py` opens a connection. Existing
environment variables take precedence over values in `.env`. Do not commit
`.env` or database credentials.

## Running the importer

The executable workflow is currently configured in the `__main__` block of
`probe_tip_import.py`. Review the department paths, IDs, and refresh flags
before running it:

```powershell
python .\probe_tip_import.py
```

The default multi-department workflow:

- Opens one PC-DMIS session.
- Clears the temporary staging directory before each department.
- Copies `.PRG` files recursively into the staging directory.
- Imports each department with its `DepartmentID`.
- Prints per-department and total added/updated counts.
- Shuts down PC-DMIS in a `finally` block.

The default paths are site-specific network paths. Change them to paths
available on the target machine before use. The temporary directory must be
writable, for example `C:\pcdmis-temp`.

### Single-directory use

Applications or scripts can call the orchestration function directly:

```python
from probe_tip_import import run_probe_and_tip_import

counts = run_probe_and_tip_import(
    directory_path=r"C:\pcdmis-temp",
    department_id=15,
    delete_unused=True,
)
print(counts)
```

The returned dictionary contains `newly_added` and `updated` counts.

## Refresh options

The importer supports these mutually composable controls:

| Option | Behavior |
| --- | --- |
| `delete_unused=True` | Marks all department records absent, then removes records that were not encountered during this run. |
| `full_refresh=True` | Deletes all existing tip-angle records for the department before importing. |
| `partial_refresh=True` with `probe_name` | Deletes records for the named program before importing it again. |
| `manage_session=False` | Reuses a PC-DMIS session owned by a multi-department import. |

Use `delete_unused` carefully: the queue must contain a complete snapshot of
the department, or valid records that are not in the snapshot will be deleted.

Each import run is recorded in `tblTipAngle_ImportRun`. An existing open run
for the department is reused; a completed run receives an `EndTime`.

## Finding old programs

Update the department paths and `CUTOFF_DATE` in `FindOldCMMPrograms.py`, then
run:

```powershell
python .\FindOldCMMPrograms.py
```

The script prints progress and writes a CSV report beside the script at
`old_cmm_programs.csv`. Missing directories and files whose timestamps cannot
be read are reported as warnings and skipped.

## Logging and troubleshooting

When a managed import starts, `probe_tip_import.py` recreates
`crash_log.txt` with the run start time, loaded module, Python executable,
working directory, and log path. Each program is logged immediately before it
is opened. This makes it possible to identify the program involved if
PC-DMIS or the COM layer terminates unexpectedly.

If processing a program fails, its source file is not deleted. Resolve the
PC-DMIS, database, or file-path problem and retry the remaining queue.

Common checks:

- Confirm the `.PRG` path is readable and the temporary directory is writable.
- Confirm PC-DMIS can be launched by the current Windows user.
- Confirm the SQL Server host, database, and credentials in `.env`.
- Check `crash_log.txt` for the last program opened.
- Run a non-destructive import without `delete_unused` before enabling cleanup.

## Tests

Run the unit tests with:

```powershell
pytest
```

The normal test suite mocks PC-DMIS and database calls. A separate live test
requires PC-DMIS and a real program file:

```powershell
$env:LIVE_PRG_FILE = 'C:\path\to\program.PRG'
pytest .\tests -v -m live
```

The live test opens the file offline and verifies that it contains at least
one probe or tip command. It should only be run on a machine with PC-DMIS
installed and should use a copy of any production program.

## Project layout

| Path | Purpose |
| --- | --- |
| `probe_tip_import.py` | PC-DMIS import and multi-department orchestration |
| `DB.py` | SQL Server connections and query helpers |
| `Setup\SetupDb.sql` | SQL Server table-creation script |
| `FindOldCMMPrograms.py` | Old `.PRG` file report generator |
| `tests/` | Unit and optional live integration tests |
| `requirements.txt` | Python dependencies |
| `crash_log.txt` | Runtime log recreated at the start of managed imports |
| `old_cmm_programs.csv` | Generated old-program report |
