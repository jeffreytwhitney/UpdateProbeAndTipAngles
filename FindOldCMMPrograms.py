import csv
from datetime import datetime
from pathlib import Path

departments = [
    r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Ortho Turning Approved",
    r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Ortho Mill Approved",
    r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Additive Approved",
    r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Pacing Approved",
    r"V:\Inspect Programs\CMM Programs\B_S Approved Programs\PDF Approved Programs\Cardio Approved",
    r"\\vrmss-fs1\DNC\CMM Programs\LEVEL 2 Approved Programs",
]

CUTOFF_DATE = datetime(2020, 1, 1)
REPORT_PATH = Path(__file__).resolve().parent / "old_cmm_programs.csv"


def find_old_prg_files(directory_path: str, cutoff_date: datetime) -> list[dict]:
    """Recursively find .PRG files under directory_path modified before cutoff_date."""
    results: list[dict] = []
    root = Path(directory_path)

    if not root.exists():
        print(f"[WARN] Path not found, skipping: {directory_path}")
        return results

    for entry in root.rglob("*.PRG"):
        if not entry.is_file():
            continue
        try:
            modified = datetime.fromtimestamp(entry.stat().st_mtime)
        except OSError as exc:
            print(f"[WARN] Could not read modified time for {entry}: {exc}")
            continue

        if modified < cutoff_date:
            results.append({
                "department": directory_path,
                "file_path": str(entry),
                "modified": modified.strftime("%Y-%m-%d %H:%M:%S"),
            })

    return results


def main() -> None:
    all_old_files: list[dict] = []

    for department in departments:
        print(f"Scanning: {department}")
        old_files = find_old_prg_files(department, CUTOFF_DATE)
        print(f"  Found {len(old_files)} old .PRG file(s)")
        all_old_files.extend(old_files)

    all_old_files.sort(key=lambda row: row["modified"])

    with open(REPORT_PATH, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=["department", "file_path", "modified"])
        writer.writeheader()
        writer.writerows(all_old_files)

    print("=" * 70)
    print(f"TOTAL old .PRG files found: {len(all_old_files)}")
    print(f"Report written to: {REPORT_PATH}")


if __name__ == "__main__":
    main()

