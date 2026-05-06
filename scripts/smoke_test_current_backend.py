from __future__ import annotations

import hashlib
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from cleaner import analyze_folder, analyze_workbook, clean_folder_from_preview, clean_workbook_from_report  # noqa: E402


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def assert_zip_ok(path: Path) -> None:
    with zipfile.ZipFile(path) as zipf:
        bad_member = zipf.testzip()
    if bad_member is not None:
        raise AssertionError(f"{path} has a broken zip member: {bad_member}")


def main() -> int:
    clean_dir = ROOT / "tests" / "fixtures" / "clean"
    dirty_dir = ROOT / "tests" / "fixtures" / "dirty"
    output_dir = ROOT / "tests" / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    clean_files = sorted(clean_dir.glob("*.xlsx")) if clean_dir.exists() else []
    if not clean_files:
        print("SKIP clean fixtures (public workbook fixtures are generated locally)")

    for source in clean_files:
        before_hash = sha256(source)
        assert_zip_ok(source)
        report = analyze_workbook(source)
        if report.needs_cleanup:
            raise AssertionError(f"Expected clean fixture to stay untouched: {source}")
        if before_hash != sha256(source):
            raise AssertionError(f"Clean fixture was modified during scan: {source}")
        print(f"OK clean {source.name}")

    dirty_files = sorted(dirty_dir.glob("*.xlsx")) if dirty_dir.exists() else []
    if not dirty_files:
        print("SKIP dirty fixtures (private samples are not committed)")

    for source in dirty_files:
        before_hash = sha256(source)
        before_size = source.stat().st_size

        report = analyze_workbook(source)
        if not report.needs_cleanup:
            raise AssertionError(f"Expected fixture to need cleanup: {source}")

        output = output_dir / f"{source.stem}_cleaned.xlsx"
        if output.exists():
            output.unlink()

        cleaned = clean_workbook_from_report(report, output)
        if not output.exists():
            raise AssertionError(f"Expected output file: {output}")

        after_hash = sha256(source)
        if before_hash != after_hash:
            raise AssertionError(f"Source file was modified: {source}")

        assert_zip_ok(source)
        assert_zip_ok(output)

        verify = analyze_workbook(output)
        if verify.needs_cleanup:
            raise AssertionError(f"Cleaned output still needs cleanup: {output}")

        saved = before_size - output.stat().st_size
        print(f"OK dirty {source.name}")
        print(f"  suspicious={report.suspicious_total}")
        print(f"  broken_validations={report.broken_validation_total}")
        print(f"  removed_anchors={cleaned.removed_anchors}")
        print(f"  removed_validations={cleaned.removed_validations}")
        print(f"  size_saved={saved}")

    batch_dir = ROOT / "tests" / "fixtures" / "batch"
    batch_dir.mkdir(parents=True, exist_ok=True)
    batch_sources = clean_files[:2] + dirty_files[:1]
    if not batch_sources:
        print("SKIP batch fixtures (public workbook fixtures are generated locally)")
        return 0
    for source in batch_sources:
        target = batch_dir / source.name
        if not target.exists() or sha256(target) != sha256(source):
            target.write_bytes(source.read_bytes())

    batch_preview = analyze_folder(batch_dir)
    if batch_preview.file_count < 3:
        if batch_preview.actionable_count != 0:
            raise AssertionError("Expected public batch fixtures to be clean-only.")
        print("OK batch folder")
        print(f"  files={batch_preview.file_count}")
        print(f"  actionable={batch_preview.actionable_count}")
        return 0
    if batch_preview.actionable_count != 1:
        raise AssertionError(f"Expected exactly one actionable batch file, got {batch_preview.actionable_count}.")

    batch_result = clean_folder_from_preview(batch_preview)
    if batch_result.failed:
        raise AssertionError(f"Expected batch cleanup without failures: {batch_result.failed}")
    if not batch_result.skipped:
        raise AssertionError("Expected batch cleanup to skip clean files.")
    print("OK batch folder")
    print(f"  files={batch_preview.file_count}")
    print(f"  actionable={batch_preview.actionable_count}")
    print(f"  skipped={len(batch_result.skipped)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
