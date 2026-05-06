from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "rust" / "excel_dr_core" / "Cargo.toml"
EXE = ROOT / "rust" / "excel_dr_core" / "target" / "release" / "excel_dr_core.exe"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def assert_zip_ok(path: Path) -> None:
    with zipfile.ZipFile(path) as zipf:
        bad_member = zipf.testzip()
    if bad_member is not None:
        raise AssertionError(f"{path} has a broken zip member: {bad_member}")


def run_core(*args: str) -> dict:
    if EXE.exists():
        command = [str(EXE), "--json", *args]
    else:
        command = ["cargo", "run", "--manifest-path", str(MANIFEST), "--quiet", "--", "--json", *args]
    completed = subprocess.run(command, cwd=ROOT, check=True, text=True, capture_output=True, encoding="utf-8")
    return json.loads(completed.stdout)


def main() -> int:
    clean_dir = ROOT / "tests" / "fixtures" / "clean"
    dirty_dir = ROOT / "tests" / "fixtures" / "dirty"
    output_dir = ROOT / "tests" / "output" / "rust_smoke"
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)

    clean_files = sorted(clean_dir.glob("*.xlsx")) if clean_dir.exists() else []
    if not clean_files:
        print("SKIP rust clean fixtures (public workbook fixtures are generated locally)")
    for source in clean_files:
        before_hash = sha256(source)
        assert_zip_ok(source)
        report = run_core("analyze-file", str(source))
        if report["sheet_plans"] and (sum(d["suspicious_total"] for s in report["sheet_plans"] for d in s["drawing_plans"]) > 0):
            raise AssertionError(f"Expected clean fixture to stay untouched: {source}")
        if sum(len(s["broken_validations"]) for s in report["sheet_plans"]) > 0:
            raise AssertionError(f"Expected no broken validations: {source}")
        if before_hash != sha256(source):
            raise AssertionError(f"Clean fixture was modified during scan: {source}")
        print(f"OK rust clean {source.name}")

    dirty_files = sorted(dirty_dir.glob("*.xlsx")) if dirty_dir.exists() else []
    if not dirty_files:
        print("SKIP rust dirty fixtures (private samples are not committed)")
    for source in dirty_files:
        before_hash = sha256(source)
        before_size = source.stat().st_size
        report = run_core("analyze-file", str(source))
        suspicious = sum(d["suspicious_total"] for s in report["sheet_plans"] for d in s["drawing_plans"])
        broken = sum(len(s["broken_validations"]) for s in report["sheet_plans"])
        if suspicious != 58175:
            raise AssertionError(f"Expected 58175 suspicious anchors, got {suspicious}")
        if broken != 1:
            raise AssertionError(f"Expected 1 broken validation, got {broken}")

        output = output_dir / f"{source.stem}_cleaned.xlsx"
        cleaned = run_core("clean-file", str(source), "--output", str(output))
        if before_hash != sha256(source):
            raise AssertionError(f"Source file was modified: {source}")
        assert_zip_ok(source)
        assert_zip_ok(output)
        verify = run_core("analyze-file", str(output))
        verify_suspicious = sum(d["suspicious_total"] for s in verify["sheet_plans"] for d in s["drawing_plans"])
        verify_broken = sum(len(s["broken_validations"]) for s in verify["sheet_plans"])
        if verify_suspicious or verify_broken:
            raise AssertionError(f"Cleaned output still needs cleanup: {output}")
        print(f"OK rust dirty {source.name}")
        print(f"  suspicious={suspicious}")
        print(f"  broken_validations={broken}")
        print(f"  removed_anchors={cleaned['removed_anchors']}")
        print(f"  removed_validations={cleaned['removed_validations']}")
        print(f"  size_saved={before_size - output.stat().st_size}")

    batch_dir = ROOT / "tests" / "fixtures" / "batch"
    if not batch_dir.exists():
        print("SKIP rust batch fixtures (public workbook fixtures are generated locally)")
        return 0
    batch = run_core("analyze-folder", str(batch_dir))
    actionable = 0
    for report in batch["reports"]:
        suspicious = sum(d["suspicious_total"] for s in report["sheet_plans"] for d in s["drawing_plans"])
        broken = sum(len(s["broken_validations"]) for s in report["sheet_plans"])
        if suspicious or broken:
            actionable += 1
    if len(batch["reports"]) < 3:
        if actionable != 0:
            raise AssertionError(f"Expected public batch fixtures to be clean-only, got {actionable}.")
    elif actionable != 1:
        raise AssertionError(f"Expected exactly one actionable batch file, got {actionable}.")
    print("OK rust batch folder")
    print(f"  files={len(batch['reports'])}")
    print(f"  actionable={actionable}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
