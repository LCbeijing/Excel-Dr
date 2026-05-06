use anyhow::{bail, Result};
use sha2::{Digest, Sha256};
use std::fs;
use std::path::{Path, PathBuf};
use zip::ZipArchive;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf()
}

fn sha256(path: &Path) -> Result<String> {
    let bytes = fs::read(path)?;
    Ok(format!("{:X}", Sha256::digest(bytes)))
}

fn assert_zip_ok(path: &Path) -> Result<()> {
    let file = fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    for index in 0..zip.len() {
        let mut member = zip.by_index(index)?;
        let mut sink = Vec::new();
        if let Err(error) = std::io::copy(&mut member, &mut sink) {
            bail!(
                "{} has a broken zip member {}: {error}",
                path.display(),
                member.name()
            );
        }
    }
    Ok(())
}

#[test]
fn clean_fixtures_do_not_need_cleanup() -> Result<()> {
    let clean_dir = repo_root().join("tests/fixtures/clean");
    if !clean_dir.exists() {
        eprintln!("No clean fixture directory found; skipping public fixture regression.");
        return Ok(());
    }
    let mut files: Vec<_> = fs::read_dir(clean_dir)?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.extension()
                .map(|ext| ext.eq_ignore_ascii_case("xlsx"))
                .unwrap_or(false)
        })
        .collect();
    files.sort();
    if files.is_empty() {
        eprintln!("No clean fixture found; skipping public fixture regression.");
        return Ok(());
    }
    for source in files {
        let before_hash = sha256(&source)?;
        assert_zip_ok(&source)?;
        let report = excel_dr_core::analyze_file(&source)?;
        if report.needs_cleanup() {
            bail!(
                "Expected clean fixture to stay untouched: {}",
                source.display()
            );
        }
        if before_hash != sha256(&source)? {
            bail!(
                "Clean fixture was modified during scan: {}",
                source.display()
            );
        }
    }
    Ok(())
}

#[test]
fn clean_file_skips_clean_fixture_without_output() -> Result<()> {
    let root = repo_root();
    let source = root.join("tests/fixtures/clean/normal_basic.xlsx");
    if !source.exists() {
        eprintln!("No clean fixture found; skipping clean-file skip regression.");
        return Ok(());
    }
    let output = root.join("tests/output/rust/normal_basic_cleaned.xlsx");
    let _ = fs::remove_file(&output);

    let report = excel_dr_core::clean_file(&source, &output)?;
    if report.needs_cleanup() {
        bail!("Expected clean fixture to be skipped");
    }
    if report.output.is_some() {
        bail!("Clean fixture should not produce an output path");
    }
    if output.exists() {
        bail!("Clean fixture should not create an output file");
    }
    Ok(())
}

#[test]
fn batch_counts_failed_files_and_keeps_failure_path() -> Result<()> {
    let root = repo_root();
    let batch_dir = root.join("tests/output/rust_bad_batch");
    if batch_dir.exists() {
        fs::remove_dir_all(&batch_dir)?;
    }
    fs::create_dir_all(&batch_dir)?;
    let invalid = batch_dir.join("invalid.xlsx");
    fs::write(&invalid, b"not a zip")?;

    let result = excel_dr_core::analyze_folder(&batch_dir)?;
    if result.file_count() != 1 {
        bail!("Expected failed file to be counted");
    }
    if result.failed.len() != 1 {
        bail!("Expected one failed file");
    }
    if result.failed[0].0.file_name() != invalid.file_name() {
        bail!("Expected failed path to keep original file name");
    }
    Ok(())
}

#[test]
fn dirty_fixture_can_be_cleaned_and_rechecked() -> Result<()> {
    let root = repo_root();
    let dirty_dir = root.join("tests/fixtures/dirty");
    let output_dir = root.join("tests/output/rust");
    fs::create_dir_all(&output_dir)?;
    if !dirty_dir.exists() {
        eprintln!("No dirty fixture directory found; skipping private dirty regression.");
        return Ok(());
    }
    let mut files: Vec<_> = fs::read_dir(dirty_dir)?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.extension()
                .map(|ext| ext.eq_ignore_ascii_case("xlsx"))
                .unwrap_or(false)
        })
        .collect();
    files.sort();
    if files.is_empty() {
        eprintln!("No dirty fixture found; skipping private dirty regression.");
        return Ok(());
    }
    for source in files {
        let before_hash = sha256(&source)?;
        let before_size = fs::metadata(&source)?.len();
        let report = excel_dr_core::analyze_file(&source)?;
        if !report.needs_cleanup() {
            bail!("Expected fixture to need cleanup: {}", source.display());
        }
        if report.suspicious_total() != 58_175 {
            bail!("Unexpected suspicious count: {}", report.suspicious_total());
        }
        if report.broken_validation_total() != 1 {
            bail!(
                "Unexpected broken validation count: {}",
                report.broken_validation_total()
            );
        }
        let output = output_dir.join(format!(
            "{}_cleaned.xlsx",
            source.file_stem().and_then(|item| item.to_str()).unwrap()
        ));
        let _ = fs::remove_file(&output);
        let cleaned = excel_dr_core::clean_file_from_report(report, &output)?;
        if !output.exists() {
            bail!("Expected output file: {}", output.display());
        }
        if before_hash != sha256(&source)? {
            bail!("Source file was modified: {}", source.display());
        }
        assert_zip_ok(&source)?;
        assert_zip_ok(&output)?;
        let verify = excel_dr_core::analyze_file(&output)?;
        if verify.needs_cleanup() {
            bail!("Cleaned output still needs cleanup: {}", output.display());
        }
        if cleaned.removed_anchors != 58_175 {
            bail!("Unexpected removed anchors: {}", cleaned.removed_anchors);
        }
        if cleaned.removed_validations != 1 {
            bail!(
                "Unexpected removed validations: {}",
                cleaned.removed_validations
            );
        }
        if before_size <= fs::metadata(&output)?.len() {
            bail!("Expected cleaned file to be smaller");
        }
    }
    Ok(())
}

#[test]
fn batch_skips_clean_files_and_cleans_actionable_files() -> Result<()> {
    let root = repo_root();
    let batch_dir = root.join("tests/fixtures/batch");
    if !batch_dir.exists() {
        eprintln!("No batch fixture directory found; skipping batch fixture regression.");
        return Ok(());
    }
    let preview = excel_dr_core::analyze_folder(&batch_dir)?;
    if preview.file_count() < 3 {
        eprintln!("Batch fixture has no private dirty file; only validating clean skip behavior.");
        if preview.actionable_count() != 0 {
            bail!(
                "Expected public batch fixtures to be clean-only, got {} actionable files",
                preview.actionable_count()
            );
        }
        return Ok(());
    }
    if preview.actionable_count() != 1 {
        bail!(
            "Expected exactly one actionable batch file, got {}",
            preview.actionable_count()
        );
    }
    let result = excel_dr_core::clean_folder_from_preview(preview)?;
    if !result.failed.is_empty() {
        bail!(
            "Expected batch cleanup without failures: {:?}",
            result.failed
        );
    }
    if result.skipped.is_empty() {
        bail!("Expected batch cleanup to skip clean files");
    }
    Ok(())
}
