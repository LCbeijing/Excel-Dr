# Excel-Dr v1.1.0

V1.1 is the MVP release after the Flutter + Rust + Rinf desktop rebuild.

## Download

- `Excel-Dr-Single.exe`
- `Excel-Dr-Flutter-portable.zip`

Note: if assets are missing, run the `Windows Release` GitHub Actions workflow with tag `v1.1.0`; it builds the Windows package on GitHub and uploads the assets from there.

## Checksums

- `Excel-Dr-Single.exe` SHA256: `6BCF8F44C0CAE0563B2563FAB9A8C4421BA4B2560A8707401B4EBC460F225AEC`
- `Excel-Dr-Flutter-portable.zip` SHA256: `46E89C72D265B54A31E8452D8704D188D1C5F441C806C8001EE3E7D8D2FF0DC4`

## Added

- Flutter + Rinf desktop shell based on the final v13 prototype.
- Rust core backend for single-file scan, single-file cleanup, batch scan, and batch cleanup.
- Windows single-file launcher: `Excel-Dr-Single.exe`.
- Flutter portable package: `Excel-Dr-Flutter-portable.zip`.
- Structured task results and progress events between Flutter and Rust.
- Rust regression tests and Flutter widget smoke test.

## Fixed And Improved

- Cleanup always writes a new file and does not overwrite the source workbook.
- Clean files are skipped during cleanup instead of producing meaningless output copies.
- Batch failures are counted and displayed with the failed file path.
- Relationship target parsing handles `/xl/...` package paths.
- Zip entry path normalization reduces unsafe extraction risk in the single-file launcher.
- The single-file launcher validates its cache and uses a mutex for concurrent startup.

## Known Limits

- Current support is focused on `.xlsx`.
- Detection focuses on abnormal hidden drawing objects and obvious `#REF!` data validation damage.
- Task cancellation and report export are planned for later versions.

## Data Safety

No real business workbook samples are committed to the public repository. Public tests skip private dirty workbook regressions when those files are absent, and synthetic clean fixtures can be generated locally when needed.
