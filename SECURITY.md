# Security Policy

Excel-Dr processes local `.xlsx` files and writes cleaned copies next to the original files. It should not modify the original workbook.

## Supported Versions

Security fixes are prioritized for the latest public release.

## Reporting a Vulnerability

Please report security issues through GitHub Issues with a minimal reproduction description. If the report involves a private workbook, do not upload the workbook publicly; describe the symptoms and the workbook structure as far as possible.

Useful details:

- Excel-Dr version
- Windows version
- Whether the issue appears in single-file or folder mode
- Whether the original file was modified
- Any error message shown by the app

## Data Handling

Excel-Dr runs locally. The app does not intentionally upload workbook contents to a server.
