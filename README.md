# Resource Audit Tool

PowerShell-based inventory and storage audit utility for Windows application environments.

This project is a sanitized public version of an internal support idea originally used to gather system and storage information for healthcare application environments. The public version keeps the useful operational concepts while removing customer-specific logic, internal paths, and embedded credentials.

## What it does

- collects host metadata
- identifies physical vs virtual machine type
- inventories fixed drives and free space
- optionally measures selected directory trees
- discovers IIS websites when available
- detects common database engines installed locally
- exports results to CSV and JSON for review
- optionally exports a multi-sheet Excel workbook when `ImportExcel` is installed

## Why it is useful

This tool is designed for support, implementation, and operations work where you need a fast view of environment health and storage footprint.

Typical use cases:

- onboarding and environment reviews
- migration planning
- support escalations
- storage cleanup efforts
- readiness checks before implementation work

Questions it helps answer:

- How much free disk space is available?
- Which application directories are consuming the most storage?
- Is this server virtual or physical?
- Is IIS installed and which sites exist?
- Which common database engines appear to be present?

## Why this matters professionally

This project demonstrates practical infrastructure and support engineering work rather than toy scripting. It shows the ability to turn a real operations need into a reusable audit utility that can help with troubleshooting, implementation planning, and system reviews.

## Files

- `Start-ResourceAudit.ps1`
  Main script that gathers the audit data and exports reports.
- `.gitignore`
  Ignores generated output.
- `sample-output/`
  Generic sample reports for GitHub visitors.

## Example

```powershell
.\Start-ResourceAudit.ps1 `
  -OutputRoot "$env:USERPROFILE\Documents\ResourceAudit" `
  -AuditPaths "C:\inetpub","C:\AppData" `
  -ExportExcel
```

## Output

The script writes:

- `resource_audit_summary_<date>.json`
- `resource_audit_drives_<date>.csv`
- `resource_audit_paths_<date>.csv`
- `resource_audit_sites_<date>.csv`
- `resource_audit_databases_<date>.csv`
- `resource_audit_<date>.xlsx` when `-ExportExcel` is used and the `ImportExcel` module is available

## Sample output

See the `sample-output` folder for generic example reports included for demonstration.

## Notes

- No credentials are stored in the script.
- No internal customer names or product-specific paths are included.
- Directory sizing is recursive, so very large paths may take time.
- Excel export is optional and only runs when the `ImportExcel` PowerShell module is installed.
