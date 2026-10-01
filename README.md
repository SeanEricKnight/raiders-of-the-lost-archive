# Raiders of the Lost Archive

A Windows PowerShell tool for file inventories, folder-size reports, and shared cleanup reviews. Excel optional. Curses excluded.

Select a local folder or network share, inventory its files and descendants, and generate reports for deciding what to keep, move, or remove. The tool does not delete, move, or change ownership of source files.

## Requirements

- Windows with Windows PowerShell 5.1 or PowerShell 7.
- Read access to the source and write access to the report destination.
- Desktop Microsoft Excel is optional. It enables formatted XLSX reports; CSV reports are always produced after a completed scan.

## Quick start

Download `FolderReview.ps1` and save it somewhere convenient, such as `C:\PS`. It does not need to be inside the folder you want to scan.

If Windows marks the download as blocked, unblock that downloaded copy once:

```powershell
Unblock-File -LiteralPath 'C:\PS\FolderReview.ps1'
& 'C:\PS\FolderReview.ps1'
```

Follow the source menu to browse for a folder, enter a local or network path, or choose Desktop, Documents, or Downloads. Then choose where to save reports and whether to scan the entire source or one immediate child folder.

The default destination is `FolderReviewReports` on the current user's Desktop. Reports must be outside the source folder. Scanning an entire drive therefore requires a report destination on another drive or a network share.

Your organization's execution policy may impose additional restrictions. This tool does not change execution policy.

## Reports

Names use the selected folder and scan date, for example:

- `Security Templates - Review - 2026-10-01.xlsx`
- `Security Templates - Inventory - 2026-10-01.csv`
- `Security Templates - Folder Sizes - 2026-10-01.csv`
- `Security Templates - Issues - 2026-10-01.csv` when issues occur.

Repeated runs add `(1)`, `(2)`, and so on to avoid overwriting reports.

### File inventory

| Column | Contents |
| --- | --- |
| FullPath | Full file path; clickable in the XLSX report |
| LastModified | Last modification timestamp; Excel displays the date |
| FileSystemOwner | Windows filesystem owner, or an unresolved SID |
| DocumentAuthor | Document author metadata when available |
| ReviewDecision | Blank for reviewers |
| ReviewerNotes | Blank unless an owner reporting alias is used |

The XLSX inventory is sorted by containing folder and path. The CSV retains scan order and full timestamps. The tool does not update an existing review workbook or publish reports to SharePoint. Copy inventory rows into your team's workbook as needed. Copying as values removes clickable formulas.

Local and UNC file links are intended for desktop Excel on Windows with access to those locations. Excel for the web supports web-address hyperlinks only. Moving source files can invalidate links.

### Folder sizes

The `Folder Sizes` worksheet and CSV rank folders by total file size, including descendants. They include folder path, bytes, file count, and GiB (1 GiB = 1,073,741,824 bytes).

Parent and child totals overlap; do not add them together. Sizes are logical file sizes, not allocated disk space. Configured exclusions and skipped links are omitted. Scan issues cause totals to be marked partial, conservatively including metadata-only issues. Files can change while a scan runs; reports are an inventory, not a filesystem snapshot.

### Field report and issues

The terminal reports file count, folder count, issue count, and overall size. When needed, a short issue summary appears above the Issues CSV location. The CSV contains full paths and detailed errors.

## Examples

```powershell
# Interactive selection
.\FolderReview.ps1

# Scan an entire source folder
.\FolderReview.ps1 -SourceRoot 'C:\Documents' -ReportFolder 'C:\Reports' -ScanRoot

# Scan one immediate child folder
.\FolderReview.ps1 -SourceRoot '\\server\share' -FolderName 'Templates' -ReportFolder 'C:\Reports'

# Exclude named top-level folders and selected filenames
.\FolderReview.ps1 -SourceRoot '\\server\share' -ScanRoot -ReportFolder 'C:\Reports' -ExcludeFolder 'Archive' -ExcludeFile '~$*','desktop.ini','Thumbs.db','*.tmp'

# Plain terminal output with numbered selection
.\FolderReview.ps1 -PlainOutput -NumberedMenu

# Skip Windows property lookups; OpenXML author lookup still runs
.\FolderReview.ps1 -SkipWindowsProperties
```

`ExcludeFolder` applies only to immediate child folders of the source. Default file exclusions are `~$*`, `desktop.ini`, and `Thumbs.db`. Supplying `ExcludeFile` replaces that list.

Optional `OwnerAlias` substitutions affect report text only; the original owner is preserved in ReviewerNotes. Filesystem ownership is not changed.

## Stopping and recovery

Press Esc during counting or scanning when console keyboard polling is available. Completed inventory rows are flushed incrementally to the inventory CSV. Stopping during counting may leave only the CSV header because file metadata scanning has not started. Stopping does not create the completed XLSX or folder-size reports.

Ctrl+C or closing the terminal may bypass completion messages; previously flushed inventory rows remain on disk. Interrupted files or storage failure can still leave incomplete output.

## Limits

- LastModified shows modification, not whether a file has been read or remains useful.
- Author metadata is not proof of current responsibility or business ownership.
- Child directory links and file links are skipped. Windows PowerShell 5.1 conservatively skips reparse points; PowerShell 7 uses LinkTarget when available.
- Long-path and metadata support depend on the Windows/.NET host and available property handlers.
- Excel worksheet row limits apply. CSV remains available if XLSX creation fails.
- Source access failures and skipped paths are reported; review them before relying on completeness.

## Project status

Interactive scans and workbook creation have been exercised by the author on Windows. Recent revisions were checked statically during development; comprehensive automated Windows/Excel tests are not yet included.

The inventory belongs in a museum. The spreadsheet will have to do.

## License

MIT License. Copyright (c) 2026 Sean Knight. See [LICENSE](LICENSE).
