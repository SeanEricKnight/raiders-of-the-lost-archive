# Changelog

## 2026-10-02

- Reworked the source picker into an arrow-key scrolling menu.
- Added separate LOCATIONS, LOCAL DRIVES, and NETWORK DRIVES sections.
- Added mapped network-drive discovery with drive-letter and UNC target display.
- Added blank, non-selectable spacing between source-menu sections.
- Changed interactive folder selection into a drill-down navigator instead of stopping after the first child selection.
- Added an explicit `[review this folder: ...]` choice at each level; leaf folders are selected automatically.
- Preserved direct local/UNC path entry, numbered-menu fallback, and top-level-only `ExcludeFolder` behavior.

## Initial release

- Interactive source menu with Windows folder browser and common-folder shortcuts.
- File inventory with filesystem owner and document author metadata.
- Incremental CSV recovery and optional Excel output with clickable file paths.
- Folder-size worksheet and CSV with descendant totals.
- Readable, collision-safe report names based on the selected folder.
- Field report with counts, total size, and grouped issue summaries.
- Wrapped messages, keyboard folder selection, and numbered/plain-output options.
