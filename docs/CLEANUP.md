# Cleanup Candidates

Files that appear to be backups, duplicates, or leftover artifacts that could be removed.

---

## `app/BACKUP/` — Old MainForm Versions

| File | Notes |
|---|---|
| `app/BACKUP/MainForm.vb` | Backup of MainForm |
| `app/BACKUP/1MainForm.vb` | Numbered backup |
| `app/BACKUP/MainForm.txt` | Text copy of MainForm |
| `app/BACKUP/MainForm - old.txt` | Older text copy |
| `app/BACKUP/New folder/MainForm.vb` | Another backup in subfolder |
| `app/BACKUP/New folder/MainForm.txt` | Another text copy in subfolder |

**Recommendation:** Archive or delete entire `app/BACKUP/` directory. Already gitignored via `app/tools/backup/` but this folder is at `app/BACKUP/` (different path — consider adding to `.gitignore` or deleting).

---

## `app/tools/backup/` — Old Tool Backups

| File | Notes |
|---|---|
| `app/tools/backup/MacroEngine.vb` | Old macro engine |
| `app/tools/backup/SolidWorksMacroRunner.vb` | Old macro runner |
| `app/tools/backup/MacroAutoDetect.vb` | Old macro auto-detection |
| `app/tools/backup/MacroSlot.vb` | Old macro slot class |
| `app/tools/backup/LicenseInfoForm.vb` | Old license info form |
| `app/tools/backup/FlexLinkProjectConfiguratorForm.vb` | Old FlexLink configurator |
| `app/tools/backup/ConveyorModeForm.vb` | Old conveyor mode form |
| `app/tools/backup/ConveyorCalculatorForm.vb` | Old conveyor calculator |
| `app/tools/backup/FlexLinkCalculatorForm.txt` | Text copy |
| `app/tools/backup/New folder/ConveyorCalculatorForm.vb` | Another copy |
| `app/tools/backup/Program testing/build.bat` | Test build script |
| `app/tools/backup/Program testing/Program.vb` | Test program |
| `app/tools/backup/Program testing/ConveyorCalculatorForm.vb` | Test conveyor form |

**Recommendation:** Already gitignored. Safe to delete if no longer needed for reference.

---

## Duplicate/Redundant Source Files

| File | Duplicate Of | Notes |
|---|---|---|
| `app/tools/MainForm_OLD1.vb` | `app/MainForm.vb` | Old version of MainForm stored in tools directory |
| `app/tools/SeatServerClient.vb` | `app/SeatServerClient.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/SeatCache.vb` | `app/SeatCache.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/MachineFingerprint.vb` | `app/MachineFingerprint.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/TelemetryJson.vb` | `app/TelemetryJson.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/TelemetryQueue.vb` | `app/TelemetryQueue.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/TelemetryService.vb` | `app/TelemetryService.vb` | Duplicate — only `app/` version is compiled |
| `app/tools/PdfMergeTool.vb` | `tools/PdfMergeTool.vb` | Duplicate — only `tools/` version is compiled |
| `app/tools/SeatEnforcer.vb` | `app/SeatEnforcer.vb` | Duplicate — only `app/` version is compiled |

**Recommendation:** Delete all `app/tools/` duplicates. The build script (`Build_EXE.bat`) only references files in `app/` and `app/tools/` by specific name — the duplicates listed above are NOT in the build and could cause confusion.

---

## Miscellaneous Junk Files

| File | Notes |
|---|---|
| `app/tools/New Text Document.txt` | Empty or scratch file |
| `app/tools/New Text Document (2).txt` | Empty or scratch file |

**Recommendation:** Delete.

---

## Possibly Unused Tool Files

These are in `app/tools/` but NOT referenced in `Build_EXE.bat`. They may be compiled but unused, or were removed from the build:

| File | Notes |
|---|---|
| `app/tools/MacroRegistry.vb` | Not in build script |
| `app/tools/FlexLinkProjectConfiguratorForm.vb` | Not in build script |
| `app/tools/FlexLinkCalculatorForm.vb` | Not in build script |
| `app/tools/PneumaticCylinderToolForm.vb` | Not in build script (but PneumaticCylinderCalculatorForm.vb IS) |
| `app/tools/Unit Converter.vb` | Not in build (UnitConverterForm.vb IS) — possible old version |
| `app/tools/Engineering Calculator.vb` | In build — but has space in filename (works but unusual) |
| `app/tools/LicenseGenerator.vb` | Not in build (separate LicenseGenerator*.vb files are) |

**Recommendation:** Verify if these are needed. If not compiled, they're dead code.

---

## `temp/` Directory

| File | Notes |
|---|---|
| `temp/merge_order.txt` | Runtime artifact |
| `temp/Combined.pdf` | Runtime artifact |
| `temp/a.pdf`, `temp/b.pdf` | Test PDFs |

**Recommendation:** Already gitignored. Clear contents periodically.

---

## Summary

| Category | Count | Action |
|---|---|---|
| Backup directories | 2 dirs (~15 files) | Delete or archive |
| Duplicate source files | 9 files | Delete from `app/tools/` |
| Junk text files | 2 files | Delete |
| Possibly unused tools | ~7 files | Audit and remove if dead |
| **Total cleanup candidates** | **~33 files** | |
