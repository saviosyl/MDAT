# MDAT – Mechanical Design Automation Tool

**By MetaMech Solutions**

MDAT is a Windows desktop application that automates SolidWorks engineering workflows — running macros for PDF generation, BOM automation, DXF export, renumbering, and more. It also includes standalone engineering calculators (conveyor, pneumatic, air consumption, torque, etc.).

---

## Architecture (6 Modules)

| Module | Description |
|---|---|
| **Main Application** (`app/MainForm.vb`) | WinForms shell — header, side panels (Design Tools / Engineering Tools), central log, theme system, license display |
| **Licensing & Seats** (`app/SeatEnforcer.vb`, `SeatServerClient.vb`, `SeatCache.vb`, `MachineFingerprint.vb`) | RSA-signed license validation, tier-based feature locks (Trial → Standard → Premium → Premium Plus), seat-count enforcement via remote server |
| **Telemetry** (`app/TelemetryClient.vb`, `TelemetryJson.vb`, `TelemetryQueue.vb`, `TelemetryService.vb`) | Usage tracking — queues events locally, flushes to sync server with action timing |
| **Macro Engine** (Config-driven + R2 delivery) | Reads `[MACROS]` from `Config.txt`, calls SolidWorks `RunMacro2`/`RunMacro` via COM late-binding, supports cloud-fetched macros with auto-cache and cleanup |
| **Engineering Tools** (`app/tools/`) | Standalone calculator forms — conveyor configurator, pneumatic cylinder, air consumption & compressor sizing, torque, unit converter, motor power, beam deflection |
| **Build System** (`build/Build_EXE.bat`) | Compiles three EXEs via `vbc.exe` (.NET Framework 4.x): MDAT.exe, LicenseGenerator.exe, PdfMergeTool.exe |

---

## Build Instructions

### Requirements
- Windows with .NET Framework 4.x installed (uses `%WINDIR%\Microsoft.NET\Framework\v4.0.30319\vbc.exe`)
- `PdfSharp-gdi.dll` in `output/` (for PdfMergeTool)

### Build
```
cd build
Build_EXE.bat
```

This compiles three executables into `output/`:
- **MDAT.exe** — Main application (x86, WinForms)
- **LicenseGenerator.exe** — License key generator (AnyCPU, WinForms)
- **PdfMergeTool.exe** — PDF merge utility (x86, console)

---

## Folder Structure

```
MDAT/
├── app/                    # Main application source
│   ├── MainForm.vb         # Primary UI (~1950 lines)
│   ├── Program.vb          # Entry point
│   ├── AboutForm.vb        # About dialog
│   ├── UITheme.vb          # Theme definitions
│   ├── SeatEnforcer.vb     # Seat-count enforcement
│   ├── SeatServerClient.vb # Remote seat server comms
│   ├── SeatCache.vb        # Local seat cache
│   ├── MachineFingerprint.vb
│   ├── Telemetry*.vb       # Telemetry pipeline
│   ├── tools/              # Engineering tool forms & utilities
│   │   ├── ConveyorCalculatorForm.vb
│   │   ├── PneumaticCylinderCalculatorForm.vb
│   │   ├── AirConsumptionForm.vb
│   │   ├── Licensing.vb / LicenseInfo.vb
│   │   ├── TierLocks.vb
│   │   ├── ThemeApplier.vb
│   │   └── ... (20+ tool files)
│   └── BACKUP/             # Old MainForm versions (cleanup candidate)
├── build/
│   └── Build_EXE.bat       # Compiler script
├── macros/                  # SolidWorks macro files (.swp)
├── tools/
│   └── PdfMergeTool.vb     # PDF merge source
├── output/                  # Build output + runtime files
│   ├── assets/              # Themes, icons, logos (tracked)
│   ├── Generate-License.ps1 # PowerShell license gen (tracked)
│   └── test_macro_server.cmd (tracked)
└── temp/                    # Transient merge/PDF files
```

---

## Security

> ⚠️ **Do NOT commit private keys, tokens, or license files to version control.**

The following are gitignored and must stay that way:
- `output/MetaMech_RSA_PRIVATE.xml` — RSA private key for license signing
- `output/Config.txt` — contains `SEAT_TOKEN`, `MACRO_TOKEN`, server URLs
- `output/license.key` — active license file
- `output/SeatConfig.txt` — seat server configuration

The **public** key (`MetaMech_RSA_PUBLIC.xml`) is safe to track.

---

## License Tiers

| Tier | Name | Access |
|------|------|--------|
| 0 | Trial | Limited design tools, limited engineering tools |
| 1 | Standard | More design tools |
| 2 | Premium | Full design tools + engineering |
| 3 | Premium Plus | Everything |

Feature gating is handled by `TierLocks.vb`. Locked buttons show 🔒.
