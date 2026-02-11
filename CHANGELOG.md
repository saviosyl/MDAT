# Changelog

## v1.0 – Baseline (2026-02-11)

### Core Application
- WinForms desktop app targeting .NET Framework 4.x (x86)
- SolidWorks COM late-binding for versions 2022–2025
- Assembly selection and auto-open with lightweight component resolution
- Central log panel with timestamped emoji-prefixed messages

### Macro Engine
- Config-driven macro slots (up to 10 design tools) loaded from `[MACROS]` section
- `RunMacro2` with `RunMacro` fallback via COM reflection
- Cloud macro delivery (R2 via Worker) with local caching
- Auto-cleanup of macro cache with retry timer (handles SolidWorks file locks)
- Action timing with per-slot running averages and estimated completion times

### Design Tool Macros Available
- PDF Automation (with merge variant)
- PDF & Index Automation
- BOM Automation (standard, Medtronic, CPS & Transitions)
- DXF Export (single, TLA)
- STEP File Automation
- Renumbering (Auto, Medtronic, CPS)
- Custom Property Update
- Quotation PDF
- Temp Change

### Engineering Tools
- Conveyor Calculator / Configurator
- Pneumatic Cylinder Calculator
- Air Consumption & Compressor Sizing
- Beam Deflection & Frame Check (placeholder)
- Motor Power Calculator
- Torque Calculator
- Unit Converter
- Engineering Notepad
- FlexLink Calculator & Project Configurator
- PDF Export & Quote Forms

### Licensing System
- RSA-signed license files (`license.key`)
- 4-tier system: Trial → Standard → Premium → Premium Plus
- Expiry tracking with warning thresholds (30d / 7d)
- Seat enforcement via remote server with local cache fallback
- Machine fingerprinting for seat binding
- License Generator (separate EXE with WinForms UI)

### Telemetry
- Event tracking: START / END / FAIL per macro action
- Local queue with flush to remote sync server
- Captures: slot, tool name, duration, assembly name, license ID, machine ID

### Theme System
- Light, Dark, MetaMech (purple), Custom (color picker)
- Theme applied to main form and child tool forms via `ThemeApplier`

### Build System
- Single `Build_EXE.bat` compiles 3 executables:
  - MDAT.exe (main app)
  - LicenseGenerator.exe (license tool)
  - PdfMergeTool.exe (PDF utility, requires PdfSharp-gdi.dll)

### Admin Tools
- Admin Login Form
- Admin Macro Editor Form
- Project Manager
