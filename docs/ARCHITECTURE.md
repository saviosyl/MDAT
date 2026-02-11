# MDAT Architecture

Detailed technical documentation of `MainForm.vb` (~1950 lines) and the overall application architecture.

---

## UI Layout Structure

```
┌─────────────────────────────────────────────────────────────┐
│  HEADER (pnlHeader, 90px)                                   │
│  ┌──────────┐  MetaMech Mechanical Design Automation        │
│  │  LOGO    │  Designed by MetaMech Solutions                │
│  └──────────┘                    [LICENSE ▾] [THEME ▾]      │
│                                  LICENCE: ACTIVE             │
│                                  PREMIUM | S:3 | 180D LEFT   │
├─────────────────────────────────────────────────────────────┤
│  [accent line - 2px]                                        │
├─────────────────────────────────────────────────────────────┤
│        [SolidWorks Year ▾]  [Select Assembly (.SLDASM)]     │
├────────┬────────────────────────────────────┬───────────────┤
│ DESIGN │                                    │ ENGINEERING   │
│ TOOLS  │         LOG PANEL                  │ TOOLS         │
│ (210px)│         (txtLog)                   │ (300px)       │
│        │                                    │               │
│ btn1   │  12:30:15  ✅ Application ready.   │ CONVEYOR CALC │
│ btn2   │  12:30:20  📁 Assembly selected:   │ PNEUMATIC     │
│ btn3   │  ...                               │ AIR CONSUMP.  │
│ ...    │                                    │ BEAM DEFLECT. │
├────────┴────────────────────────────────────┴───────────────┤
│  FOOTER: Registered Name: John Byrne Conveyors ...          │
└─────────────────────────────────────────────────────────────┘
```

### Key UI Controls
- **pnlHeader** — Top panel: logo, title, subtitle, THEME/LICENSE buttons, license status labels
- **pnlHeaderLine** — 2px accent-colored divider
- **cmbSW** — ComboBox for SolidWorks version (2022–2025)
- **btnSelectFile** — Opens file dialog for `.sldasm` selection
- **txtLog** — Multiline read-only textbox, anchored to fill center
- **pnlDesign / pnlDesignContent** — Left sidebar, scrollable, auto-populated with macro buttons
- **pnlEngineering / pnlEngineeringContent** — Right sidebar (300px wide), hardcoded 4 buttons
- **lblFooter** — Registered name at bottom

---

## Major Methods

### Initialization
| Method | Description |
|---|---|
| `New()` (constructor) | Builds all UI, loads macros from config, loads action stats, validates license, loads seat/telemetry config, ensures macro cache, flushes telemetry queue, builds buttons, applies tier locks and theme |
| `OnShown()` | Forces first header layout via `HeaderResize` |

### UI Construction
| Method | Description |
|---|---|
| `BuildHeader()` | Creates header panel with logo, title, subtitle, THEME/LICENSE buttons, license labels |
| `HeaderResize()` | Repositions THEME/LICENSE buttons and labels on resize |
| `BuildCentre()` | Creates SW version combo, assembly select button, log textbox |
| `BuildSidePanels()` | Creates left (Design Tools, 210px) and right (Engineering Tools, 300px) panels |
| `BuildFooter()` | Creates registered name label |
| `BuildButtons()` | Populates design panel from `macroDisplayMap` (slots 1–10), hardcodes 4 engineering buttons (slots 11–14) |
| `AddButton()` | Helper: creates a styled button with slot tag in a panel |

### Assembly / SolidWorks
| Method | Description |
|---|---|
| `SelectAssembly()` | Opens file dialog for `.sldasm`, stores path |
| `IsAssemblySelected()` | Returns whether path is set |
| `EnsureSolidWorksRunning()` | Attaches to running SW via `GetActiveObject` or creates new instance via `CreateObject`. Maps year → ProgID suffix (2022→30, 2023→31, etc.) |
| `IsComObjectAlive()` | Checks if COM object responds to `Visible` property |
| `TrySetSwVisible()` | Sets SW `Visible` property via reflection |
| `EnsureAssemblyOpen()` | Opens assembly via `OpenDoc`, resolves lightweight components, force rebuilds |
| `GetSWProgIDSuffix()` | Maps year string to ProgID number |

### Macro Execution Flow
| Method | Description |
|---|---|
| `RunDesignTool()` | Button click handler: validates assembly selected → checks seat → ensures SW running → ensures assembly open → calls `RunMacroSlot` |
| `RunMacroSlot(slot)` | Core execution: looks up macro path from `macroExecMap`, auto-fetches if missing (R2 delivery), logs action start with ETA, sends telemetry START, invokes `RunMacro2` then fallback `RunMacro`, logs end, sends telemetry END/FAIL, triggers cache cleanup |
| `InvokeRunMacro2()` | Calls `swApp.RunMacro2(path, module, proc, 0, 0)` via COM reflection |
| `InvokeRunMacro()` | Fallback: calls `swApp.RunMacro(path, module, proc)` via COM reflection |
| `MacroResultIsSuccess()` | Interprets COM return: Integer 0 = success, Boolean True = success |

### Macro Execution Flow Diagram
```
User clicks Design Tool button
  → RunDesignTool()
    → IsAssemblySelected()? (abort if no)
    → EnsureSeatForAction()? (abort if seat blocked)
    → EnsureSolidWorksRunning()? (attach or create COM)
    → EnsureAssemblyOpen()? (OpenDoc + resolve + rebuild)
    → RunMacroSlot(slot)
      → Lookup macro config from macroExecMap
      → If .swp file missing: EnsureMacroAvailable() (fetch from R2)
      → LogActionStart() (with ETA from historical stats)
      → SendTelemetrySafe("START", ...)
      → InvokeRunMacro2()
        → Success? → log, telemetry END, cleanup cache → return
      → InvokeRunMacro() (fallback)
        → Success? → log, telemetry END, cleanup cache → return
      → Both failed → telemetry FAIL, log error
```

### Macro Config Loading
| Method | Description |
|---|---|
| `LoadMacrosFromConfig()` | Parses `Config.txt` for `[MACROS]` section, populates `macroDisplayMap` (slot→display name) and `macroExecMap` (slot→"path\|module\|proc") |
| `HasMacrosSection()` | Checks if config file contains `[MACROS]` header |

### Macro Delivery (Cloud Fetch)
| Method | Description |
|---|---|
| `EnsureMacroAvailable()` | If .swp exists locally, use it. Otherwise fetch from `macroDeliveryUrl` via HTTP GET with Bearer token, save to cache dir |
| `EnsureMacroCacheFolder()` | Creates and optionally hides the macro cache directory |
| `EnsureTls12()` | Forces TLS 1.2 for HTTP calls |

### Cache Cleanup
| Method | Description |
|---|---|
| `StartMacroCacheCleanupRetrySilent()` | Attempts immediate cleanup; on failure, starts 2-second retry timer (up to 30 retries = ~60s) |
| `CleanupTimerTickSilent()` | Timer callback: retries cleanup |
| `TryCleanMacroCacheOnce()` | Deletes all files in cache dir (with retry per file for locked files), removes empty dirs, re-hides folder |
| `RemoveEmptyDirsSafe()` | Recursively removes empty subdirectories bottom-up |

### Action Timing
| Method | Description |
|---|---|
| `LogActionStart()` | Records start time, logs tool info, shows ETA from historical average |
| `LogActionEnd()` | Calculates elapsed time, updates running average, logs result |
| `UpdateActionStats()` | Maintains per-slot running average (count + avg seconds) |
| `GetEstimatedTimeString()` | Returns formatted ETA from stats |
| `LoadActionStats()` | Reads `action_times.txt` (format: `slot|count|avg_seconds`) |
| `SaveActionStats()` | Writes stats back to file |
| `FormatDuration()` | Formats seconds as `MM:SS` |

### Theme System
| Method | Description |
|---|---|
| `ShowThemeMenu()` | Context menu: Light / Dark / MetaMech / Custom… |
| `PickCustomTheme()` | Color picker dialog for custom accent |
| `SetTheme()` | Sets `currentTheme` enum and calls `ApplyTheme` |
| `ApplyTheme()` | Applies background/panel colors based on theme mode. Updates header line to accent color |
| `ApplyThemeToToolForm()` | Applies current theme to child engineering tool forms via `ThemeApplier.ApplyTheme()` |

#### Theme Definitions
| Theme | Background | Panel | Dark? |
|---|---|---|---|
| Light | `(245,246,248)` | `(230,232,235)` | No |
| Dark | `(26,30,36)` | `(38,42,50)` | Yes |
| MetaMech | `(242,240,248)` | `(225,220,240)` | No |
| Custom | Uses Light bg | Light panel | No (accent color only) |

### Licensing Flow
| Method | Description |
|---|---|
| `ResolveTierFromLicense()` | Reads `LicenseInfo`, sets `licenseValid`, `currentTier`, updates `lblLicence`/`lblValidity` with tier name, seat count, days remaining, color-coded warnings |
| `GetTierNameSafe()` | Maps tier int → string (0=TRIAL, 1=STANDARD, 2=PREMIUM, 3=PREMIUM PLUS) |
| `ShowLicensePopup()` | MessageBox with current license status |
| `ApplyTierLocks()` | Iterates all buttons in both panels; disables (🔒) those not allowed by `TierLocks.CanRunDesignTool`/`CanRunEngineeringTool` for current tier |
| `StyleDisabledButton()` | Grays out button, prepends 🔒, adds tooltip |

#### Licensing Flow Diagram
```
Startup
  → Licensing.GetLicenseInfo() → reads/validates license.key
  → ResolveTierFromLicense()
    → Check IsValid → set licenseValid
    → Extract tier (0–3), seats, expiry
    → Update UI labels with color-coded status
  → ApplyTierLocks()
    → For each button: TierLocks.CanRun*Tool(slot, tier)?
      → No → StyleDisabledButton (🔒 + gray + tooltip)

Before each design action:
  → EnsureSeatForAction()
    → License valid? Expired?
    → SeatEnforcer.EnsureSeatOrThrow(licenseId, tier, maxSeats)
      → Checks remote seat server (with local cache fallback)
```

### Seat Enforcement
| Method | Description |
|---|---|
| `EnsureSeatForAction()` | Pre-action gate: validates license not expired, extracts license ID via reflection, calls `SeatEnforcer.EnsureSeatOrThrow()` |
| `GetLicenseIdSafe()` | Reflection-based: tries property names (LicenseId, LICENSEID, Id, etc.) then fields |

### Config Loading
| Method | Description |
|---|---|
| `LoadSeatConfigFromConfig()` | Parses `Config.txt` for: SEAT_SERVER, SEAT_TOKEN/CLIENT_TOKEN, SYNC_URL, MACRO_DELIVERY_URL, MACRO_TOKEN, MACRO_CACHE_DIR, HIDE_MACRO_CACHE, AUTO_CLEAN_MACRO_CACHE, TELEMETRY |

### Telemetry
| Method | Description |
|---|---|
| `SendTelemetrySafe()` | Wraps `TelemetryService.SendEvent()` — sends status, slot, action name, exe version, license ID, machine name, duration, log text |
| `GetExeVersionSafe()` | Assembly version or "1.0" |
| `GetAssemblyNameSafe()` | Filename of selected assembly |

### Engineering Tools
| Method | Description |
|---|---|
| `RunEngineeringTool()` | Button click handler: maps slot 11→ConveyorCalculatorForm, 12→PneumaticCylinderCalculatorForm, 13→AirConsumptionForm, 14→placeholder. Applies theme, shows form |

### Logging
| Method | Description |
|---|---|
| `Log(msg)` | Appends timestamped line to txtLog |
| `LogA(icon, msg)` | Log with emoji prefix |

---

## Engineering Tools Available

From `app/tools/`:

| Form | Slot | Description |
|---|---|---|
| `ConveyorCalculatorForm.vb` | 11 | Conveyor configurator |
| `PneumaticCylinderCalculatorForm.vb` | 12 | Pneumatic cylinder sizing |
| `AirConsumptionForm.vb` | 13 | Air consumption & compressor sizing |
| (Beam Deflection) | 14 | Placeholder — "coming soon" |
| `TorqueCalculatorForm.vb` | — | Torque calculator |
| `MotorPowerForm.vb` | — | Motor power calculator |
| `UnitConverterForm.vb` | — | Unit converter |
| `EngineeringNotepadForm.vb` | — | Engineering notepad |
| `FlexLinkCalculatorForm.vb` | — | FlexLink conveyor calculator |
| `FlexLinkProjectConfiguratorForm.vb` | — | FlexLink project configurator |
| `ConveyorModeForm.vb` | — | Conveyor mode selector |
| `PdfExportForm.vb` | — | PDF export utility |
| `PdfQuoteForm.vb` | — | PDF quotation form |
| `AdminLoginForm.vb` | — | Admin authentication |
| `AdminMacroEditorForm.vb` | — | Macro configuration editor |
| `ProjectManager.vb` | — | Project management |

### Supporting Modules
- `Licensing.vb` / `LicenseInfo.vb` — License parsing and validation
- `LicenseGenerator*.vb` — Separate EXE for generating license keys
- `TierLocks.vb` — Feature gating per tier
- `ThemeApplier.vb` — Applies theme to child forms
- `SolidWorksGate.vb` — SolidWorks integration helper
- `MacroRegistry.vb` — Macro registration
- `ResultContext.vb` — Result passing between forms
- `UnitSystem.vb` — Unit conversion engine
- `SeatPool.vb` — Seat pool management
- `MachineId.vb` / `MachineFingerprint.vb` — Machine identification
