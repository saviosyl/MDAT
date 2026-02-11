# MDAT UI Overhaul Plan

> Comprehensive visual upgrade for MDAT — VB.NET 4.0 WinForms, no external libraries.
> All changes are pure cosmetic/structural unless noted. No WPF, no NuGet, no new dependencies.

---

## Table of Contents

1. [Premium Light Theme (Default)](#1-premium-light-theme-default)
2. [Premium Dark Theme](#2-premium-dark-theme)
3. [Theme System Improvements](#3-theme-system-improvements)
4. [Header Overhaul](#4-header-overhaul)
5. [Macro Button Panel (Design Automation)](#5-macro-button-panel-design-automation)
6. [Engineering Tools Panel](#6-engineering-tools-panel)
7. [Log Window](#7-log-window)
8. [Footer](#8-footer)
9. [Calculator Forms Standardization](#9-calculator-forms-standardization)
10. [Splash Screen](#10-splash-screen)
11. [Risk Assessment Summary](#11-risk-assessment-summary)
12. [Implementation Order](#12-implementation-order)

---

## Color Reference

### Brand Colors
| Name | Hex | RGB | Usage |
|------|-----|-----|-------|
| MetaMech Teal | `#00B4FF` | `(0, 180, 255)` | Primary accent, interactive elements |
| MetaMech Cyan | `#00E5D0` | `(0, 229, 208)` | Dark theme highlights, success |
| Navy Base | `#0B1E34` | `(11, 30, 52)` | Dark theme main background |
| Navy Panel | `#0F2A44` | `(15, 42, 68)` | Dark theme panel background |
| Navy Card | `#162C45` | `(22, 44, 69)` | Dark theme card/elevated surface |

### Status Colors (Both Themes)
| Status | Hex | RGB |
|--------|-----|-----|
| Connected/Success | `#00C853` | `(0, 200, 83)` |
| Warning | `#FFB300` | `(255, 179, 0)` |
| Error | `#F44336` | `(244, 67, 54)` |
| Info | `#00B4FF` | `(0, 180, 255)` |

---

## 1. Premium Light Theme (Default)

**Risk: LOW** — Cosmetic color changes only.

### Color Palette

| Element | Hex | RGB | Notes |
|---------|-----|-----|-------|
| Form Background | `#F8F9FA` | `(248, 249, 250)` | Clean off-white |
| Panel Background | `#FFFFFF` | `(255, 255, 255)` | Pure white cards |
| Header Background | `#FFFFFF` | `(255, 255, 255)` | Clean white header |
| Side Panel Background | `#F1F3F5` | `(241, 243, 245)` | Subtle gray distinction |
| Card/Button Surface | `#FFFFFF` | `(255, 255, 255)` | White cards with border |
| Card Border | `#DEE2E6` | `(222, 226, 230)` | Subtle separator |
| Card Border Hover | `#00B4FF` | `(0, 180, 255)` | Teal highlight on hover |
| Primary Text | `#1A1A2E` | `(26, 26, 46)` | Near-black, high contrast |
| Secondary Text | `#6C757D` | `(108, 117, 125)` | Muted/subtitle text |
| Accent Primary | `#00B4FF` | `(0, 180, 255)` | Active states, progress, links |
| Accent Hover | `#0096D6` | `(0, 150, 214)` | Darker teal for hover |
| Button Background | `#FFFFFF` | `(255, 255, 255)` | Flat white + border |
| Button Hover BG | `#E8F7FF` | `(232, 247, 255)` | Light teal wash |
| Selected Button BG | `#00B4FF` | `(0, 180, 255)` | Solid teal fill |
| Selected Button Text | `#FFFFFF` | `(255, 255, 255)` | White on teal |
| Disabled Button BG | `#F1F3F5` | `(241, 243, 245)` | Gray wash |
| Disabled Button Text | `#ADB5BD` | `(173, 181, 189)` | Muted |
| Header Line | `#00B4FF` | `(0, 180, 255)` | 2px accent separator |
| Footer BG | `#F1F3F5` | `(241, 243, 245)` | Subtle distinction |

### Font Specifications
| Element | Font | Size | Style |
|---------|------|------|-------|
| App Title | Segoe UI | 18pt | Bold |
| Subtitle | Segoe UI | 9pt | Regular |
| Section Headers | Segoe UI | 10pt | Bold |
| Button Text | Segoe UI | 9.5pt | SemiBold (use Bold) |
| Body/Labels | Segoe UI | 9pt | Regular |
| Log Text | Consolas | 9pt | Regular |
| Footer | Segoe UI | 8pt | Regular |
| Version Badge | Segoe UI | 8pt | Regular |

### Card Shadow Simulation (WinForms Compatible)
WinForms doesn't support CSS-like shadows. Simulate depth with:
- 1px border using `#DEE2E6` around cards
- For "elevated" cards: draw a 1px `#CED4DA` line on bottom and right edges via `Paint` event
- Use `Panel` with `BorderStyle = None` and custom `Paint` handler:

```vb
Private Sub PaintCardShadow(sender As Object, e As PaintEventArgs)
    Dim p As Panel = DirectCast(sender, Panel)
    Dim r As Rectangle = New Rectangle(0, 0, p.Width - 1, p.Height - 1)
    ' Border
    Using pen As New Pen(Color.FromArgb(222, 226, 230))
        e.Graphics.DrawRectangle(pen, r)
    End Using
    ' Bottom shadow line
    Using pen As New Pen(Color.FromArgb(206, 212, 218))
        e.Graphics.DrawLine(pen, 2, p.Height - 1, p.Width - 1, p.Height - 1)
        e.Graphics.DrawLine(pen, p.Width - 1, 2, p.Width - 1, p.Height - 1)
    End Using
End Sub
```

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` | Update `BG_LIGHT`, `PANEL_LIGHT` constants; update `ApplyTheme()` to apply new palette to all controls |
| `output/assets/themes/light.theme` | Complete rewrite with new palette |
| `app/UITheme.vb` | Update default colors to match premium light palette |

---

## 2. Premium Dark Theme

**Risk: LOW** — Cosmetic color changes only.

### Color Palette

| Element | Hex | RGB | Notes |
|---------|-----|-----|-------|
| Form Background | `#0B1E34` | `(11, 30, 52)` | Navy base (matches website) |
| Panel Background | `#0F2A44` | `(15, 42, 68)` | Elevated navy |
| Header Background | `#0F2A44` | `(15, 42, 68)` | Navy panel |
| Card Surface | `#162C45` | `(22, 44, 69)` | Card background |
| Card Border | `#1E3A5F` | `(30, 58, 95)` | Subtle navy border |
| Card Border Hover | `#00B4FF` | `(0, 180, 255)` | Blue highlight |
| Primary Text | `#EAF4FF` | `(234, 244, 255)` | Off-white, easy on eyes |
| Secondary Text | `#A9C7E8` | `(169, 199, 232)` | Muted blue-gray |
| Accent Primary | `#00E5D0` | `(0, 229, 208)` | Teal highlights |
| Accent Interactive | `#00B4FF` | `(0, 180, 255)` | Buttons, links |
| Accent Hover | `#33C4FF` | `(51, 196, 255)` | Lighter blue hover |
| Button Background | `#162C45` | `(22, 44, 69)` | Navy card surface |
| Button Hover BG | `#1A3555` | `(26, 53, 85)` | Slightly lighter |
| Selected Button BG | `#00B4FF` | `(0, 180, 255)` | Solid blue |
| Selected Button Text | `#FFFFFF` | `(255, 255, 255)` | White |
| Disabled Button BG | `#0D2236` | `(13, 34, 54)` | Darker navy |
| Disabled Button Text | `#4A6A8A` | `(74, 106, 138)` | Muted |
| Header Line | `#00E5D0` | `(0, 229, 208)` | Teal separator |
| Footer BG | `#0A1929` | `(10, 25, 41)` | Darkest navy |

### Accessibility
- Primary text on navy BG: contrast ratio ~12:1 ✅ (WCAG AAA)
- Secondary text on navy BG: ~7:1 ✅ (WCAG AA)
- Accent teal on navy BG: ~8:1 ✅ (WCAG AA)

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` | Update `BG_DARK`, `PANEL_DARK` constants to navy palette |
| `output/assets/themes/dark.theme` | Already close — update `card` and add button colors |
| `app/tools/ThemeApplier.vb` | Update dark button default from `(40,45,58)` to `(22,44,69)` |

---

## 3. Theme System Improvements

**Risk: MEDIUM** — Restructuring theme persistence and child form theming. No logic changes.

### Current State
- Theme set via `ContextMenuStrip` from THEME button
- Themes only partially apply — `ApplyTheme()` in MainForm only changes `BackColor` of form, header, and header line
- Child forms get themed via `ApplyThemeToToolForm()` → `ThemeApplier.ApplyTheme()`
- No persistence — reverts to Light on restart
- Theme files (`dark.theme`, `light.theme`) exist but are NOT loaded at runtime

### Proposed Changes

#### 3a. Theme Persistence via `ui.settings`
Create a simple INI-style settings file at `output/ui.settings`:

```ini
[UI]
THEME=light
CUSTOM_ACCENT=#00B4FF
```

**New class: `app/tools/UISettings.vb`**
```vb
Public Module UISettings
    Private Const SETTINGS_FILE As String = "ui.settings"
    
    Public Function LoadThemeName() As String
        ' Read THEME= from ui.settings, return "light"/"dark"/"metamech"/"custom"
        ' Default: "light"
    End Function
    
    Public Sub SaveThemeName(name As String)
        ' Write THEME= to ui.settings
    End Sub
    
    Public Function LoadCustomAccent() As Color
        ' Read CUSTOM_ACCENT= hex, parse to Color
    End Function
    
    Public Sub SaveCustomAccent(c As Color)
        ' Write CUSTOM_ACCENT= as hex
    End Sub
End Module
```

#### 3b. Theme Switcher UI
Replace `ContextMenuStrip` with a styled `ComboBox` in the header:

```vb
Private cmbTheme As ComboBox  ' replaces btnTheme

' In BuildHeader():
cmbTheme = New ComboBox() With {
    .DropDownStyle = ComboBoxStyle.DropDownList,
    .Size = New Size(100, 28),
    .Font = New Font("Segoe UI", 8.5F),
    .FlatStyle = FlatStyle.Flat
}
cmbTheme.Items.AddRange(New String() {"Light", "Dark"})
cmbTheme.SelectedIndex = 0
AddHandler cmbTheme.SelectedIndexChanged, AddressOf OnThemeChanged
pnlHeader.Controls.Add(cmbTheme)
```

> **Note:** Remove MetaMech and Custom themes. Simplify to Light/Dark only for a cleaner UX. Custom accent can be a future feature.

#### 3c. Full-Form Theme Application
Expand `ApplyTheme()` in MainForm to recursively theme ALL controls (use `ThemeApplier.ApplyTheme()` on `Me`):

```vb
Private Sub ApplyTheme()
    Dim bg, panel, accent As Color
    Dim isDark As Boolean
    
    Select Case currentTheme
        Case ThemeMode.PANEL_LIGHT
            bg = Color.FromArgb(248, 249, 250)
            panel = Color.FromArgb(255, 255, 255)
            accent = Color.FromArgb(0, 180, 255)
            isDark = False
        Case ThemeMode.PANEL_DARK
            bg = Color.FromArgb(11, 30, 52)
            panel = Color.FromArgb(15, 42, 68)
            accent = Color.FromArgb(0, 229, 208)
            isDark = True
    End Select
    
    ThemeApplier.ApplyTheme(Me, bg, panel, accent, isDark)
    
    ' Override specific elements
    pnlHeaderLine.BackColor = accent
    ' ... header-specific styling
End Sub
```

#### 3d. Load Theme Files at Runtime (Future)
Currently the `.theme` files are not parsed. Add a `ThemeFileParser` module that reads INI-style theme files into a `ThemeColors` structure. This enables user-created themes without recompilation.

**New file: `app/tools/ThemeFileParser.vb`** (future, not MVP)

### Files to Change/Create
| File | Action | Risk |
|------|--------|------|
| `app/tools/UISettings.vb` | **NEW** — theme persistence | LOW |
| `app/MainForm.vb` | Rewrite `ApplyTheme()`, replace `btnTheme` with `cmbTheme`, load saved theme on startup | MEDIUM |
| `app/tools/ThemeApplier.vb` | Update dark-mode button colors to navy palette | LOW |
| `output/assets/themes/light.theme` | Rewrite with new palette | LOW |
| `output/assets/themes/dark.theme` | Update to match navy palette | LOW |

---

## 4. Header Overhaul

**Risk: LOW** — Visual enhancements only. Layout positions are LOCKED.

### What Is LOCKED (Cannot Move)
Per user's rule, the header layout structure is locked:
- **Logo:** Left side, `(20, 10)`, 160×70
- **Title:** Right of logo, `(picLogo.Right + 15, 20)`
- **Subtitle:** Below title, `(picLogo.Right + 15, 55)`
- **Theme control:** Top-right area (repositioned by `HeaderResize`)
- **License button:** Left of theme control
- **License labels:** Below buttons, right-aligned

### What CAN Be Enhanced Visually

#### 4a. Header Panel Height
Increase from 90px to **100px** for breathing room.

#### 4b. Header Background
- Light theme: Pure white `#FFFFFF` with a subtle bottom border
- Dark theme: Navy panel `#0F2A44`

#### 4c. Title Styling
- Keep font: `Segoe UI 18pt Bold`
- Light theme: `#1A1A2E` text
- Dark theme: `#EAF4FF` text

#### 4d. Subtitle Enhancement
- Change to `Segoe UI 9pt Regular` (remove italic — looks more professional)
- Color: `#6C757D` (light) / `#A9C7E8` (dark)

#### 4e. Version Badge
Add a small version label near the subtitle:
```vb
lblVersion = New Label() With {
    .Text = "v" & GetExeVersionSafe(),
    .Font = New Font("Segoe UI", 7.5F),
    .ForeColor = Color.FromArgb(108, 117, 125),
    .AutoSize = True,
    .Location = New Point(lblSubtitle.Right + 12, 58)
}
pnlHeader.Controls.Add(lblVersion)
```

#### 4f. Header Accent Line
Keep the 2px `pnlHeaderLine` but ensure it uses the theme accent:
- Light: `#00B4FF`
- Dark: `#00E5D0`

#### 4g. License Labels Styling
- Valid license: Green `#00C853`
- Expiring (<30 days): Amber `#FFB300`
- Expired/Invalid: Red `#F44336`
- Already implemented — just ensure colors apply in both themes

#### 4h. Theme Switcher & License Button Styling
Style as flat buttons with accent border:
```vb
btnTheme.FlatStyle = FlatStyle.Flat
btnTheme.FlatAppearance.BorderSize = 1
btnTheme.FlatAppearance.BorderColor = Color.FromArgb(0, 180, 255)
btnTheme.BackColor = Color.Transparent
btnTheme.ForeColor = Color.FromArgb(26, 26, 46) ' dark text for light theme
btnTheme.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
```

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` `BuildHeader()` | Visual styling only — colors, fonts, version label |
| `app/MainForm.vb` `ApplyTheme()` | Theme header controls properly |

---

## 5. Macro Button Panel (Design Automation)

**Risk: MEDIUM** — Restructuring button rendering. Same click behavior.

### Current State
- `AddButton()` creates flat `Button` controls, 34px tall, purple border `(150,90,190)`, dark background `(40,45,58)`
- No hover effects, no selected state, no status indicators
- Buttons are plain text only

### Proposed: Card-Style Buttons

#### 5a. New Button Dimensions
- Height: **52px** (up from 34px) — room for label + subtitle
- Width: `parent.Width - 24` (12px margin each side)
- Margin between cards: **8px** (up from 6px)

#### 5b. Card Layout Per Button
```
┌─────────────────────────────────────────┐
│  [●]  Tool Name                   ~12s  │
│       Slot 1 • Design Tool              │
└─────────────────────────────────────────┘
```
- Left 8px: Status dot (● using a small painted circle — green if last run succeeded, gray if never run)
- Center: Tool name (Segoe UI 9.5pt Bold)
- Below name: Subtitle line (Segoe UI 7.5pt, muted color) showing slot info
- Right: Average run time badge (e.g., "~12s") in small muted text

#### 5c. Implementation — Custom Panel Instead of Button
Use a `Panel` per card with child labels, and handle `Click`/`MouseEnter`/`MouseLeave`:

```vb
Private Sub AddCardButton(parent As Panel, title As String, slot As Integer, y As Integer, clickHandler As EventHandler)
    Dim card As New Panel() With {
        .Location = New Point(12, y),
        .Size = New Size(parent.Width - 24, 52),
        .BackColor = Color.White, ' themed
        .Cursor = Cursors.Hand,
        .Tag = slot
    }
    card.BorderStyle = BorderStyle.None
    AddHandler card.Paint, AddressOf PaintCardBorder
    
    ' Status dot — painted in Paint handler
    
    Dim lblName As New Label() With {
        .Text = title,
        .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
        .ForeColor = Color.FromArgb(26, 26, 46),
        .Location = New Point(28, 8),
        .AutoSize = True,
        .BackColor = Color.Transparent,
        .Cursor = Cursors.Hand
    }
    card.Controls.Add(lblName)
    
    ' ETA badge
    Dim eta As String = GetEstimatedTimeString(slot)
    If eta <> "" Then
        Dim lblEta As New Label() With {
            .Text = eta,
            .Font = New Font("Segoe UI", 7.5F),
            .ForeColor = Color.FromArgb(108, 117, 125),
            .AutoSize = True,
            .Anchor = AnchorStyles.Right Or AnchorStyles.Top,
            .BackColor = Color.Transparent
        }
        lblEta.Location = New Point(card.Width - lblEta.PreferredWidth - 10, 10)
        card.Controls.Add(lblEta)
    End If
    
    ' Click handler — propagate from card and all children
    AddHandler card.Click, clickHandler
    AddHandler lblName.Click, Sub(s, ev) clickHandler(card, ev)
    
    ' Hover effects
    AddHandler card.MouseEnter, Sub(s, ev)
        card.BackColor = Color.FromArgb(232, 247, 255) ' #E8F7FF
    End Sub
    AddHandler card.MouseLeave, Sub(s, ev)
        card.BackColor = Color.White ' reset to theme
    End Sub
    
    parent.Controls.Add(card)
End Sub
```

#### 5d. Selected State
When a card is clicked and running, set its background to accent color:
```vb
' On run start:
card.BackColor = Color.FromArgb(0, 180, 255)  ' accent
lblName.ForeColor = Color.White

' On run end:
card.BackColor = Color.White  ' reset
lblName.ForeColor = Color.FromArgb(26, 26, 46)
```

#### 5e. Hover Effects
- Light theme: BG shifts from `#FFFFFF` → `#E8F7FF` on hover, border shifts from `#DEE2E6` → `#00B4FF`
- Dark theme: BG shifts from `#162C45` → `#1A3555`, border shifts from `#1E3A5F` → `#00B4FF`

#### 5f. Tier-Locked Styling
Locked cards get:
- Muted background: `#F1F3F5` (light) / `#0D2236` (dark)
- 🔒 icon prepended to title
- Tooltip: "Upgrade to unlock"
- No hover effect
- Reduced opacity text

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` `AddButton()` | Replace with `AddCardButton()` — new panel-based cards |
| `app/MainForm.vb` `BuildButtons()` | Use new card method, adjust Y spacing to 60px per card |
| `app/MainForm.vb` `StyleDisabledButton()` | Update for panel-based cards |
| `app/MainForm.vb` `ApplyTierLocks()` | Update control type check from `Button` to `Panel` |
| `app/MainForm.vb` `RunDesignTool()` | Update sender extraction (Tag from Panel) |

> **Important:** The `Tag` property carries the slot integer. All click handlers must extract `Tag` from the Panel, not a Button. This is the main risk area — test thoroughly.

---

## 6. Engineering Tools Panel

**Risk: MEDIUM** — Same card treatment as §5.

### Current State
- 4 hardcoded buttons (slots 11–14) in `pnlEngineeringContent`
- Same `AddButton()` styling as design tools
- Panel is 300px wide (wider than design panel's 210px)

### Proposed Changes

#### 6a. Same Card Treatment
Use `AddCardButton()` from §5 with engineering-specific subtitles:

| Slot | Title | Subtitle |
|------|-------|----------|
| 11 | Conveyor Calculator | Belt & chain conveyor sizing |
| 12 | Pneumatic Calculator | Cylinder force & bore sizing |
| 13 | Air Consumption | Compressor & consumption sizing |
| 14 | Beam Deflection | Frame stress & deflection check |

#### 6b. Tool Icons
SVG icons exist in `output/assets/icons/`. WinForms can't render SVGs directly. Options:
1. **Pre-convert SVGs to 24×24 PNGs** and load via `PictureBox` in each card (recommended)
2. Ship both SVG + PNG in `output/assets/icons/`

Add a 24×24 `PictureBox` at `(8, 14)` in each card panel, left of the title label. Shift title label to `(40, 8)`.

#### 6c. Section Header
Style the "ENGINEERING TOOLS" label at top of panel:
```vb
Dim lblHeader As New Label() With {
    .Text = "ENGINEERING TOOLS",
    .Dock = DockStyle.Top,
    .Height = 36,
    .Font = New Font("Segoe UI", 9.0F, FontStyle.Bold),
    .TextAlign = ContentAlignment.MiddleCenter,
    .ForeColor = Color.FromArgb(0, 180, 255),  ' accent color
    .BackColor = Color.Transparent
}
```

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` `BuildButtons()` | Use `AddCardButton()` for engineering tools too |
| `output/assets/icons/` | Add 24×24 PNG versions of tool icons |

---

## 7. Log Window

**Risk: LOW** — Cosmetic + minor functional (scroll-lock). No logic changes.

### Current State
- Plain `TextBox` with `Multiline = True`, `ReadOnly = True`
- Default system colors, default font
- `Log()` and `LogA()` append timestamped text with emoji prefixes

### Proposed Changes

#### 7a. Replace TextBox with RichTextBox
A `RichTextBox` enables per-line coloring for log levels.

```vb
' Replace:
Private txtLog As TextBox
' With:
Private rtbLog As RichTextBox
```

#### 7b. Console-Style Appearance (Both Themes)
The log should always have a dark background — like an IDE console:

```vb
rtbLog = New RichTextBox() With {
    .Multiline = True,
    .ReadOnly = True,
    .ScrollBars = RichTextBoxScrollBars.Vertical,
    .BackColor = Color.FromArgb(18, 18, 24),   ' #121218 — near-black
    .ForeColor = Color.FromArgb(200, 210, 220), ' light gray text
    .Font = New Font("Consolas", 9.0F),
    .BorderStyle = BorderStyle.None,
    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
    .Location = New Point(250, 150),
    .Size = New Size(Me.ClientSize.Width - 620, Me.ClientSize.Height - 220),
    .DetectUrls = False
}
```

#### 7c. Colored Log Levels
Update `LogA()` to append colored text:

```vb
Private Sub LogA(icon As String, msg As String)
    If rtbLog Is Nothing Then Return
    
    Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss") & "  "
    Dim fullMsg As String = icon & " " & msg & vbCrLf
    
    rtbLog.SelectionStart = rtbLog.TextLength
    rtbLog.SelectionLength = 0
    
    ' Timestamp in muted color
    rtbLog.SelectionColor = Color.FromArgb(100, 110, 120)
    rtbLog.AppendText(timestamp)
    
    ' Message in level color
    Dim levelColor As Color = Color.FromArgb(200, 210, 220) ' default
    If icon = "✅" Then levelColor = Color.FromArgb(0, 200, 83)      ' green
    If icon = "⚠️" Then levelColor = Color.FromArgb(255, 179, 0)     ' amber
    If icon = "❌" Then levelColor = Color.FromArgb(244, 67, 54)      ' red
    If icon = "⛔" Then levelColor = Color.FromArgb(244, 67, 54)      ' red
    If icon = "📁" Then levelColor = Color.FromArgb(0, 180, 255)      ' teal/info
    If icon = "🔗" Then levelColor = Color.FromArgb(169, 199, 232)    ' muted blue
    If icon = "☁️" Then levelColor = Color.FromArgb(0, 180, 255)      ' info
    
    rtbLog.SelectionColor = levelColor
    rtbLog.AppendText(fullMsg)
    
    ' Auto-scroll
    If Not scrollLocked Then
        rtbLog.SelectionStart = rtbLog.TextLength
        rtbLog.ScrollToCaret()
    End If
End Sub
```

#### 7d. Scroll-Lock Toggle
Add a small checkbox or toggle button above the log:

```vb
Private scrollLocked As Boolean = False

Private chkScrollLock As CheckBox
' In BuildCentre():
chkScrollLock = New CheckBox() With {
    .Text = "🔒 Scroll Lock",
    .AutoSize = True,
    .Font = New Font("Segoe UI", 7.5F),
    .Location = New Point(rtbLog.Right - 100, rtbLog.Top - 18),
    .Anchor = AnchorStyles.Top Or AnchorStyles.Right
}
AddHandler chkScrollLock.CheckedChanged, Sub(s, ev) scrollLocked = chkScrollLock.Checked
Me.Controls.Add(chkScrollLock)
```

#### 7e. Log Panel Border
Wrap `rtbLog` in a Panel with a 1px border for clean separation:
```vb
Dim pnlLog As New Panel() With {
    .BackColor = Color.FromArgb(30, 40, 55),  ' slightly lighter border color
    .Padding = New Padding(1),
    .Location = rtbLog.Location,
    .Size = New Size(rtbLog.Width + 2, rtbLog.Height + 2),
    .Anchor = rtbLog.Anchor
}
rtbLog.Dock = DockStyle.Fill
pnlLog.Controls.Add(rtbLog)
Me.Controls.Add(pnlLog)
```

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` | Replace `txtLog` with `rtbLog`, update `Log()`, `LogA()`, add scroll-lock |
| `app/tools/ThemeApplier.vb` | Ensure RichTextBox log keeps dark BG even in light theme (skip theming for controls named `rtbLog`) |

---

## 8. Footer

**Risk: LOW** — Cosmetic restructuring of footer label.

### Current State
- Single `lblFooter` label with registered company name
- Positioned at bottom-left, anchored to bottom
- No status information shown

### Proposed: Status Bar Footer

#### 8a. Footer Panel
Replace the single label with a proper status bar panel:

```vb
Private pnlFooter As Panel
Private lblFooterSW As Label       ' SolidWorks status
Private lblFooterTier As Label     ' License tier
Private lblFooterSeat As Label     ' Seat status
Private lblFooterTheme As Label    ' Theme indicator
Private lblFooterReg As Label      ' Registered name

Private Sub BuildFooter()
    pnlFooter = New Panel() With {
        .Dock = DockStyle.Bottom,
        .Height = 28,
        .BackColor = Color.FromArgb(241, 243, 245),  ' themed
        .Padding = New Padding(8, 0, 8, 0)
    }
    
    ' Separator line at top
    Dim sep As New Panel() With {
        .Dock = DockStyle.Top,
        .Height = 1,
        .BackColor = Color.FromArgb(222, 226, 230)
    }
    pnlFooter.Controls.Add(sep)
    
    ' SolidWorks connection
    lblFooterSW = New Label() With {
        .Text = "● SW: Not Connected",
        .Font = New Font("Segoe UI", 8.0F),
        .ForeColor = Color.FromArgb(108, 117, 125),
        .AutoSize = True,
        .Location = New Point(10, 6)
    }
    pnlFooter.Controls.Add(lblFooterSW)
    
    ' License tier
    lblFooterTier = New Label() With {
        .Text = "TRIAL",
        .Font = New Font("Segoe UI", 8.0F, FontStyle.Bold),
        .ForeColor = Color.FromArgb(0, 180, 255),
        .AutoSize = True,
        .Location = New Point(200, 6)
    }
    pnlFooter.Controls.Add(lblFooterTier)
    
    ' Seat status
    lblFooterSeat = New Label() With {
        .Text = "S:1/3",
        .Font = New Font("Segoe UI", 8.0F),
        .AutoSize = True,
        .Location = New Point(320, 6)
    }
    pnlFooter.Controls.Add(lblFooterSeat)
    
    ' Theme indicator (right-aligned)
    lblFooterTheme = New Label() With {
        .Text = "☀ Light",
        .Font = New Font("Segoe UI", 8.0F),
        .AutoSize = True,
        .Anchor = AnchorStyles.Right Or AnchorStyles.Top,
        .Location = New Point(pnlFooter.Width - 80, 6)
    }
    pnlFooter.Controls.Add(lblFooterTheme)
    
    Me.Controls.Add(pnlFooter)
End Sub
```

#### 8b. SolidWorks Status Updates
After `EnsureSolidWorksRunning()` succeeds:
```vb
lblFooterSW.Text = "● SW: Connected (" & year & ")"
lblFooterSW.ForeColor = Color.FromArgb(0, 200, 83)  ' green
```

On failure:
```vb
lblFooterSW.Text = "● SW: Not Connected"
lblFooterSW.ForeColor = Color.FromArgb(244, 67, 54)  ' red
```

#### 8c. Theme Indicator
Updated in `ApplyTheme()`:
```vb
lblFooterTheme.Text = If(currentThemeIsDark, "🌙 Dark", "☀ Light")
```

### Files to Change
| File | Changes |
|------|---------|
| `app/MainForm.vb` `BuildFooter()` | Full rewrite to status bar panel |
| `app/MainForm.vb` | Update footer labels on SW connect, theme change, license load |

---

## 9. Calculator Forms Standardization

**Risk: MEDIUM** — Restructuring UI layout of multiple forms. Same calculation logic.

### Current State
- Each calculator form (Conveyor, Pneumatic, Air Consumption) builds its own UI independently
- Inconsistent spacing, colors, fonts
- Each has its own `ApplyMDATTheme()` hook
- PDF export buttons exist but placement varies

### Proposed: Base Visual Standards

#### 9a. Common Layout Pattern
All calculator forms should follow this structure:

```
┌─────────────────────────────────────────────────────┐
│  HEADER: Tool Name                         [× Close]│
├──────────────────┬──────────────────────────────────┤
│  INPUT PANEL     │  RESULTS PANEL                   │
│  (Left, ~360px)  │  (Right, fill)                   │
│                  │                                   │
│  [Group: Load]   │  ┌─ Summary ──────────────────┐  │
│   Label: Input   │  │  Final result text          │  │
│   Label: Input   │  │  (monospace, bordered)      │  │
│                  │  └─────────────────────────────┘  │
│  [Group: Cyl]    │                                   │
│   Label: Input   │  ┌─ Calculations ─────────────┐  │
│   Label: Input   │  │  Step-by-step              │  │
│                  │  │  (scrollable)              │  │
│  ┌─────────────┐ │  └─────────────────────────────┘  │
│  │ CALCULATE   │ │                                   │
│  │ RESET  PDF  │ │  ┌─ Warnings ─────────────────┐  │
│  └─────────────┘ │  │  (colored RichTextBox)     │  │
│                  │  └─────────────────────────────┘  │
├──────────────────┴──────────────────────────────────┤
│  STATUS BAR: Ready | Last calc: 0.2s                 │
└─────────────────────────────────────────────────────┘
```

#### 9b. Standard Colors for Calculator Forms
| Element | Light Theme | Dark Theme |
|---------|------------|------------|
| Form BG | `#F8F9FA` | `#0B1E34` |
| Input Panel BG | `#FFFFFF` | `#0F2A44` |
| Results Panel BG | `#F8F9FA` | `#0B1E34` |
| GroupBox Border | `#DEE2E6` | `#1E3A5F` |
| Input Field BG | `#FFFFFF` | `#162C45` |
| Input Field Border | `#CED4DA` | `#1E3A5F` |
| Calculate Button | `#00B4FF` bg, white text | Same |
| Reset Button | `#6C757D` bg, white text | Same |
| PDF Button | `#00C853` bg, white text | Same |
| Result Summary BG | `#F1F8E9` (light green tint) | `#0D2818` (dark green tint) |
| Warning BG | `#FFF8E1` (light amber tint) | `#2D1F00` (dark amber tint) |

#### 9c. Standard Button Bar
All calculator forms should have a consistent bottom button bar:

```vb
' Standard button bar (inside input panel, at bottom)
Dim pnlButtons As New FlowLayoutPanel() With {
    .Dock = DockStyle.Bottom,
    .Height = 45,
    .FlowDirection = FlowDirection.LeftToRight,
    .Padding = New Padding(5)
}

btnCalc = MakeStyledButton("CALCULATE", Color.FromArgb(0, 180, 255), Color.White)
btnCalc.Size = New Size(120, 34)

btnReset = MakeStyledButton("RESET", Color.FromArgb(108, 117, 125), Color.White)
btnReset.Size = New Size(80, 34)

btnPdf = MakeStyledButton("📄 PDF", Color.FromArgb(0, 200, 83), Color.White)
btnPdf.Size = New Size(80, 34)
```

#### 9d. Report-Style Results
The results `TextBox` should look like a professional engineering report:
- Monospace font: `Consolas 9pt`
- Bordered panel wrapper
- Clear section separators using `═══════` or similar
- Units always shown
- Pass/fail indicators with colored text (via RichTextBox)

#### 9e. New Helper: `CalculatorFormBase` Considerations
Creating a base class would be ideal but is HIGH RISK (inheritance changes). Instead, create a **helper module**:

**New file: `app/tools/CalcFormStyler.vb`**
```vb
Public Module CalcFormStyler
    Public Function MakeStyledButton(text As String, bg As Color, fg As Color) As Button
        Dim b As New Button() With {
            .Text = text,
            .BackColor = bg,
            .ForeColor = fg,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.5F, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        b.FlatAppearance.BorderSize = 0
        Return b
    End Function
    
    Public Sub StyleGroupBox(gb As GroupBox, isDark As Boolean)
        gb.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        gb.ForeColor = If(isDark, Color.FromArgb(169, 199, 232), Color.FromArgb(0, 180, 255))
    End Sub
    
    Public Sub StyleInputField(tb As TextBox, isDark As Boolean)
        tb.Font = New Font("Segoe UI", 9.5F)
        tb.BorderStyle = BorderStyle.FixedSingle
        tb.BackColor = If(isDark, Color.FromArgb(22, 44, 69), Color.White)
        tb.ForeColor = If(isDark, Color.FromArgb(234, 244, 255), Color.FromArgb(26, 26, 46))
    End Sub
End Module
```

### Files to Change/Create
| File | Action | Risk |
|------|--------|------|
| `app/tools/CalcFormStyler.vb` | **NEW** — shared styling helpers | LOW |
| `app/tools/ConveyorCalculatorForm.vb` | Apply standard colors/layout | MEDIUM |
| `app/tools/PneumaticCylinderCalculatorForm.vb` | Apply standard colors/layout | MEDIUM |
| `app/tools/AirConsumptionForm.vb` | Apply standard colors/layout | MEDIUM |
| All other calculator forms | Apply standard colors/layout | MEDIUM |

---

## 10. Splash Screen

**Risk: LOW** — New form, no interaction with existing logic.

### Design

```
┌─────────────────────────────────────────┐
│                                         │
│           [MetaMech Logo]               │
│                                         │
│     MDAT — Mechanical Design            │
│        Automation Tool                  │
│                                         │
│          v2.1.0                         │
│                                         │
│    ████████████████░░░░░  78%           │
│    Loading license...                   │
│                                         │
│    © 2024 MetaMech Solutions            │
└─────────────────────────────────────────┘
```

### Implementation

**New file: `app/SplashForm.vb`**

```vb
Public Class SplashForm
    Inherits Form
    
    Private lblAppName As Label
    Private lblVersion As Label
    Private lblStatus As Label
    Private prgBar As ProgressBar
    Private picLogo As PictureBox
    Private lblCopyright As Label
    
    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(480, 320)
        Me.BackColor = Color.FromArgb(11, 30, 52)  ' Navy
        Me.ShowInTaskbar = False
        
        ' Logo
        picLogo = New PictureBox() With {
            .Size = New Size(200, 80),
            .Location = New Point(140, 30),
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BackColor = Color.Transparent
        }
        Dim logoPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets\logo\logo.png")
        Try
            If IO.File.Exists(logoPath) Then
                picLogo.Image = New Bitmap(logoPath)
            End If
        Catch
        End Try
        Me.Controls.Add(picLogo)
        
        ' App Name
        lblAppName = New Label() With {
            .Text = "MDAT",
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = Color.FromArgb(234, 244, 255),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Size = New Size(480, 45),
            .Location = New Point(0, 120),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblAppName)
        
        ' Subtitle
        Dim lblSub As New Label() With {
            .Text = "Mechanical Design Automation Tool",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.FromArgb(169, 199, 232),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Size = New Size(480, 22),
            .Location = New Point(0, 165),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblSub)
        
        ' Version
        lblVersion = New Label() With {
            .Text = "v" & GetVersionSafe(),
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(100, 140, 180),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Size = New Size(480, 18),
            .Location = New Point(0, 190),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblVersion)
        
        ' Progress bar
        prgBar = New ProgressBar() With {
            .Style = ProgressBarStyle.Continuous,
            .Size = New Size(360, 6),
            .Location = New Point(60, 230),
            .ForeColor = Color.FromArgb(0, 180, 255)
        }
        Me.Controls.Add(prgBar)
        
        ' Status
        lblStatus = New Label() With {
            .Text = "Initializing...",
            .Font = New Font("Segoe UI", 8.5F),
            .ForeColor = Color.FromArgb(100, 140, 180),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Size = New Size(480, 18),
            .Location = New Point(0, 245),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblStatus)
        
        ' Copyright
        lblCopyright = New Label() With {
            .Text = "© 2024 MetaMech Solutions",
            .Font = New Font("Segoe UI", 7.5F),
            .ForeColor = Color.FromArgb(74, 106, 138),
            .TextAlign = ContentAlignment.MiddleCenter,
            .Size = New Size(480, 16),
            .Location = New Point(0, 290),
            .BackColor = Color.Transparent
        }
        Me.Controls.Add(lblCopyright)
    End Sub
    
    Public Sub UpdateProgress(pct As Integer, status As String)
        prgBar.Value = Math.Min(pct, 100)
        lblStatus.Text = status
        Application.DoEvents()
    End Sub
    
    Private Function GetVersionSafe() As String
        Try
            Dim v As Version = Reflection.Assembly.GetExecutingAssembly().GetName().Version
            If v IsNot Nothing Then Return v.ToString()
        Catch
        End Try
        Return "1.0"
    End Function
End Class
```

### Usage in Program.vb / MainForm
```vb
' In Program.vb Main():
Dim splash As New SplashForm()
splash.Show()
splash.UpdateProgress(10, "Loading configuration...")

Dim main As New MainForm()

splash.UpdateProgress(80, "Validating license...")
' ... license init happens in MainForm constructor

splash.UpdateProgress(100, "Ready")
Threading.Thread.Sleep(300)
splash.Close()
splash.Dispose()

Application.Run(main)
```

### Files to Create/Change
| File | Action | Risk |
|------|--------|------|
| `app/SplashForm.vb` | **NEW** | LOW |
| `app/Program.vb` (or MainForm startup) | Add splash display logic | LOW |
| `Build_EXE.bat` | Add `SplashForm.vb` to compilation | LOW |

---

## 11. Risk Assessment Summary

| Section | Risk | Rationale |
|---------|------|-----------|
| §1 Premium Light Theme | **LOW** | Color constant changes only |
| §2 Premium Dark Theme | **LOW** | Color constant changes only |
| §3 Theme System | **MEDIUM** | New persistence file, restructured `ApplyTheme()`, ComboBox replaces ContextMenu |
| §4 Header Overhaul | **LOW** | Visual-only changes within locked layout |
| §5 Macro Button Panel | **MEDIUM** | Button → Panel refactor, click handler routing changes |
| §6 Engineering Tools Panel | **MEDIUM** | Same as §5 |
| §7 Log Window | **LOW** | TextBox → RichTextBox swap, same append pattern |
| §8 Footer | **LOW** | New status bar panel, no logic coupling |
| §9 Calculator Standardization | **MEDIUM** | Per-form UI changes across 5+ forms |
| §10 Splash Screen | **LOW** | New standalone form, no existing code touched |

### HIGH RISK items (none proposed, but noted):
- Changing `TierLocks.vb` logic → NOT in scope
- Changing `Licensing.vb` / seat enforcement → NOT in scope
- Creating a `CalculatorFormBase` class hierarchy → avoided in favor of helper module
- Modifying macro execution flow → NOT in scope

---

## 12. Implementation Order

Recommended phased approach:

### Phase 1: Foundation (LOW RISK)
1. Update color constants in MainForm (`BG_LIGHT`, `PANEL_LIGHT`, `BG_DARK`, `PANEL_DARK`)
2. Update `ThemeApplier.vb` dark-mode colors to navy
3. Update `.theme` files
4. Create `UISettings.vb` for theme persistence
5. Create `SplashForm.vb`

### Phase 2: MainForm Visual Upgrade (LOW-MEDIUM)
1. Header visual enhancements (§4)
2. Log window upgrade — TextBox → RichTextBox (§7)
3. Footer status bar (§8)
4. Theme switcher ComboBox (§3b)
5. Full `ApplyTheme()` rewrite using ThemeApplier on self (§3c)

### Phase 3: Card-Style Panels (MEDIUM)
1. Implement `AddCardButton()` for design tools (§5)
2. Apply to engineering tools (§6)
3. Update `ApplyTierLocks()` for panel-based cards
4. Update click handlers for Panel tag extraction
5. **Thorough testing** of all button clicks + tier locks

### Phase 4: Calculator Forms (MEDIUM)
1. Create `CalcFormStyler.vb`
2. Standardize ConveyorCalculatorForm
3. Standardize PneumaticCylinderCalculatorForm
4. Standardize AirConsumptionForm
5. Standardize remaining calculator forms

---

## New Files Summary

| File | Purpose |
|------|---------|
| `app/tools/UISettings.vb` | Theme persistence (read/write `ui.settings`) |
| `app/tools/CalcFormStyler.vb` | Shared calculator form styling helpers |
| `app/SplashForm.vb` | Branded splash screen |
| `output/assets/icons/*.png` | 24×24 PNG tool icons (converted from existing SVGs) |

## Modified Files Summary

| File | Scope of Changes |
|------|-----------------|
| `app/MainForm.vb` | Theme colors, `ApplyTheme()`, `AddButton()` → `AddCardButton()`, log RichTextBox, footer, header styling |
| `app/UITheme.vb` | Update default colors to premium light palette |
| `app/tools/ThemeApplier.vb` | Navy dark-mode colors, skip log panel theming |
| `output/assets/themes/light.theme` | Full rewrite with premium palette |
| `output/assets/themes/dark.theme` | Update to navy palette |
| `app/tools/ConveyorCalculatorForm.vb` | Standardized styling |
| `app/tools/PneumaticCylinderCalculatorForm.vb` | Standardized styling |
| `app/tools/AirConsumptionForm.vb` | Standardized styling |
| `Build_EXE.bat` | Add new `.vb` files to compilation |
