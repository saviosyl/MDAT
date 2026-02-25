# MetaMech MDAT — User Manual

**Mechanical Design Automation Tool**  
Version 2026 | MetaMech Solutions | Ireland  
[metamechsolutions.com](https://metamechsolutions.com)

---

## Table of Contents

1. [Getting Started](#1-getting-started)
2. [Installation](#2-installation)
3. [Licensing](#3-licensing)
4. [Interface Overview](#4-interface-overview)
5. [Design Tools (SolidWorks Automation)](#5-design-tools-solidworks-automation)
6. [Engineering Tools](#6-engineering-tools)
7. [PDF Merge Tool](#7-pdf-merge-tool)
8. [Settings & Themes](#8-settings--themes)
9. [Troubleshooting](#9-troubleshooting)
10. [Support & Contact](#10-support--contact)

---

## 1. Getting Started

MetaMech MDAT is a professional SolidWorks automation suite that eliminates repetitive, error-prone tasks mechanical engineers deal with every day. Built for engineers who value their time.

### System Requirements

- **OS:** Windows 10 or 11 (64-bit)
- **.NET Framework:** 4.0 or later (pre-installed on most Windows PCs)
- **SolidWorks:** 2018 or later (required for Design Tools)
- **Internet:** Required for first-time activation only

### What's Included

| File | Purpose |
|------|---------|
| `MDAT.exe` | Main application |
| `PdfMergeTool.exe` | PDF merge engine |
| `PdfSharp-gdi.dll` | PDF processing library |
| `Config.txt` | Server & tool configuration |
| `MetaMech_RSA_PUBLIC.xml` | License verification key |
| `assets\` | Icons, themes, logos |

---

## 2. Installation

### Step 1: Download

Download `MetaMech-Trial.zip` from [metamechsolutions.com/download](https://metamechsolutions.com/download)

### Step 2: Extract

1. Right-click the ZIP file → **Extract All**
2. Choose a location (e.g., `C:\MetaMech\` or your Desktop)
3. Click **Extract**

> **Important:** Do NOT extract into `Program Files` — Windows permissions can cause issues. Use your Desktop or a folder like `C:\MetaMech\`.

### Step 3: Launch

Double-click `MDAT.exe` to start MetaMech.

On first launch, the app will automatically contact the MetaMech server to activate your **free 3-day trial**. You'll see a confirmation message:

> *"Welcome to MetaMech! Your 3-day free trial has been activated."*

No license key needed — it's automatic.

### Step 4: Connect SolidWorks

1. Open SolidWorks
2. In MetaMech, **select your SolidWorks version** from the dropdown (e.g., 2022)
3. Click **"Select Assembly (.SLDASM)"** to choose your assembly file
4. Now you can use the Design Tools

> **Important:** You must select the SolidWorks version and assembly BEFORE clicking any Design Tool button.

---

## 3. Licensing

### Free Trial

- **Duration:** 3 days from first launch
- **Features:** All tools unlocked
- **Limits:** One trial per machine — cannot be reset by reinstalling
- **No key required:** Activates automatically on first launch

### Purchasing a License

When your trial expires, visit [metamechsolutions.com/pricing](https://metamechsolutions.com/pricing) to purchase a license.

**License Tiers:**

| Tier | Name | Access |
|------|------|--------|
| 1 | Standard | Core design automation tools |
| 2 | Premium | All tools including engineering calculators |
| 3 | Premium Plus | Everything + priority support |

### Activating a Paid License

1. After purchase, you'll receive a `license.key` file
2. Place it in the same folder as `MDAT.exe`
3. Launch MetaMech — your license will be detected automatically

### License Information Display

The main window shows your license status:
- **Tier** (Trial / Standard / Premium / Premium Plus)
- **Expiry date**
- **Days remaining**
- **Seats** (how many PCs can use this license simultaneously)

### Seat Management

Your license allows a set number of simultaneous users (seats). If all seats are in use:
- A message will appear: *"No seats available"*
- Close MetaMech on another PC, or wait — inactive seats are automatically released after 6 hours

### Offline Use

After initial activation, MetaMech works offline for a grace period:
- Trial: 1 day offline
- Standard: 3 days offline
- Premium: 7 days offline
- Premium Plus: 14 days offline

After the grace period, connect to the internet briefly to re-validate.

---

## 4. Interface Overview

MetaMech has a clean, modern interface with two main sections:

### Main Window

```
┌─────────────────────────────────────────┐
│  MetaMech MDAT                    [─][□][×]
├─────────────────────────────────────────┤
│  LICENCE: ACTIVE | TIER: PREMIUM | SEATS: 3
│  ✅ 340 days remaining
├──────────────┬──────────────────────────┤
│              │                          │
│  DESIGN      │     Activity Log         │
│  TOOLS       │                          │
│              │  🧩 ACTION START — Slot 1│
│  ┌────────┐  │  📄 Script: SMART BOM   │
│  │SMART   │  │  ✅ Completed (4.2s)    │
│  │BOM     │  │                          │
│  └────────┘  │                          │
│  ┌────────┐  │                          │
│  │SMART   │  │                          │
│  │PDF     │  │                          │
│  └────────┘  │                          │
│  ...         │                          │
│              │                          │
│  ENGINEERING │                          │
│  TOOLS       │                          │
│  ┌────────┐  │                          │
│  │Conveyor│  │                          │
│  │Calc    │  │                          │
│  └────────┘  │                          │
│  ...         │                          │
├──────────────┴──────────────────────────┤
│  Assembly: C:\Projects\Machine.SLDASM   │
└─────────────────────────────────────────┘
```

### Left Panel — Tool Buttons
Click any tool button to run it. Tools are grouped into:
- **Design Tools** — SolidWorks automation macros
- **Engineering Tools** — Standalone calculators and utilities

### Right Panel — Activity Log
Shows real-time progress, timing, and results for each action.

### Status Bar
Shows the currently loaded SolidWorks assembly.

---

## 5. Design Tools (SolidWorks Automation)

These tools automate SolidWorks tasks.

**Before using any Design Tool:**
1. Open SolidWorks on your PC
2. In MetaMech, select your **SolidWorks version** from the dropdown
3. Click **"Select Assembly (.SLDASM)"** to load your assembly
4. Then click the tool button

> **Important:** Always select version and assembly first. Clicking a tool button without an assembly loaded will fail.

### SMART BOM
**Generate Bill of Materials from any assembly.**

1. Open your assembly in SolidWorks
2. Click **SMART BOM** in MetaMech
3. BOM is extracted and exported to Excel automatically

**Output:** Excel file with part numbers, descriptions, quantities — flat or indented.

### SMART PDF
**Batch-export all drawings to PDF.**

1. Open your assembly in SolidWorks
2. Click **SMART PDF**
3. All associated drawings are found and exported to PDF
4. PDFs are saved to the output folder

**Output:** Individual PDF files for each drawing.

### SMART STEP
**Batch-export assembly parts to STEP format.**

1. Open your assembly in SolidWorks
2. Click **SMART STEP**
3. All parts and sub-assemblies are exported to STEP files

**Output:** `.stp` files for each component — ready for sharing with vendors or other CAD systems.

### SMART DXF
**Batch-export flat patterns and profiles to DXF.**

1. Open your assembly in SolidWorks
2. Click **SMART DXF**
3. Sheet metal flat patterns and profiles are exported to DXF

**Output:** `.dxf` files ready for laser cutting, CNC, or waterjet.

### SMART TEMPLATE CHANGE
**Push standardised templates and properties across files.**

1. Open your assembly in SolidWorks
2. Click **SMART TEMPLATE CHANGE**
3. Custom properties and drawing templates are updated across all files

**Use case:** Enforce company standards, update title blocks, sync properties.

### MDAT ONE-CLICK
**Run the complete automation pipeline in one click.**

1. Open your assembly in SolidWorks
2. Click **MDAT ONE-CLICK**
3. Runs BOM + PDF + STEP exports in sequence automatically

**Use case:** Full release package — BOM, PDFs, and STEP files generated in one go.

---

## 6. Engineering Tools

Standalone calculators that work without SolidWorks.

### Conveyor Calculator
Calculate belt conveyors — speed, capacity, motor power, belt tension.

- Select conveyor type (flat, inclined, decline)
- Enter: belt width, speed, material density, length, incline angle
- Get: throughput, required motor power, belt tension

### Motor Power Calculator
Calculate required motor power for mechanical systems.

- Enter: torque, speed (RPM), efficiency
- Get: power in kW/HP, current draw estimates

### Torque Calculator
Calculate torque for various mechanical applications.

- Enter: force, radius/lever arm
- Get: torque in Nm, lb-ft, and other units

### Pneumatic Cylinder Calculator
Size pneumatic cylinders for your application.

- Enter: required force, pressure, stroke length
- Get: bore size, rod diameter, air consumption

### Air Consumption Calculator
Calculate compressed air usage for pneumatic systems.

- Enter: cylinder sizes, cycle rates, pressure
- Get: total air consumption (l/min, CFM)

### Unit Converter
Convert between engineering units:
- Length (mm, in, ft, m)
- Force (N, lbf, kgf)
- Pressure (bar, PSI, MPa, kPa)
- Torque (Nm, lb-ft, kgf-cm)
- And more

### Engineering Notepad
A simple notepad for engineering notes, calculations, and project documentation.

---

## 7. PDF Merge Tool

Merge multiple PDF files into a single document with an auto-generated index page.

### How to Use

1. Click the **PDF Merge** tool in MetaMech
2. **Add files:** Click "Add PDF Files" or "Add Folder" to select PDFs
3. **Reorder:** Use Move Up / Move Down buttons to set the order
4. **Merge:** Click one of:
   - **"Merge with Index"** — creates an index page with clickable links + page numbers
   - **"Merge (No Index)"** — simple merge without index
5. Choose where to save the output
6. Done — merged PDF opens automatically

### Index Page Features
- Hierarchical numbering (1., 1.1., 1.2., etc.)
- Drawing name or custom description
- Start page number for each document
- Clickable links — click any entry to jump to that drawing
- Auto page numbering on every page ("Page 1 of 47")

### Command-Line Usage (Advanced)
PdfMergeTool can also be run from command line:

```
PdfMergeTool.exe -out "C:\output\Combined.pdf" -list "C:\output\merge_order.txt"
PdfMergeTool.exe -noindex -out "C:\output\Combined.pdf" -list "C:\output\merge_order.txt"
```

The `merge_order.txt` file contains one PDF path per line. Use tabs for indentation (creates hierarchy) and a second tab-separated column for custom display names.

---

## 8. Settings & Themes

### Dark / Light Theme
MetaMech supports dark and light themes. Toggle via the theme button in the main window.

Your preference is saved and remembered between sessions.

### Configuration (Config.txt)
The `Config.txt` file contains server and tool settings. You generally don't need to edit this unless instructed by support.

```
SEAT_SERVER=https://metamech-license-server.saviosyl.workers.dev
CLIENT_TOKEN=<your token>
TELEMETRY=OFF
```

> **Note:** Do not share your `Config.txt` file — it contains your connection credentials.

---

## 9. Troubleshooting

### "License not valid" / "Activate your licence"
- **First launch?** Make sure you're connected to the internet. The trial activates automatically.
- **Trial expired?** Purchase a license at [metamechsolutions.com/pricing](https://metamechsolutions.com/pricing)
- **Have a license?** Place your `license.key` file in the same folder as `MDAT.exe`

### "No seats available"
All license seats are in use on other machines. Either:
- Close MetaMech on another PC
- Wait 6 hours — inactive seats release automatically
- Contact support if the issue persists

### "SolidWorks not available"
- Make sure SolidWorks is running before clicking Design Tools
- Open an assembly in SolidWorks first
- Supported: SolidWorks 2018–2026

### Design Tool runs but nothing happens
- Check that an assembly is open (not a part or drawing)
- Check the Activity Log for error messages
- Make sure the assembly is fully resolved (no lightweight/suppressed components causing issues)

### "Trial already used on this machine"
The free trial is limited to one per machine. To continue using MetaMech, purchase a license.

### PDF Merge fails
- Ensure `PdfMergeTool.exe` and `PdfSharp-gdi.dll` are in the same folder as `MDAT.exe`
- Check that all input PDFs are valid and not password-protected

### App won't start / crashes immediately
- Make sure .NET Framework 4.0+ is installed
- Try running as Administrator (right-click → Run as Administrator)
- Check that all files from the ZIP are extracted (not running from inside the ZIP)

---

## 10. Support & Contact

**Website:** [metamechsolutions.com](https://metamechsolutions.com)  
**Contact:** [metamechsolutions.com/contact](https://metamechsolutions.com/contact)  
**Pricing:** [metamechsolutions.com/pricing](https://metamechsolutions.com/pricing)

---

*© 2026 MetaMech Solutions, Ireland. All rights reserved.*
