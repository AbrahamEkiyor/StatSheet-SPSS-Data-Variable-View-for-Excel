# StatSheet — SPSS Data & Variable View for Excel

**StatSheet** brings the **Data View**, **11-column Variable View**, and **reversible Value Labels (`1 ↔ A`) toggle** of SPSS directly to Microsoft Excel and the browser—complete with missing-value definitions, validation rules, portable JSON/CSV codebooks, descriptive statistics, and a **true zero-server offline Excel VBA (`.xlam`) add-in**.

---

## Key Capabilities

- **SPSS Data View**
  - High-density spreadsheet grid with sticky case numbers, variable headers (with `Scale` / `Ordinal` / `Nominal` measure icons), and an active-cell inspector displaying both raw numeric code and resolved label.
  - Inline editing with automatic searchable dropdowns when Value Labels are active, plus undo/redo history (`40` steps), row insertion/deletion, live search, and validation-issue filtering.
- **Full 11-Column SPSS Variable View (+ Class Bins)**
  - Inline editing for **Name**, **Type** (`Numeric`, `Comma`, `Dot`, `Scientific`, `Date`, `Dollar`, `Custom Currency`, `String`), **Width**, **Decimals**, **Label**, **Values**, **Missing**, **Class Bins**, **Columns**, **Align**, **Measure** (`Scale`, `Ordinal`, `Nominal`), and **Role** (`Input`, `Target`, `Both`, `None`, `Partition`, `Split`).
  - Built-in SPSS variable-naming validation, atomic column renaming across all cases, row reordering, and variable duplication.
- **Reversible Value Labels Toggle (`1 ↔ A`)**
  - Switch spreadsheets seamlessly between raw codes (`1`, `2`, `-9`) and human-readable text labels (`Strongly Disagree`, `Disagree`, `Skipped Question`).
  - Preserves original raw codes in hidden `_SPSS_RawBackup` and `_SPSS_Metadata` worksheets so toggling Value Labels off restores numeric values without corrupting formulas or statistical coding.
- **Value Labels Editor**
  - Add, update, or remove `code = label` pairs, apply quick presets (`5-pt Likert`, `Yes/No`, `Frequency 1–5`), copy label maps from another variable, or bulk-paste `code = label` lines.
- **Missing Values & Data Validation**
  - Configure SPSS discrete missing codes (up to 3 values) or numeric missing ranges (`low thru high`).
  - Enforce numeric `Min`/`Max` bounds, required-field checks, and **Restrict to defined Value Labels** (which also generates native Excel in-cell dropdown validation on sync).
- **Scale Class Boundaries & Reusable Binning Templates**
  - Group continuous numeric/scale variables via **Equal Count**, **Fixed Width** (with optional start bound & nice rounded cutpoints), or **Custom Intervals** (with open-ended `−∞` / `+∞` bounds, custom labels, and live overlap/gap validation).
  - Configurable endpoint inclusion rule: **Lower-inclusive `[a, b)`** (default) or **Upper-inclusive `(a, b]`**.
  - Check **Remember these binning settings** to persist the interval configuration in browser `localStorage` and apply the saved template to other variables in one click.
  - **Recode into New Variable** generates a coded **Ordinal** column (`1, 2, …`) with matching Value Labels and preserved user-missing codes.
- **Descriptives & Grouped Frequency Distributions**
  - Slide-over inspector calculating **Total N**, **Valid N**, **User-Missing N**, **System-Missing N**, **Mean**, **Median**, **Std. Dev.**, **Min**, **Max**, and both **Ungrouped** and **Grouped Class Interval** frequency tables (with **Midpoint**, `Frequency`, `Percent`, `Valid %`, `Cumulative %`) excluding missing codes.
- **Codebooks & Workbook Portability**
  - Import and export `.xlsx` and `.csv` workbooks via SheetJS (preserving hidden `_SPSS_RawBackup` and `_SPSS_Metadata` sheets, including `ClassBinsJSON`, on `.xlsx` round-trips).
  - Export and import portable **JSON Codebooks** (`statsheet-codebook-v1`) and **CSV Data Dictionaries** (including class boundary definitions).
- **Built-in Sample Datasets & Task Pane Preview**
  - Includes an *Employee Engagement Survey (N=24)* and *Cardiology Clinical Trial (N=12)* dataset out of the box, plus a **Task Pane Preview** (`420px`) toggle to test narrow Excel side-pane ergonomics.

---

## Three Ways to Run StatSheet

| Mode | Best For | Internet / Server Required? |
| :--- | :--- | :--- |
| **1. Offline VBA Add-in (`.xlam`)** | Air-gapped Desktop Excel (Windows & macOS) | **No** — 100% offline, zero server |
| **2. Office.js Task Pane (`manifest.xml`)** | Rich React UI inside Excel Desktop or Excel on the Web | Requires HTTPS host URL (`localhost` or deployed URL) |
| **3. Browser Simulator** | Evaluating, cleaning `.xlsx`/`.csv` files, and building codebooks in any browser | Runs locally via Vite or hosted static build |

---

## Local Development Setup

### Prerequisites
- **Node.js** `18+` and **npm**

### Install & Run

```bash
# 1. Install dependencies
npm install

# 2. Start the development server (runs on http://localhost:3000)
npm run dev
```

### Build & Typecheck

```bash
# TypeScript check
npm run lint

# Production build (outputs to dist/)
npm run build

# Preview production build locally
npm run preview
```

---

## Installing in Desktop Excel (Offline `.xlam` VBA Add-in)

The **Offline VBA Add-in** runs purely in cross-platform Excel VBA (`StatSheet_SPSS.bas`) with no ActiveX/Windows-only dependencies, so it works on both **Windows** and **macOS** Desktop Excel with zero network access.

1. Open StatSheet in the browser and click **Excel Add-in** in the top navigation bar.
2. Under the **Offline VBA Add-in (.xlam)** tab, download **`StatSheet_SPSS.bas`** (and optionally **`Build_StatSheet_xlam.ps1`** for Windows or **`customUI14.xml`** for a custom Ribbon tab).

### Option A — Automated Windows Installer (`.ps1`)
1. Save `StatSheet_SPSS.bas` and `Build_StatSheet_xlam.ps1` in the same folder.
2. Open PowerShell in that folder and run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Build_StatSheet_xlam.ps1
   ```
   *(Requires `File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model` to be checked once).*
3. The script builds `%APPDATA%\Microsoft\AddIns\StatSheet_SPSS.xlam`, embeds the custom **`SPSS View`** Ribbon tab (`customUI14.xml`), and enables the add-in automatically.

### Option B — Manual Import (Windows & macOS, 30 Seconds)
1. Open a blank workbook in Excel and open the VBA Editor (`Alt + F11` on Windows, `Option + F11` on macOS).
2. Choose **File → Import File…** and select `StatSheet_SPSS.bas`.
3. Return to Excel, choose **File → Save As**, and set the format to **Excel Add-in (`*.xlam`)**.
4. Enable the add-in:
   - **Windows**: **File → Options → Add-ins → Manage: Excel Add-ins → Go…** → check `StatSheet_SPSS`.
   - **macOS**: **Tools → Excel Add-ins…** → check `StatSheet_SPSS`.

### Offline VBA Ribbon, Shortcuts & Menu Commands
Once enabled, StatSheet adds a native **`SPSS View`** Ribbon tab (via Option A), a **StatSheet SPSS** submenu when right-clicking any cell, an **Add-ins → StatSheet SPSS** menu, and two global shortcuts:
- `Ctrl + Shift + L` — **Toggle Value Labels (`1 ↔ A`)** reversibly (includes an option to suppress repeat confirmation pop-ups while keeping status-bar updates).
- `Ctrl + Shift + V` — **Switch between Data View and `Variable_View`**.
- **Class Boundaries & Grouped Stats**: **Define Class Boundaries (Scale)** (`AUTO:k`, `WIDTH:w@start`, or custom intervals with `[a, b)` / `(a, b]` rule stored in `Variable_View` column `O`), **Descriptives & Frequencies Report** (with automatic grouped interval + Midpoint + Cumulative % output), and **Recode Binned Variable** (creates a coded Ordinal column with Value Labels).
- Additional menu/Ribbon actions: **Sync / Initialize `Variable_View`**, **Edit Value Labels for Selection**, **Validate Rules & Flag Missing Values**, **Clear Validation Highlights**, **Export Codebook (`.json`)**, and **Toggle Pop-up Confirmations (On/Off)**.

---

## Installing in Excel via Office.js (`manifest.xml`)

For the full React Task Pane inside Excel Desktop or Excel on the Web:

1. Open the **Excel Add-in** dialog → **Office.js Task Pane (`manifest.xml`)** tab.
2. Enter a reachable `https://` host URL (either your publicly deployed HTTPS URL or a local HTTPS server such as `https://localhost:3000` with `office-addin-dev-certs`), then click **Download `manifest.xml`**.
3. Sideload `statsheet-excel-manifest.xml`:
   - **Windows Desktop**: Place the XML in a shared network folder (`\\YOUR-PC\StatSheet`), register it under **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, check **Show in Menu**, restart Excel, and open **Home → Add-ins → More Add-ins → Shared Folder**.
   - **macOS Desktop**: Copy the XML to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`, restart Excel, and open **Home → Add-ins → My Add-ins**.
   - **Excel on the Web**: Open **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in**.

---

## Workbook Storage & Hidden Sheets Architecture

To prevent Value Labels from destroying numeric formulas or underlying codes, StatSheet maintains dedicated hidden worksheets inside `.xlsx` workbooks and live Excel sessions:

| Sheet Name | Visibility | Purpose |
| :--- | :--- | :--- |
| `Data` *(or active sheet)* | Visible | Primary case-by-variable spreadsheet view (raw codes or toggled Value Labels) |
| `Variable_View` *(VBA mode)* | Visible | 11-column SPSS variable definitions plus `Val_Min`, `Val_Max`, `Restrict_Labels`, and `Class_Bins` |
| `_SPSS_RawBackup` | Hidden (`xlSheetVeryHidden`) | Snapshot of exact raw numeric/text codes before applying Value Labels |
| `_SPSS_Metadata` | Hidden (`xlSheetVeryHidden`) | Serialized Variable View schema, value labels, missing rules, validation bounds, and `ClassBinsJSON` |
| `_SPSS_State` *(VBA mode)* | Hidden (`xlSheetVeryHidden`) | Tracks active `RAW`/`LABELS` display mode and target data sheet name |
| `_SPSS_Output` *(VBA mode)* | Visible on run | Generated Descriptive Statistics & Ungrouped / Grouped Frequency Distribution report |

---

## Project Structure

```text
├── index.html                        # Entry HTML, Google Fonts, and metadata
├── metadata.json                     # App metadata
├── package.json                      # Scripts and dependencies
├── public/                           # 16px, 32px, 80px PNG icons for Office manifest
└── src/
    ├── App.tsx                       # Main workspace shell, undo/redo, host & modal orchestration
    ├── main.tsx                      # React root with error boundary
    ├── index.css                     # Tailwind CSS v4 theme tokens and grid styles
    ├── types/spss.ts                 # TypeScript models for Variable View, Data View, stats, codebook
    ├── data/sampleDatasets.ts        # Built-in Employee Survey & Clinical Trial datasets
    ├── bridge/workbookBridge.ts      # Dual Office.js Excel.run + Simulator / SheetJS (.xlsx/.csv) bridge
    ├── engine/
    │   ├── spssEngine.ts             # Formatting, missing/validation checks, stats, codebook, manifest.xml
    │   └── vbaAddinGenerator.ts      # Pure cross-platform VBA (.bas), PowerShell .xlam builder, RibbonX XML
    └── components/
        ├── TopBar.tsx                # 3-zone header navigation and Sync / Export CTA
        ├── RibbonToolbar.tsx         # Value Labels (1↔A) toggle, undo/redo, presets, import, filter
        ├── DataViewGrid.tsx          # SPSS Data View grid with formula bar and label dropdown editors
        ├── VariableViewGrid.tsx      # 11-column SPSS Variable View editor
        ├── BottomViewTabs.tsx        # Classic Data View / Variable View sheet switcher
        ├── ValueLabelsModal.tsx      # Code-to-label editor, presets, copy, and bulk paste
        ├── MissingAndValidationModal.tsx # Missing value rules and numeric/label validation bounds
        ├── ClassBoundariesModal.tsx  # Equal-width & custom class intervals, templates, and recoding
        ├── StatsInspectorDrawer.tsx  # Slide-over Descriptives & Ungrouped/Grouped Frequency inspector
        ├── CodebookModal.tsx         # JSON Codebook & CSV Data Dictionary import/export
        ├── ExcelAddinModal.tsx       # Offline VBA (.xlam) & Office.js manifest.xml installer hub
        └── MeasureIcon.tsx           # SPSS Scale, Ordinal, and Nominal SVG icons
```

---

## Tech Stack

- **Frontend**: React 19, TypeScript, Vite, Tailwind CSS v4, Lucide React
- **Spreadsheet Engine**: SheetJS (`xlsx`) for `.xlsx`/`.csv` import & export, Office.js (`@types/office-js`) for live Excel integration
- **Offline Desktop Add-in**: Cross-platform Excel VBA (`StatSheet_SPSS.bas`), Windows PowerShell COM builder, RibbonX XML
