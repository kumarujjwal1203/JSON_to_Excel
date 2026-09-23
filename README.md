# JSON_to_Excel — GST JSON to Excel Converter (Octa GST Edition)

A professional, standalone Windows desktop application built with **.NET 8 (LTS)**, **C#**, **WPF**, and **ClosedXML** that converts GST-related JSON & ZIP files (`GSTR-1`, `GSTR-2A`, `GSTR-2B`, `GSTR-3B`) into Excel (`.xlsx`) format matching the **Exact Octa GST Layout** with **100% Data Accuracy** and **Zero Data Loss**.

---

## 📥 Direct Download (`.exe` & Setup Installer)

| Package | Direct Download Link | Size | Description |
| :--- | :--- | :--- | :--- |
| **1-Click Setup Installer (Recommended)** | [**Download `GSTJsonToExcel_Setup.exe`**](publish/installer/GSTJsonToExcel_Setup.exe) | ~89 MB | Universal Windows Installer (64-bit & 32-bit) with Desktop & Start Menu Shortcuts |
| **Portable 64-Bit `.exe`** | [**Download `GSTJsonToExcel.exe` (x64)**](publish/win-x64/GSTJsonToExcel.exe) | ~71 MB | Direct standalone 64-bit executable (No installation required) |
| **Portable 32-Bit `.exe`** | [**Download `GSTJsonToExcel.exe` (x86)**](publish/win-x86/GSTJsonToExcel.exe) | ~66 MB | Direct standalone 32-bit executable (No installation required) |

---

## 🌟 Key Features

### 1. Official Application Branding & Logo
* **Embedded Multi-Resolution Windows Icon**: The application now features the official circular branding logo (**GST JSON TO EXCEL — SMART • FAST • ACCURATE**):
  * Embedded directly into `GSTJsonToExcel.exe` (both 64-bit and 32-bit builds).
  * Embedded into the universal installer `GSTJsonToExcel_Setup.exe` (shows the logo on download, in Explorer, and when shared).
  * Automatically applies to Windows Desktop shortcuts, Start Menu shortcuts, Taskbar, and window title bar.
  * Rendered as a circular badge inside the application's header.

### 2. 1-to-1 Individual File Conversion (Default Mode)
* **Separate Excel per JSON File**: When multiple JSON files are selected (e.g. 10 `R2B` JSON files), the application generates **10 individual Excel files** named directly after the source files (e.g. `R2B_001.xlsx`, `R2B_002.xlsx`), each formatted in the exact Octa GST layout for that return type.
* **No Unwanted Merging**: Data from different JSON files is kept strictly in separate workbooks by default.
* **Optional Consolidation Toggle**: A "Consolidate by GST Return Type" mode is also available for users who want to combine files into single `R1.xlsx`, `R2B.xlsx`, etc.

### 3. Interactive & Modern Windows UI
* **Interactive Scanned Files DataGrid**:
  * Checkbox column (`✓`) to selectively include/exclude files from conversion.
  * Color-coded type badges (`R1`, `R2A`, `R2B`, `R3A`, `Duplicate`, `Ignored`).
  * Target output Excel filename column (e.g. `R2B_001.xlsx`).
  * Real-time file processing status pills (`Pending` ⏳, `Converting...` ⚙️, `Completed` ✅, `Failed` ❌).
* **Search & Quick Category Filters**:
  * Category filter chips: **All Files**, **R1**, **R2A**, **R2B**, **R3A**.
  * Quick **Select All** / **Deselect All** action links.
  * Real-time search box (`🔍`) to filter files by filename instantly.
* **Direct File Actions**:
  * **"Open Excel"** button to open the converted file immediately.
  * **"Show in Folder"** button to highlight the generated file in Windows Explorer.
  * **"Open Output Folder"** button for the entire batch.

### 4. Exact Octa GST Layout & Styling
* Strictly adheres to the **Octa GST** workbook structures:
  * **`R1`**: Sheets: `Overview`, `Sales` (30 cols), `Sales Summary` (15 cols), `SalesHSN` (15 cols), `Disclosed` (8 cols), `All_Data_Index`.
  * **`R2A`**: Sheets: `Overview`, `Purchase` (33 cols), `ISD` (18 cols), `TDS` (9 cols), `TCS` (11 cols), `All_Data_Index`.
  * **`R2B`**: Sheets: `Overview`, `Purchase` (39 cols), `ISD` (19 cols), `All_Data_Index`.
* **Overview Sheet Styling**:
  * Top banner: Merged royal blue header (`#0070C0`) **"MAP & Associates"** in bold white text.
  * Metadata card: Column A in royal blue with bold white labels (`Company Name`, `Contents`, `Company GSTIN`, `Period`, `Creation Time`), Column B with extracted data.
  * GSTIN state lookup: Decodes 2-digit Indian GST state codes into state names (e.g. `07ACWFS8659K2ZV (Delhi)`, `06ACWFS8659K1ZY (Haryana)`).
  * Bottom banner: Right-aligned **"Created by Octa GST"** in royal blue.

### 5. Multi-Source Input & Drag-and-Drop
* **Select Files**: Pick single or multiple `.json` files via the file picker dialog.
* **Select Folder**: Select an entire folder containing hundreds of files.
* **Drag & Drop**: Drag multiple files or entire directory trees straight into the application.

### 6. Smart Filtering (Zero File Deletion)
* The application automatically identifies valid candidate GST JSON files and **safely ignores non-JSON files** (`.pdf`, `.xml`, `.png`, `.txt`, `.docx`, etc.).
* **All original user files and folders remain 100% untouched on disk!**
* Ignored non-GST files are tracked and reported in the classification dashboard.

### 7. Automatic GST Type Classification (2-Stage Detection)
* Automatically categorizes input files into:
  * **`R1`** (GSTR-1 Outward Supplies)
  * **`R3A`** (GSTR-3B / 3A Returns)
  * **`R2A`** (GSTR-2A Inward Auto-Drafted)
  * **`R2B`** (GSTR-2B Static Auto-Drafted ITC)
  * **`Unknown / Review Required`** (Ambiguous files with detailed explanation)
* **Priority 1**: Filename pattern matching (`R1_*.json`, `*GSTR1*.json`, `*R3A*.json`, `*R2A*.json`, `*R2B*.json`).
* **Priority 2**: Deep JSON inspection of root and nested properties (e.g. `sup_details` / `taxpayble` for R3A; `b2b` + `b2cs`/`hsn` for R1; `itc_summ` for R2B; `cfs` supplier filing status for R2A).

### 5. SHA-256 Duplicate Content Detection
* Calculates SHA-256 content hashes of the JSON files.
* Detects exact duplicate content even if files have different names (e.g. `R1_September.json` and `R1_September_Copy.json`).
* Skips duplicates during export so tax totals are never double-counted, while clearly recording them in the conversion audit report.

### 6. Separate Excel File for Each GST Type
* Output files are organized in a dedicated directory:
  ```text
  GST_Converted/
  ├── R1.xlsx        (all R1 files formatted in exact Octa GST format)
  ├── R3A.xlsx       (all R3A files aggregated)
  ├── R2A.xlsx       (all R2A files formatted in exact Octa GST format)
  └── R2B.xlsx       (all R2B files formatted in exact Octa GST format)
  ```
* **No data mixing between GST return types.**
* Types with 0 files are **not** created.

### 7. Zero Data Loss Guarantee (110% Accuracy)
* Sensitive identifiers (GSTIN, Invoice Numbers, HSN/SAC, Reference numbers) and numbers with leading zeros (e.g. `00001234`) are strictly written as Excel Text (`@` format) so leading zeros are never dropped.
* High-precision decimals (e.g. `58917516.44`, `123456.7891`) use `decimal` parsing to avoid floating-point loss.
* Safety net `All_Data_Index` sheet in every workbook records every JSON leaf element by JSONPath.

### 8. Universal 32-bit & 64-bit Windows Compatibility
* Standalone single-file executables for both **64-bit** (`win-x64`) and **32-bit** (`win-x86`).
* Universal Inno Setup installer (`GSTJsonToExcel_Setup.exe`) that detects the host OS at install time.
* **100% Offline & Private**: No data leaves the machine; requires no pre-installed .NET runtime or Microsoft Excel.

---

## 📁 Repository Structure

```text
GSTIN/
│
├── GSTJsonToExcel.sln                  # Visual Studio Solution (.NET 8)
│
├── src/
│   └── GSTJsonToExcel/                 # WPF MVVM Application
│       ├── Models/
│       │   ├── GstFileType.cs          # R1, R3A, R2A, R2B, Unknown
│       │   ├── ScannedFileItem.cs      # Scanned item with type, hash, status
│       │   ├── BatchScanSummary.cs     # Grouped classification counts
│       │   ├── BatchConversionResult.cs# Generated Excel files & error reports
│       │   ├── JsonFlatTable.cs        # Relational table definition
│       │   └── ...
│       ├── Services/
│       │   ├── Interfaces/
│       │   │   ├── IFileScannerService.cs
│       │   │   ├── IGstClassifierService.cs
│       │   │   ├── IDuplicateDetectorService.cs
│       │   │   ├── IBatchConversionService.cs
│       │   │   ├── IJsonParserService.cs
│       │   │   └── IExcelExportService.cs
│       │   └── Implementations/
│       │       ├── FileScannerService.cs
│       │       ├── GstClassifierService.cs
│       │       ├── DuplicateDetectorService.cs
│       │       ├── OctaGstBuilderService.cs
│       │       ├── BatchConversionService.cs
│       │       ├── JsonParserService.cs
│       │       └── ExcelExportService.cs
│       ├── ViewModels/
│       │   └── MainViewModel.cs        # Orchestrates scanning, classification, and batch conversion
│       └── Views/
│           └── MainWindow.xaml         # Dashboard with badges, dual buttons, progress, & file list
│
├── tests/
│   └── GSTJsonToExcel.Tests/           # Automated Test Suite (30 Tests, 100% Pass)
│       ├── OctaGstFormatTests.cs
│       ├── BatchAndClassificationTests.cs
│       ├── SampleGstr3bConversionTests.cs
│       ├── LeadingZerosAndPrecisionTests.cs
│       ├── EdgeCasesAndIntegrityTests.cs
│       └── HierarchyAndValidationTests.cs
│
├── publish/
│   ├── installer/
│   │   └── GSTJsonToExcel_Setup.exe    # 1-Click Universal Windows Installer (93 MB)
│   ├── win-x64/
│   │   └── GSTJsonToExcel.exe          # Standalone 64-bit portable app (171 MB)
│   └── win-x86/
│       └── GSTJsonToExcel.exe          # Standalone 32-bit portable app (161 MB)
│
└── test_data/
    └── batch_sample/                   # Ready-to-test mixed folder
        ├── R1_001.json
        ├── R1_002.json
        ├── R1_001_Copy.json            # Duplicate file
        ├── R3A_001.json
        ├── R2A_001.json
        ├── R2B_001.json
        ├── sample.pdf                  # Ignored non-JSON file
        ├── document.xml                # Ignored non-JSON file
        ├── notes.txt                   # Ignored non-JSON file
        └── image.png                   # Ignored non-JSON file
```

---

## 🚀 How to Run the Application

### Option 1: Universal 1-Click Installer (Recommended)
1. Double-click:
   ```text
   publish\installer\GSTJsonToExcel_Setup.exe
   ```
2. Click **Next → Install → Finish**.
3. It detects **32-bit vs 64-bit** Windows automatically and creates Start Menu and Desktop shortcuts.

### Option 2: Portable Standalone Executables
* **Any Windows PC (32-bit or 64-bit)**: Run `publish\win-x86\GSTJsonToExcel.exe`.
* **64-bit Windows PCs**: Run `publish\win-x64\GSTJsonToExcel.exe`.

---

## 🧪 Automated Test Suite (32 Tests)

Run the automated tests using:
```powershell
dotnet test GSTJsonToExcel.sln
```

All 32 tests pass with 100% success:
* `ConvertBatch_IndividualMode_GeneratesSeparateExcelPerJsonFile_NeverLumpsIntoOne` (10 JSON files generate 10 individual Excel files, no merging)
* `ConvertBatch_IndividualMode_UncheckedItemsAreSkipped` (selective file conversion)
* `OctaGstBuilder_R1_GeneratesExactSheetsAndColumns` (Overview banner, Sales 30 cols, Sales Summary 15 cols, SalesHSN 15 cols, Disclosed 8 cols)
* `OctaGstBuilder_R2A_GeneratesExactSheetsAndColumns` (Overview, Purchase 33 cols, ISD 18 cols, TDS 9 cols, TCS 11 cols)
* `OctaGstBuilder_R2B_GeneratesExactSheetsAndColumns` (Overview, Purchase 39 cols, ISD 19 cols)
* `GstStateHelper_FormatsGstinCorrectly` (Indian GST 2-digit state code mapping)
* `BatchConversionService_WithOctaBuilder_GeneratesOctaForR1_AndStandardForR3A`
* `ClassifyByFileName_AccuratelyIdentifiesTypes` (all GST return types)
* `ClassifyFileAsync_FallsBackToContentInspection_ForAmbiguousNames`
* `DuplicateDetector_IdentifiesIdenticalContent_ByHash`
* `FileScanner_IgnoresNonJsonFiles_LeavesUserFilesUntouched`
* `BatchConversion_GeneratesSeparateExcelPerGstType_NeverMixesData`
* `Convert_PromptSampleGstr3b_ZeroDataLoss_IntegrityPassed`
* `Convert_PreservesLeadingZeros_NeverConvertsToInteger`
* `Convert_EdgeCases_PreservesNulls_EmptyArrays_AndFutureFields`
* `Convert_Gstr1_ParentKeysInheritedByChildRecords`
* `ValidateJsonFile_HandlesMalformedAndMissingFilesGracefully`
