**English** | [Русский](README.md)

---

# 📦 Snippets Library for Macro Tools VBA Add-in

> Ready-to-use VBA code snippets to accelerate Excel development.

This repository provides a curated collection of code snippets designed for use with the [Macro Tools VBA](https://github.com/vbatools/Addin_MacroToolsVBA) add-in. Snippets sync automatically from this repository, giving you instant access to reliable, production-tested code templates.

---

## 📖 Purpose

The snippets library extends **Macro Tools VBA** with:
- ✅ Pre-built templates for common VBA tasks
- ✅ Automatic updates via repository sync
- ✅ Logical categorization by functionality
- ✅ IntelliSense-friendly naming and auto-completion support

---

## 🛠️ Setup & Import Workflow

To manage snippets in your environment:

1. Open the **Macro Tools VBA** toolbar
2. Navigate to **Other tools**
3. Choose one of the following actions:

| Action | Description |
|--------|-------------|
| 📂 **Export Snippets to Folder** | Exports all add-in snippets to a local `ADDIN_MACRO_TOOLS_SNIPPETS` folder, preserving category structure |
| 📄 **Add local JSON from Folder** | Compiles exported snippet folders into a single JSON configuration file |
| 📄 **Import Snippets from local JSON** | Imports snippets from a locally generated JSON file |
| 🌐 **Import Snippets from GitHub** | Fetches the latest snippet library directly from this repository *(warning: overwrites existing snippets)* |

---

## 💻 Using Snippets in Your Code

1. Open the VBA Editor (`ALT + F11`) and navigate to your target module
2. Start typing a snippet keyword (e.g., `pusub`, `arr`, `LastRow`)
3. Select **Insert Code** from the context menu
4. ✅ The code inserts automatically; dependent modules or forms load as needed

### ⚡ Quick Access
- `ALT + Z` — Trigger snippet insertion *(requires [MacroToolsVBAHotKeys](https://github.com/vbatools/MacroToolsVBAHotKeys) module)*

> 💡 **Pro Tip:** Forgot a snippet name? Use the **Insert Module** tool in the VBE toolbar to generate a helper module named `SNIPPETS`. It contains a full catalog of available templates with descriptions and enables IntelliSense (e.g., `SNIPPETS.App.prsub`).  
> You can safely delete the `SNIPPETS` module anytime via the **Delete Module** button.

---

## 📁 Library Structure

Snippets are organized into functional categories for easy navigation:

| Category | Description | Count |
|----------|-------------|-------|
| [App](#app-12) | Core application procedures and utilities | 12 |
| [Custom](#custom-36) | Custom helpers and domain-specific functions | 36 |
| [DoLoop](#doloop-8) | Loop templates: `For`, `Do While`, `Do Until` | 8 |
| [FSO](#fso-9) | File system operations via `FileSystemObject` | 9 |
| [Forms](#forms-16) | UserForm controls and dynamic UI components | 16 |
| [Http](#http-3) | HTTP requests and web content handling | 3 |
| [Iff](#iff-3) | Conditional logic and branching templates | 3 |
| [LogFile](#logfile-3) | Logging and debugging utilities | 3 |
| [TxtFile](#txtfile-9) | Text file I/O, encoding, and CSV export | 9 |
| **Total** | | **99** |

---

### App (12)

| Snippet | Description |
|---------|-------------|
| `CalAu` | Enable automatic calculation: `Calculation = xlCalculationAutomatic` |
| `CalMa` | Disable calculation & screen updating for macro performance |
| `Enum_` | Template for `Enum` declaration |
| `pusub` | `Public Sub` template |
| `prsub` | `Private Sub` template |
| `pufun` | `Public Function` template |
| `prfun` | `Private Function` template |
| `pucon` | `Public Const` template |
| `ScrF` | Quick toggle: disable `ScreenUpdating` and `Calculation` |
| `ScrT` | Restore `ScreenUpdating` and `Calculation` post-execution |
| `ThWb` | `With ThisWorkbook.Worksheets(SHName)` shortcut |
| `Wth` | Generic `With` block template |

### Custom (36)

| Snippet | Description |
|---------|-------------|
| `addCheckBox` | Create a `CheckBox` in a range with macro binding |
| `arr` | Load Excel range into VBA array |
| `arrDemens` | Detect and validate array dimensionality |
| `arrFor` | `For` loop template for 1D array iteration |
| `arrFor2` | Loop template for parallel processing of two arrays |
| `arrSort` | Sort a 1D array (ascending/descending) |
| `arrSort2DArray` | Sort 2D array by specified column index |
| `arrTo2D` | Convert 1D array to 2D structure |
| `arrTranspose` | Transpose array (rows ↔ columns) |
| `arrUniqueValuesFrom` | Extract unique values from an array |
| `BlendColor` | Blend two RGB colors programmatically |
| `Clear` | Clear range values while preserving formatting |
| `delAllFilesInPath` | Bulk-delete all files in a directory |
| `dialogFile` / `FileDialog` | Open file picker dialog (single/multi-select) |
| `extractAllFilesFromZip` | Extract all contents from a ZIP archive |
| `fileIsBusy` | Check if a file is locked by another process |
| `FileHave` | Verify file existence at a given path |
| `getPath` | Resolve and normalize folder paths |
| `GetListFiles` | Return array of filenames in a directory |
| `getLetterByNomer` | Convert column number to letter (1 → "A") |
| `GetUniqueValueFromRange` | Extract unique values from an Excel range |
| `hash` | Generate string hash (MD5/SHA variants) |
| `HaveSheetInFile` | Check if a worksheet exists in an external workbook |
| `LastCol` / `LastRow` | Find last used column/row in a worksheet |
| `Path` | Validate and sanitize file/folder paths |
| `replaceSymbols` | Replace characters in string using pattern/rules |
| `Resize` | Write array to range with auto-resizing |
| `ScreenUpdatingCalculation` | Toggle performance settings for macro execution |
| `SheetCopy` / `SheetDelete` | Safely copy or delete worksheets |
| `Status` / `StatusBar` | Display messages in Excel's status bar |
| `wbIsOpen` | Check if a workbook is already open in the session |
| `WriterLogSub` | Universal logging subroutine for macro tracing |

### DoLoop (8)

| Snippet | Description |
|---------|-------------|
| `DoLoUn` | `Do ... Loop Until` template |
| `DoLoWh` | `Do ... Loop While` template |
| `DoUnLo` | `Do Until ... Loop` template |
| `DoWhLo` | `Do While ... Loop` template |
| `ForEa` | `For Each ... In` template |
| `ForEaW` | Loop through all worksheets in active workbook |
| `ForI` | Basic `For I = 1 To N` loop |
| `ForIS` | `For` loop with explicit step: `Step 1` |

### FSO (9)

| Snippet | Description |
|---------|-------------|
| `delPath` | Delete folder (recursive if non-empty) |
| `MoveFile` | Move or rename a file |
| `sFileExists` | Check file existence via FSO |
| `sFolderExists` | Check folder existence via FSO |
| `sGetBaseName` | Get filename without extension |
| `sGetExtensionName` | Get file extension |
| `sGetFileName` | Get full filename with extension |
| `sGetParentFolderName` | Get parent directory path |
| `sGetTempName` | Generate unique temporary filename |

### Forms (16)

| Snippet | Description |
|---------|-------------|
| `clsAnchors` | Class for responsive ("anchored") control positioning |
| `clsButtonIcon` | Button class with icon support and styling |
| `clsCalendarDate` / `clsCalendarTime` | Date/time picker for `TextBox` controls |
| `clsComboBox` / `clsComboBoxForm` | Enhanced `ComboBox` with multi-select & modal dialog |
| `clsContextMenu` | Dynamic context menu for form controls |
| `cls ListBox` | Advanced `ListBox` handling class |
| `clsModernStyle` | Apply modern UI theme to UserForms |
| `clsProgresBar` | Animated progress bar control |
| `clsSlider` | Configurable slider with event hooks |
| `clsTextArea` | Scrollable multi-line `TextBox` |
| `clsTextboxMask` | `TextBox` with input masking (dates, phones, etc.) |
| `fint` | Center UserForm on screen |
| `ftxtbtn` / `ftxtbtnclick` | `DropButton` for `TextBox` + click handler template |

### Http (3)

| Snippet | Description |
|---------|-------------|
| `responseTextHttp` | Perform GET/POST request and retrieve response text |
| `webPageHTML` | Fetch raw HTML of a web page |
| `webPageText` | Extract visible text content from a web page |

### Iff (3)

| Snippet | Description |
|---------|-------------|
| `IfEl` | `If ... Else ... End If` template |
| `IfEn` | Minimal `If ... End If` block |
| `SeCa` | `Select Case ... End Select` template |

### LogFile (3)

| Snippet | Description |
|---------|-------------|
| `ShowLog` | Display log file contents in a message box or debug window |
| `WriteErrorLog` | Log errors with timestamp and call stack |
| `WriteLog` | General-purpose log entry writer |

### TxtFile (9)

| Snippet | Description |
|---------|-------------|
| `Base64` | Encode/decode strings using Base64 |
| `base64ToFile` / `fileToBase64` | Convert between file and Base64 string |
| `loadTextFromTextFile` | Read text file with encoding support |
| `Range2CSV` | Export Excel range to CSV format |
| `saveTextToFile` | Save string content to a text file |
| `TXTAddIntoTXTFile` | Append text to existing file |
| `TXTReadALLFile` | Read entire file content into a single string |
| `TXTWriteFile` | Overwrite file with new content |

---

## 👤 Author

**VBATools** 

## 📄 License

Distributed under the [Apache License 2.0](LICENSE).  
You are free to use, modify, and distribute this code in compliance with the license terms.