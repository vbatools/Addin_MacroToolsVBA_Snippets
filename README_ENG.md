**English** | [Русский](README.md)
---

# 📦 Snippet Database for Macro Tools VBA Add-in

This repository contains a collection of snippets for the [Macro Tools VBA](https://github.com/vbatools/Addin_MacroToolsVBA) add-in.

## 📖 Description

The snippet database is designed for use with the Macro Tools VBA add-in. Snippets are automatically loaded from this repository into the add-in, providing quick access to frequently used VBA code fragments.

## 📁 Repository Structure

The repository is organized by functional categories:

| Category | Description | Snippet Count |
|----------|-------------|---------------|
| [App](#app-10) | Main application procedures and functions | 10 |
| [Custom](#custom-46) | Custom procedures and functions | 46 |
| [DoLoop](#doloop-8) | Loops (For, Do While, Do Until) | 8 |
| [FSO](#fso-9) | File system operations | 9 |
| [Forms](#forms-16) | Form controls | 16 |
| [Http](#http-3) | HTTP requests | 3 |
| [Iff](#iff-3) | Conditional statements | 3 |
| [LogFile](#logfile-3) | Log file operations | 3 |
| [TxtFile](#txtfile-9) | Text file operations | 9 |
| **Total** | | **110** |

### App (10)

| Snippet | Description |
|---------|-------------|
| `CalAu` | Calculation = xlCalculationAutomatic |
| `CalMa` | Calculation = xlCalculationManual ScreenUpdating = False |
| `Enum_` | Enum |
| `pusub` | Public Sub |
| `prsub` | Private Sub |
| `pufun` | Public Function |
| `prfun` | Private Function |
| `pucon` | Public Const |
| `ScrF` | ScreenUpdating = False Calculation = xlCalculationManual |
| `ScrT` | Calculation = xlCalculationAutomatic ScreenUpdating = True |
| `ThWb` | With ThisWorkbook.Worksheets(SHName) |
| `Wth` | With |
| `pusub` | Public Sub |
| `pufun` | Public Function |
| `pucon` | Public Const |
| `prsub` | Private Sub |
| `prfun` | Private Function |

### Custom (46)

| Snippet | Description |
|---------|-------------|
| `addCheckBox` | Create CheckBox in range with macro binding |
| `arr` | Load range into array |
| `arrDemens` | Check array dimension |
| `arrFor` | Loop through array |
| `arrFor2` | Loop through 2 arrays |
| `arrSort` | Sort array |
| `arrSort2DArray` | Sort 2D array |
| `arrTo2D` | Convert 1D array to 2D |
| `arrTranspose` | Transpose array |
| `arrUniqueValuesFrom` | Unique values from array |
| `BlendColor` | Mix two colors |
| `Clear` | Clear range of values |
| `delAllFilesInPath` | Delete all files in folder |
| `dialogFile` | File selection |
| `FileDialog` | File dialog |
| `extractAllFilesFromZip` | Extract files from archive |
| `fileIsBusy` | Check if file is busy |
| `FileHave` | Check file existence |
| `getPath` | Get folder path |
| `GetListFiles` | List files in folder |
| `getLetterByNomer` | Column letter by number |
| `getSheetNameCloseBook` | Sheet names from closed book |
| `GetSheetNameCloseBook` | Get sheet names from closed book |
| `GetUniqueValueFromRange` | Unique values from range |
| `hash` | Hash functions |
| `HaveSheetInFile` | Check sheet in file |
| `LastCol` | Last column |
| `LastRow` | Last row |
| `Path` | Path check |
| `replaceSymbols` | Replace symbols in string |
| `Resize` | Output array to range |
| `ScreenUpdatingCalculation` | Toggle ScreenUpdating/Calculation |
| `SheetCopy` | Copy sheet |
| `SheetDelete` | Delete sheet |
| `Status` | Status bar message |
| `StatusBar` | Status bar |
| `Status` | Status bar |
| `wbIsOpen` | Check if workbook is open |
| `WriterLogSub` | Logger |

### DoLoop (8)

| Snippet | Description |
|---------|-------------|
| `DoLoUn` | Do Loop Until |
| `DoLoWh` | Do Loop While |
| `DoUnLo` | Do Until Loop |
| `DoWhLo` | Do While Loop |
| `ForEa` | For Each ... In |
| `ForEaW` | For Each sheet in book |
| `ForI` | For I = 1 To 10 |
| `ForIS` | For I = 1 To 10 Step 1 |

### FSO (9)

| Snippet | Description |
|---------|-------------|
| `delPath` | Delete folder |
| `MoveFile` | Move file |
| `sFileExists` | Check file exists |
| `sFolderExists` | Check folder exists |
| `sGetBaseName` | File name without extension |
| `sGetExtensionName` | File extension |
| `sGetFileName` | File name |
| `sGetParentFolderName` | Parent folder |
| `sGetTempName` | Random file name |

### Forms (16)

| Snippet | Description |
|---------|-------------|
| `clsAnchors` | Responsive form |
| `clsButtonIcon` | Button with icon |
| `clsCalendarDate` | Calendar for TextBox |
| `clsCalendarTime` | Time for TextBox |
| `clsComboBox` | ComboBox Multi class |
| `clsComboBoxForm` | ComboBox with modal window |
| `clsContextMenu` | Context menu |
| `cls ListBox` | ListBox class |
| `clsModernStyle` | UserForms style |
| `clsProgresBar` | Progress bar |
| `clsSlider` | Slider |
| `clsTextArea` | TextArea control |
| `clsTextboxMask` | TextBox with mask |
| `fint` | Center form |
| `ftxtbtn` | DropButton for TextBox |
| `ftxtbtnclick` | DropButton click |

### Http (3)

| Snippet | Description |
|---------|-------------|
| `responseTextHttp` | HTTP request |
| `webPageHTML` | HTML page |
| `webPageText` | HTML page text |

### Iff (3)

| Snippet | Description |
|---------|-------------|
| `IfEl` | IF Else |
| `IfEn` | IF |
| `SeCa` | Select Case |

### LogFile (3)

| Snippet | Description |
|---------|-------------|
| `ShowLog` | Show log |
| `WriteErrorLog` | Error log |
| `WriteLog` | Log write |

### TxtFile (9)

| Snippet | Description |
|---------|-------------|
| `Base64` | Base64 encode/decode |
| `base64ToFile` | Base64 to file |
| `fileToBase64` | File to Base64 |
| `loadTextFromTextFile` | Load text from file |
| `Range2CSV` | Range to CSV |
| `saveTextToFile` | Save text to file |
| `TXTAddIntoTXTFile` | Add to file |
| `TXTReadALLFile` | Read all file |
| `TXTWriteFile` | Write to file |

## 🚀 Usage

To use snippets, install the [Macro Tools VBA](https://github.com/vbatools/Addin_MacroToolsVBA) add-in and configure it to use this repository as the snippet source.

## 👤 Author

**VBATools**

## 📄 License

[Apache License 2.0](LICENSE)
