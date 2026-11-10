# ReusableComponents_UiPathLibrary

A UiPath **Library** of reusable workflows for Ashling automation projects. It covers the plumbing most automations need: config loading, logging, Excel and file handling, DataTable helpers, SMTP email, PDF utilities, and an invoice PDF splitter. Install it as a dependency and each public workflow shows up as an activity in Studio.

## Project description

| | |
|---|---|
| Project type | Library |
| Target framework | Windows |
| Expression language | VB.NET |
| Studio version | 26.0.204 |
| Project version | 1.0.0 |
| Workflows | 32 (5 Common, 2 DataTable, 1 Email, 4 Excel, 5 Files, 2 Framework, 2 PDF, 11 PDF Split) |

### Dependencies

| Package | Version | Used by |
|---|---|---|
| UiPath.System.Activities | 26.10.0 | All workflows |
| UiPath.Excel.Activities | 3.6.1 | ReadExcelFile, SaveExcelFile, FormatExcelSheet, InitAllSettings |
| UiPath.PDF.Activities | 4.3.0 | PDF Split workflows (Read PDF Text, Extract Page Range, Page Count) |
| itext7 | 9.8.0 | GetPdfOrientation, RotatePdf |

### Requirements on the robot machine

- Windows robot (library targets Windows).
- **Desktop Microsoft Excel** for `DeleteExcelSheet` and `FormatExcelSheet` (COM interop / VBA).
- A reachable **SMTP server** for `SendEmail_SMTP`: an anonymous relay, or one that accepts a user name and password (optionally over SSL/TLS).
- **Orchestrator connection** for `InitAllSettings` if your config uses assets or credentials.
- Read/write access to any folders you pass to the file, Excel and PDF workflows.

## Components

### Common utilities

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `ConvertBytesToPdf` | Rebuilds a PDF file from a DataTable where each row holds a byte-array chunk and writes it to disk. | `in_str_InputFilePath`, `in_dt_ArrayBytes` | – |
| `ConvertPdfToBytes` | Reads a PDF and returns the full byte array, plus a DataTable with the bytes split into 1,024-byte chunks (one chunk per row). | `in_str_InputFilePath` | `out_arrByte_Data`, `out_dt_arrayBytes` |
| `CreateHtmlTable` | Converts a DataTable into an HTML table string (header + rows), mainly for email bodies. | `in_dt_Input` | `out_HtmlTable` |
| `CreateSummaryReport` | Builds a 2-column summary DataTable (Entity, Count) from a report table: Success count vs Fail/Exception/BusinessRuleException count. | `in_dt_Report` | `out_dt_Summary` |
| `CustomLogging` | Wrapper around Log Message. Can pull log text by code from Config, supports Info/Error/Warn/Fatal/Trace, and adds up to 4 custom log fields (Project, Application, Workflow, Extra Logs) visible in Orchestrator. | `in_bool_PickLogFromConfig`, `in_Config`, `in_int_LogCode`, `in_str_LogLevel`, `in_str_LogMessage`, `in_str_ProjectName`, `in_str_ApplicationName`, `in_str_WorkflowName`, `in_str_ExtraLogs` | – |

### DataTable helpers

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `AddDataRow` | Adds one row to a DataTable from a string array. | `in_arrRow_RowValues`, `io_dt_Input` | `io_dt_Input` |
| `AddNewColumns` | Adds columns to an existing DataTable, or creates a new DataTable from a list of column names. | `in_arrStr_ColumnNames`, `in_bool_isNewDatatable`, `io_dt_Input` | `io_dt_Input` |

### Email

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `SendEmail_SMTP` | Sends an HTML email over SMTP. Recipients and attachments are split on ";". Missing attachment files are skipped silently. Optional authentication (user name + SecureString password) and SSL/TLS. Releases attachment files after sending. | `in_str_MailPort`, `in_str_EmailSender`, `in_str_EmailServer`, `in_str_EmailBody`, `in_str_Attachment`, `in_str_EmailSubject`, `in_str_EmailTo`, `in_str_Username`, `in_secureStr_Password`, `in_bool_EnableSsl` | – |

### Excel

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `DeleteExcelSheet` | Deletes a named sheet from an Excel file, saves and closes the workbook. | `in_str_ExcelFilePath`, `in_str_ExcelSheetName` | – |
| `FormatExcelSheet` | Formats one or more sheets by running a VBA macro (autofit, wrap text, autofilter on row 1). Can run any custom VBA method from your own file. | `in_str_ExcelFilePath`, `in_str_VBAfilePath`, `in_str_VBAmethodName`, `in_bool_ChangeVBAfilePath`, `in_arrStr_Sheets` | – |
| `ReadExcelFile` | Reads a sheet or range from an Excel file into a DataTable. Throws on failure. | `in_str_SheetName`, `in_str_InputFilePath`, `in_bool_HasHeaders`, `in_str_Range` | `out_dt_ExcelData` |
| `SaveExcelFile` | Writes a DataTable to a sheet in an Excel file, with or without headers, at a given start range. | `in_str_FilePath`, `in_bool_HasHeaders`, `in_str_Range`, `in_str_SheetName`, `in_dt_InputData` | – |

### Files and folders

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `CopyFilesAndFolder` | Copies a file or a folder (chosen by in_str_FileOrFolder = "File" or "Folder"), with overwrite option. | `in_str_DestinationPath`, `in_str_SourcePath`, `in_bool_ShouldOverwrite`, `in_str_FileOrFolder` | – |
| `CreateTempFolder` | Creates a root folder if missing and, optionally, a subfolder under it. Returns the full subfolder path in io_str_SubFolder. | `in_Config`, `in_bool_CreateSubFolder`, `in_str_RootFolderPath`, `io_str_SubFolder` | `io_str_SubFolder` |
| `DeleteTempFolder` | Deletes a folder if it exists, inside Try/Catch. | `in_str_FolderPath` | – |
| `GetListOfMultipleTypeFiles` | Returns a list of file paths and a count for a folder, filtered by multiple patterns (e.g. {"*.xlsx","*.pdf"}). | `in_str_FolderPath`, `in_arrStr_FilesType` | `out_listStr_FileNames`, `out_int_FilesCount` |
| `MoveFilesAndFolder` | Moves a file or a folder (chosen by in_str_FileOrFolder), with overwrite option. | `in_str_FileOrFolder`, `in_bool_ShouldOverwrite`, `in_str_SourcePath`, `in_str_DestinationPath` | – |

### Framework

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `InitAllSettings` | Builds the Config dictionary from an Excel config file (Settings, Constants, Credentials) and merges Orchestrator assets. Asset values overwrite same-name settings/constants. | `in_ConfigFile`, `in_ConfigSheets` | `out_Config` |
| `KillAllProcesses` | Kills every process in a semicolon-separated list using the Kill Process activity. | `in_str_KillProcesses`, `in_Config` | – |

### PDF utilities

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `GetPdfOrientation` | Reads the rotation of every page and returns a text list ("Page 1: Rotation 0°", ...). | `in_str_InputFile` | `out_str_CurrentOrientations` |
| `RotatePdf` | Rotates only the pages you specify, using a dictionary {page: degrees}, and saves to a new file. | `in_str_InputFile`, `in_str_OutputFilePath`, `in_dict_PagesToRotate` | – |

### Invoice PDF splitting

| Workflow | What it does | Inputs | Outputs |
|---|---|---|---|
| `CheckKeywordsExist` *(private)* | Checks whether any keyword from a list appears in a text (case-insensitive). | `in_str_Text`, `io_bool_KeywordFound`, `in_arr_listOfKeywords`, `in_str_SearchingForKeywords` | `io_bool_KeywordFound` |
| `CheckNextPlus1PageIsEmpty` *(private)* | Looks ahead to see whether the following page has no text (e.g. a scanned supporting page). | `int_NextPlusOneCount`, `in_int_TotalPDFpageCount`, `bool_IsEmpty`, `in_str_InputFile`, `str_ReadTempPDFText` | `int_NextPlusOneCount`, `bool_IsEmpty`, `str_ReadTempPDFText`, `int_NextPageCounter`, `io_bool_ExitLoop` |
| `CheckPageCombinations` *(private)* | Matches "Page X of Y" style patterns (from the PageCombinations sheet) against page text to find invoice page ranges. | `in_int_FirstNumber`, `in_int_SecondNumber`, `in_str_PdfText`, `in_dt_PageCombinations` | `out_bool_MatchFound` |
| `CheckVendorName` *(private)* | Checks the PDF text against a vendor list. If a vendor matches, the file is kept as is (not split). | `in_str_FilePath`, `in_dt_Vendors`, `io_listStr_Files`, `io_bool_ShouldSplit` | `io_listStr_Files`, `io_bool_ShouldSplit` |
| `EachPageAsInvoice` *(private)* | Treats each page as its own invoice. Blank following pages are attached to the same invoice. | `io_int_LoopCounter`, `int_NextPageCounter`, `bool_IsEmpty`, `in_int_TotalPDFpageCount`, `str_SplitPDfFileName`, `in_str_SplitFolderPath`, `in_str_InputFile`, `io_listStr_SplitFiles`, `str_ReadTempPDFText` | `io_int_LoopCounter`, `int_NextPageCounter`, `bool_IsEmpty`, `str_SplitPDfFileName`, `str_ReadTempPDFText`, `io_bool_ExitLoop` |
| `GetKeywords` *(private)* | Reads the keyword sheets from an Excel file into a Dictionary(String, String). | `in_str_FilePath` | `out_dict_Keywords` |
| `LoopPdfPages` *(private)* | Main page loop. Decides page by page whether pages form a single-page or multi-page invoice. | `in_str_TempFolderPath`, `in_str_InputFile`, `in_str_SplitFolderPath`, `io_int_LoopCounter`, `in_int_TotalPDFpageCount`, `io_bool_ExitLoop`, `io_listStr_TempFiles`, `io_listStr_SplitFiles`, `in_dt_PageCombinations` | `io_int_LoopCounter`, `io_bool_ExitLoop`, `io_listStr_TempFiles`, `io_listStr_SplitFiles` |
| `MultiplePageAsInvoice` *(private)* | Groups consecutive pages into one invoice using page-combination patterns. | `io_int_LoopCounter`, `int_NextPageCounter`, `in_int_TotalPDFpageCount`, `io_bool_ExitLoop`, `in_str_SplitFolderPath`, `in_str_InputFile`, `io_listStr_SplitFiles`, `in_str_ReadTempPDFText`, `in_dt_PageCombinations` | `io_int_LoopCounter`, `int_NextPageCounter`, `io_bool_ExitLoop` |
| `SavePdfRange` *(private)* | Extracts a validated page range into a new PDF and adds it to the split-file list. | `int_NextPageCounter`, `io_bool_ExitLoop`, `in_str_ReadTempPDFText`, `in_str_InputFile`, `in_int_TotalPDFpageCount`, `io_int_LoopCounter`, `in_str_SplitFolderPath`, `io_listStr_SplitFiles` | `int_NextPageCounter`, `io_bool_ExitLoop`, `io_int_LoopCounter`, `io_listStr_SplitFiles` |
| `SplitPDF` | Top-level splitter. Reads page-combination and vendor tables plus keywords, checks vendor, loops pages, validates output, moves files to the split folder. | `io_str_InvoiceType`, `in_str_InputFile`, `in_str_SplitFolderPath`, `in_str_TempFolderPath` | `io_str_InvoiceType`, `out_listStr_Pdfs` |
| `ValidatePdfs` *(private)* | Filters split PDFs using keyword lists (Pdf, Exclude, PO identifiers) and sets the invoice type. | `in_dict_Keywords`, `io_listStr_SplitFiles`, `io_str_InvoiceType` | `io_listStr_SplitFiles`, `io_str_InvoiceType` |

## Folder structure

```
Common/            CustomLogging, ConvertBytesToPdf, ConvertPdfToBytes, CreateHtmlTable, CreateSummaryReport
Datatable/         AddDataRow, AddNewColumns
Email/             SendEmail_SMTP
Excel/             ReadExcelFile, SaveExcelFile, FormatExcelSheet, DeleteExcelSheet
FilesAndFolders/   CopyFilesAndFolder, MoveFilesAndFolder, CreateTempFolder, DeleteTempFolder, GetListOfMultipleTypeFiles
Framework/         InitAllSettings, KillAllProcesses
PDF/               GetPdfOrientation, RotatePdf
PDF/SplitPdfs/     SplitPDF (public) + 10 private helpers
PDF/Files/         Sample Keywords.xlsx and PageCombinations_And_VendorList.xlsx
CommonFiles/       Sample Config.xlsx, VBA/FormatExcel.txt
```

## Support files

| File | Purpose |
|---|---|
| `CommonFiles/Sample Config.xlsx` | Template for `InitAllSettings`. Sheets: Settings, Constants, Assets, Credentials. |
| `CommonFiles/VBA/FormatExcel.txt` | Default VBA macro used by `FormatExcelSheet` (autofit, wrap text, autofilter on row 1). |
| `PDF/Files/Keywords.xlsx` | Keyword lists for `GetKeywords` / `ValidatePdfs` (sheets: PO_Identifiers, EmailSubject, EmailBody, Pdf, Word, Exclude). |
| `PDF/Files/PageCombinations_And_VendorList.xlsx` | "Page X of Y" patterns and vendor names that must not be split. |

## How the PDF splitter fits together

1. `SplitPDF` reads the page-combination and vendor tables and the keyword file (via `ReadExcelFile` and `GetKeywords`).
2. `CheckVendorName` – if the PDF matches a vendor on the list, it is kept as one file.
3. `LoopPdfPages` walks the pages and calls `CheckPageCombinations`, `EachPageAsInvoice` or `MultiplePageAsInvoice` to decide the page ranges.
4. `SavePdfRange` and `CheckNextPlus1PageIsEmpty` write each range and attach blank/scanned follow-up pages to the same invoice.
5. `ValidatePdfs` filters the split files with keyword lists (`Pdf_Keywords`, `Exclude_Keywords`, `PO_Identifiers_Keywords`) and sets the invoice type.
6. `MoveFilesAndFolder` moves the results to the split folder.

## Notes and known limitations

| # | Note |
|---|---|
| 1 | `SendEmail_SMTP` uses the .NET `SmtpClient` with basic authentication only. If you send through Exchange Online / Microsoft 365, check with your admin first: Microsoft has announced that SMTP AUTH Basic authentication will be disabled by default for existing tenants from the end of December 2026 (admins can still turn it back on), and OAuth is the supported method. For Microsoft 365 use a mail-flow connector (anonymous relay) or the Microsoft 365 / Outlook mail activities instead. |
| 2 | `GetPdfOrientation` returns the rotation of every page as text (`Page 1: Rotation 0°`), not as numbers. Parse it if you need values. |
| 3 | `DeleteExcelSheet` throws a `BusinessRuleException` when the sheet is not found, is the last sheet, or Excel fails. |
| 4 | Excel interop workflows (`DeleteExcelSheet`, `FormatExcelSheet`) need desktop Excel and a Windows robot. |

## Security note: SecureString and `InitAllSettings`

`InitAllSettings` reads each credential from Orchestrator and keeps the password as a `SecureString` in the `Config` dictionary (`Config("<name>_Password")`). It is never converted to plain text inside this library. Workflow Analyzer rule **ST-SEC-009 (SecureString Misusage)** may still raise a warning here, because the `SecureString` is stored in a long-lived dictionary instead of a narrow local scope. This is intentional and accepted: it is the standard framework pattern, and consumers should pass the value straight to activities that accept `SecureString` (for example `in_secureStr_Password` on `SendEmail_SMTP`). Do not cast it to String. If a security review ever requires it, the alternative is to stop storing credentials in `Config` and call Get Credential in each workflow that needs one; that would be a breaking change (major version).

## Usage

1. Publish the library (Studio > Publish) to your Orchestrator libraries feed or a custom NuGet feed.
2. In the consuming project: Manage Packages > install `ReusableComponents_UiPathLibrary`.
3. Search for the workflow name in the Activities panel and drag it in.
4. For config: copy `Sample Config.xlsx`, fill in Settings/Constants/Assets/Credentials, and call `InitAllSettings` first to get the `Config` dictionary.

## Conventions

- Arguments: `in_` / `out_` / `io_` + type prefix, e.g. `in_str_FilePath`, `out_dt_ExcelData`.
- Every workflow has a root annotation (purpose, precondition, postcondition).
- No hardcoded paths, credentials or client data. Pass values as arguments or through Config.
- Renaming an argument is a breaking change for anyone using the library. Bump the major version.
