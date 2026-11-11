# Automation Anywhere Utilities 🛠️

Welcome to the Automation Anywhere Utilities repository. This centralized library houses production-ready accelerators designed to standardize automation, enforce strict error contracts, and accelerate delivery across Windows environments.

## 🌿 Repository Structure (Folders)

To keep tasks isolated and maintain clean version control, this repository uses a folder-based architecture.

*   **utilities_windows**: Contains the core logging utility (`TB_WindowsUtil_LogToFile`) for structured, pipe-delimited audit trails.
*   **util_w_excel**: Contains reusable workflows (`Util_Excel_ReadFromLocal`, `Util_Excel_WriteToLocal`) for memory-safe DataTable extraction and writing.
*   **util_w_outlook**: Contains the email dispatcher (`Util_Outlook_SendEmail`) for routing HTML-formatted process notifications.

## ⚙️ Engineering Standards

All components across our folders adhere to strict enterprise guidelines:

*   **Dynamic Configuration**: No hardcoded credentials or local environment paths.
*   **Strict Error Contracts**: Clear separation between business rule exceptions and system exceptions for reliable unattended execution.
*   **Scalability**: Modular architecture designed to plug seamlessly into high-volume enterprise frameworks.

---

## 🤖 Bot Analysis & Documentation

> **Scope:** This document analyzes the four Automation Anywhere bot exports included in the supplied archive. It describes their apparent purpose, inputs/outputs, execution flow, dependencies, observations, and recommended improvements.
>
> **Source format:** The files are Automation Anywhere JSON bot definitions (bot code version 7), not executable source code. The analysis is based on the exported node structure and variable metadata. Some `If` condition details are not fully represented in the visible export structure, so conditional behavior is described cautiously.

### Overview

| Bot | Purpose | Main dependency |
|---|---|---|
| `TB_WindowsUtil_LogToFile` | Creates/appends structured process log entries | File, Folder, Datetime, LogToFile |
| `Util_Excel_ReadFromLocal` | Reads a worksheet into an Automation Anywhere DataTable | Excel MS package |
| `Util_Excel_WriteToLocal` | Creates an Excel workbook and writes a DataTable | Excel MS package |
| `Util_Outlook_SendEmail` | Sends process notification emails through Outlook | Email package / Outlook |

All four exports have no configured triggers. They appear to be reusable task-bot utilities intended to be invoked by a parent automation or another bot. All use an ErrorHandler `try` wrapper. The export does not expose enough detail to confirm the configured exception-handling behavior (for example, catch/finally actions or error propagation).

---

## 1. `TB_WindowsUtil_LogToFile`

**Location:** `Utilities_Windows/TB_WindowsUtil_LogToFile`

### Purpose
Builds a pipe-delimited log record containing a timestamp, machine name, log level, process name, and message, then appends it to a log file using UTF-8 encoding.

### Inputs and outputs

| Variable | Direction | Type | Description / default |
|---|---|---|---|
| `in_str_LogDirectory` | Input | String | Target log folder; default `C:\Logs\ProcessName` |
| `in_str_LogMessage` | Input | String | Message to log |
| `in_str_ProcessName` | Input | String | Process name |
| `in_str_LogLevel` | Input | String | Expected level: Info, Warn, Error, Fatal |
| `in_str_TimeFormat` | Input | String | Timestamp format; default `yyyyMMdd_HHmmss` |
| `io_str_FullLogFilePath` | Input/Output | String | Full log file path |
| `out_bool_IsLogUpdated` | Output | Boolean | Indicates whether the log was updated |
| `out_str_LogStatus` | Output | String | Status text |

*Internal variables: `str_CurrentTimestamp` and `str_LogEntry`.*

### Execution flow
1. Formats the system date using `in_str_TimeFormat`.
2. Checks whether a full log file path was supplied.
3. The nested file-existence check appears intended to determine whether the log file already exists and to handle the file/folder setup.
4. Builds a record in this format: `[timestamp]|[machine]|[level]|[process]|[message]`
5. Appends the record to the target file with UTF-8 encoding and with the LogToFile action's automatic timestamp disabled.
6. Sets the success Boolean and status text.
7. Displays a MessageBox containing the status and Boolean result.

### Suggested log record contract
```text
timestamp|machine|level|process|message
```
Use a consistent timestamp/time-zone convention and document whether timestamps are local time or UTC.

---

## 2. `Util_Excel_ReadFromLocal`

**Location:** `Utilities_Windows/Util_W_Excel/Util_Excel_ReadFromLocal`

### Purpose
Opens a workbook from a local or shared-drive path, reads all rows from a specified worksheet as cell text, returns the result as a DataTable, and closes the workbook session.

### Inputs and outputs

| Variable | Direction | Type | Description |
|---|---|---|---|
| `in_str_FilePath` | Input | String | Excel workbook path |
| `in_str_SheetName` | Input | String | Worksheet name |
| `out_dt_Input` | Output | Table | DataTable containing the read data |
| `out_str_Status` | Output | String | Status text |
| `excelSession_Workbook` | Internal | Session | Excel MS workbook session |

### Execution flow
1. Opens the workbook using the desktop-file-path option.
2. Opens the requested worksheet by name.
3. Reads all rows using `GetMultipleCells` with `READ_CELL_TEXT`.
4. Closes the workbook session.
5. Builds a status/detail string containing the file path and returned row count.

### Notable configuration
- `containsHeader = false`
- `fileAccessMode = EDIT`
- `excludeHiddenSheets = false`
- `disableUpdateLinks = false`
- The read operation uses `ALL_ROWS` and `READ_CELL_TEXT`.

---

## 3. `Util_Excel_WriteToLocal`

**Location:** `Utilities_Windows/Util_W_Excel/Util_Excel_WriteToLocal`

### Purpose
Creates an Excel workbook, writes an input DataTable to a named worksheet beginning at cell `A1`, saves the workbook, and closes the session.

### Inputs and outputs

| Variable | Direction | Type | Description |
|---|---|---|---|
| `in_str_FilePath` | Input | String | Destination workbook path |
| `in_str_SheetName` | Input | String | Worksheet name |
| `in_dt_Input` | Input | Table | DataTable to write |
| `out_str_Status` | Output | String | Status text |
| `out_bool_IsExcelUpdateFailed` | Output | Boolean | Failure indicator |
| `excelSession_Workbook` | Internal | Session | Excel MS workbook session |

### Execution flow
1. Creates a spreadsheet at the supplied path.
2. Uses `isOverwrite = false`.
3. Writes the input DataTable to the specified worksheet, starting at `A1`.
4. Does not retain existing cell formatting.
5. Saves and closes the workbook.
6. Builds a status/detail string containing the destination path.
7. Sets the failure Boolean to `true` in the visible action sequence.

---

## 4. `Util_Outlook_SendEmail`

**Location:** `Utilities_Windows/Util_W_Outlook/Util_Outlook_SendEmail`

### Purpose
Sends an HTML email through Outlook with recipient, CC, subject, body, and optional attachments. The bot contains success and exception HTML templates and a subject-template variable.

### Inputs and outputs

| Variable | Direction | Type | Description / default |
|---|---|---|---|
| `in_str_EmailRecipients` | Input | String | To recipients |
| `in_str_EmailCC` | Input | String | CC recipients |
| `in_str_EmailSubject` | Input | String | Email subject |
| `in_str_EmailType` | Input | String | Described as Success, Exception, or BusinessException |
| `in_str_EmailBody` | Input | String | Email body/detail |
| `in_str_ProcessName` | Input | String | Process name |
| `in_str_Environment` | Input | String | DEV, UAT, or PROD |
| `in_listStr_Attachments` | Input | List of File | Optional attachments |
| `out_str_Status` | Output | String | Status text |
| `out_bool_IsEmailSent` | Output | Boolean | Email-sent indicator |

*Internal variables include success and exception HTML templates and a subject-template variable.*

### Intended execution flow
1. Builds a subject from environment, process name, email subject, and a timestamp placeholder.
2. Selects and personalizes a success or exception HTML template based on the email type.
3. Sends the message through Outlook with To, CC, HTML body format, and optional attachments.

---

## Cross-bot findings

### Strengths
- The bots are separated into focused reusable utilities for logging, Excel read/write, and email notifications.
- Inputs and outputs are generally exposed as bot variables, which supports reuse by parent automations.
- Excel and email operations use explicit package actions rather than opaque scripts.
- The email utility supports HTML content and file attachments.
- The logging utility uses a structured record and UTF-8 append behavior.

## Runtime and package notes

The exports identify these package versions:

| Package | Version(s) observed |
|---|---|
| ErrorHandler | `2.13.0-20241115-120032` |
| Excel_MS | `6.23.20` |
| Email | `3.33.8` |
| LogToFile | `3.11.1` |
| Datetime | `2.17.3` |
| File | `6.16.2` |
| Folder | `6.13.0` |

*Package versions are taken from the supplied exports. Confirm compatibility with the target Automation Anywhere Control Room / Bot Agent environment before deployment.*