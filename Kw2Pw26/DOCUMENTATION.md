# Kw2Pw26 — Technical Documentation

> ⚠️ **Purpose Notice**
> This application is designed **exclusively and explicitly** to import legacy data from
> **KWorks** and **Abraxas** — 1980s practice-management software products that are the same
> application sold under two different names, originally written as a **FoxPro** application.
> It is not a general-purpose CSV importer.

---

## Table of Contents

1. [Application Overview](#1-application-overview)
2. [System Requirements](#2-system-requirements)
3. [User Interface Reference](#3-user-interface-reference)
4. [Import Workflow](#4-import-workflow)
5. [Schema Loading — DBF.CSV & DBFCHK.CSV](#5-schema-loading--dbfcsv--dbfchkcsv)
6. [Database Handling](#6-database-handling)
7. [Column & Data Type Mapping](#7-column--data-type-mapping)
8. [Null & Blank-Row Handling](#8-null--blank-row-handling)
9. [Optional NewID Column](#9-optional-newid-column)
10. [INI Config Export](#10-ini-config-export)
11. [Logging](#11-logging)
12. [Activity Console](#12-activity-console)
13. [STOP / Cancellation](#13-stop--cancellation)
14. [Publish & Deployment](#14-publish--deployment)
15. [Code Structure Reference](#15-code-structure-reference)

---

## 1. Application Overview

Kw2Pw26 bridges **KWorks / Abraxas** legacy FoxPro data exports and a modern **SQL Server 2022**
database. KWorks and Abraxas are the same practice-management application — released under two
brand names during the 1980s — originally built on Microsoft FoxPro. When a site migrates away
from that platform, their data can be exported as `.CSV` files. Kw2Pw26 reads those files and
the accompanying schema descriptors (`DBF.CSV`, `DBFCHK.CSV`) to recreate the original table
structure inside SQL Server and populate it with the exported records.

---

## 2. System Requirements

| Requirement | Minimum |
|---|---|
| OS | Windows 10 x64 or later |
| Runtime | None — self-contained executable |
| SQL Server | SQL Server 2022 (or a compatible instance) |
| Permissions | `dbcreator` role (or equivalent) on the SQL Server instance |
| Source data | KWorks/Abraxas `.CSV` exports + `DBF.CSV` + `DBFCHK.CSV` |

---

## 3. User Interface Reference

### Connection Row
- **Connection String** — Full ADO.NET connection string pointing at the SQL Server instance.
  The importer connects to `master` first to create databases if needed, then switches to the
  target database.
- **TEST** — Verifies connectivity. Must succeed before IMPORT is enabled.

### Folder Rows
- **CSV Folder** — Folder containing the KWorks/Abraxas `.CSV` data exports.
- **CSV Config Folder** — Folder containing `DBF.CSV` and `DBFCHK.CSV`.
- **CHECK** — Validates that `DBF.CSV` and `DBFCHK.CSV` are present and parseable. Must succeed
  before IMPORT is enabled.

### Database Rows
- **System DB Name** — Name for the system-level SQL Server database.
- **Company DB Name** — Name for the company-level SQL Server database.

### INI Row
- **INI Folder** — Destination folder for `PcaWorks26.ini`.
- **INI Filename** — Name of the `.ini` file to write (default: `PcaWorks26.ini`).

### Import Options
- **If table exists** — `Skip`, `Overwrite (ask)`, or `Overwrite (auto)`.
- **Import System Tables** / **Import Company Tables** — Toggle each group independently.
- **Log Level** — `None`, `Errors Only`, or `All`.
- **Max errors before abort** — Stop importing after this many row-level errors.
- **Batch progress size** — How many rows between activity console progress updates.
- **Create ID field?** *(checked by default)* — Prepends a `NewID BIGINT IDENTITY(1,1)` column
  to every created table.

### Bottom Bar
| Button | State | Action |
|---|---|---|
| **STOP** | Enabled during import | Signals a graceful stop after the current file finishes |
| **CLEAR** | Always | Resets all controls to defaults |
| **IMPORT** | After TEST + CHECK pass | Begins the import process |
| **VIEW LOG** | After any log entry is written | Opens the latest log in Notepad |
| **EXIT** | Always | Closes the application |

---

## 4. Import Workflow

```
User clicks IMPORT
  │
  ├─ Validate inputs (connection string, folder paths, DB names)
  ├─ Connect to master → create System DB if missing → create Company DB if missing
  ├─ Read DBF.CSV + DBFCHK.CSV → build in-memory schema (_dbfFields, _dbfTables)
  ├─ Enumerate .CSV files in CSV Folder; match each to a table entry in DBF.CSV
  ├─ Sort: system tables first, then company tables (by DIR field)
  │
  └─ For each CSV file (in order):
        ├─ Check cancellation token (STOP requested?)
        ├─ Determine target DB (system or company)
        ├─ Connect to target DB
        ├─ Handle existing table per "If table exists" option
        ├─ CreateTableFromSchema → DDL based on DBF.CSV/DBFCHK.CSV field definitions
        └─ ImportFile → stream CSV rows → InsertRow (skip blank rows, map nulls)
              └─ Log batch progress every N rows to activity console
```

---

## 5. Schema Loading — DBF.CSV & DBFCHK.CSV

`DBF.CSV` describes each **table**:

| Column | Meaning |
|---|---|
| `FILE` | Base name of the `.CSV` data file (without extension) |
| `DIR` | `COMP` → Company DB, anything else → System DB |
| *(other columns)* | Metadata used for ordering and routing |

`DBFCHK.CSV` describes each **column** within each table:

| Column | Meaning |
|---|---|
| `FILE` | Matches `FILE` in `DBF.CSV` |
| `FIELD` | Column name |
| `TYPE` | FoxPro data type (`C`, `N`, `D`, `L`, `M`, …) |
| `LEN` | Field length |
| `DEC` | Decimal places (for numeric types) |

The loader builds `_dbfTables` (keyed by `FILE`) and `_dbfFields` (list of field definitions per
table) which drive all `CREATE TABLE` DDL generation.

---

## 6. Database Handling

On each import run the application:

1. Opens a connection to `master` using the supplied connection string (with the `Initial Catalog`
   replaced by `master`).
2. Issues `CREATE DATABASE [name] IF NOT EXISTS` for both the System and Company databases.
3. Reconnects to the appropriate target database for each file group.

This ensures the target databases exist before any schema or data operations are attempted,
preventing the *"Cannot open database … requested by the login. The login failed"* error that
arises when pointing directly at a non-existent database.

---

## 7. Column & Data Type Mapping

FoxPro types are mapped to SQL Server types as follows:

| FoxPro TYPE | SQL Server type |
|---|---|
| `C` (Character) | `NVARCHAR(len)` |
| `N` (Numeric, dec = 0) | `INT` |
| `N` (Numeric, dec > 0) | `DECIMAL(len, dec)` |
| `D` (Date) | `DATE` |
| `T` (DateTime) | `DATETIME` |
| `L` (Logical) | `BIT` |
| `M` (Memo) | `NVARCHAR(MAX)` |
| *(anything else)* | `NVARCHAR(255)` |

Fields whose name contains `MONEY`, `AMT`, `AMOUNT`, `PRICE`, or `COST` are mapped to
`DECIMAL(18,2)` regardless of the FoxPro type.

All columns are created **nullable** (`NULL`) to maximise import resilience.

---

## 8. Null & Blank-Row Handling

- **Blank rows** — A CSV row where every field is empty or whitespace is detected before any
  insert attempt and silently skipped. No error is logged.
- **Empty field values** — Any field value that is empty or whitespace is inserted as SQL `NULL`
  (`DBNull.Value`) rather than an empty string, preventing type-conversion errors on numeric,
  date, and bit columns.

---

## 9. Optional NewID Column

When **Create ID field?** is checked (default: on):

```sql
[NewID] BIGINT IDENTITY(1,1) NOT NULL
```

This column is added as the **first column** of every created table. SQL Server assigns a unique
auto-incrementing integer to each inserted row. `BIGINT` supports values up to
9,223,372,036,854,775,807, comfortably covering any legacy dataset size.

> Note: because the column is `IDENTITY`, it must **not** be included in the `INSERT` column
> list — the application handles this automatically.

---

## 10. INI Config Export

After a successful import, Kw2Pw26 writes a `PcaWorks26.ini` file to the configured INI folder.
This file contains the SQL Server connection string and the System / Company database names, and
is consumed by companion PcaWorks26 applications.

Example output:

```ini
[Database]
ConnectionString=Server=.\SQLEXPRESS;Integrated Security=True;
SystemDatabase=PcaWorks26
CompanyDatabase=PcaWorks26Co
```

---

## 11. Logging

| Log Level | What is recorded |
|---|---|
| `None` | Nothing written to disk |
| `Errors Only` | Row-level errors and abort events |
| `All` | Full verbose log matching the activity console |

Log files are created **lazily** (only when the first entry is written) and placed **beside the
executable** using the naming pattern:

```
<yyyyMMdd_HHmmss>_Kw2Pw26.log
```

**VIEW LOG** opens the most recent log file in Notepad. The button is disabled until at least one
log file exists in that location.

---

## 12. Activity Console

The black panel with green `Consolas` text at the bottom of the window streams real-time progress
during import:

- File being processed and its target database
- Table creation / skip / overwrite events
- Batch progress (e.g., `Records 0–2,000 inserted…`)
- Row-level errors
- Completion or abort summary

The console auto-scrolls to the latest entry using `TextBox.ScrollToEnd()` called on the UI
dispatcher after each append.

---

## 13. STOP / Cancellation

Clicking **STOP** during an import sets a `CancellationToken`. The import loop checks this token
**between files** and raises `ImportAbortException` when it is signalled, causing the import to
finish the current file cleanly and then stop. STOP is disabled when no import is in progress.

---

## 14. Publish & Deployment

The application is published as a **self-contained single-file executable** for `win-x64`:

```
PublishSingleFile=true
SelfContained=true
RuntimeIdentifier=win-x64
IncludeNativeLibrariesForSelfExtract=true
PublishReadyToRun=true
```

No .NET runtime installation is required on the target machine. Simply copy `Kw2Pw26.exe` to any
location and run it.

---

## 15. Code Structure Reference

| Symbol | Location | Purpose |
|---|---|---|
| `BtnTest_Click` | `MainWindow.xaml.cs` | Test SQL Server connection |
| `BtnBrowse_Click` | `MainWindow.xaml.cs` | Browse for CSV data folder |
| `BtnCheck_Click` | `MainWindow.xaml.cs` | Validate and load DBF/DBFCHK config |
| `LoadDbfConfig` | `MainWindow.xaml.cs` | Parse DBF.CSV and DBFCHK.CSV into memory |
| `BtnImport_Click` | `MainWindow.xaml.cs` | Validate and start import on background thread |
| `RunImport` | `MainWindow.xaml.cs` | Orchestrate full import sequence |
| `EnsureDatabase` | `MainWindow.xaml.cs` | Create database if it does not exist |
| `ImportFile` | `MainWindow.xaml.cs` | Stream one CSV file into SQL Server |
| `InsertRow` | `MainWindow.xaml.cs` | Insert a single CSV row, handling nulls |
| `CreateTableFromSchema` | `MainWindow.xaml.cs` | Generate and execute CREATE TABLE DDL |
| `MapSqlType` | `MainWindow.xaml.cs` | FoxPro → SQL Server type mapping |
| `BtnStop_Click` | `MainWindow.xaml.cs` | Signal cancellation token |
| `BtnViewLog_Click` | `MainWindow.xaml.cs` | Open latest log in Notepad |
| `AppendConsole` | `MainWindow.xaml.cs` | Thread-safe write + scroll to activity console |
| `AppendLog` | `MainWindow.xaml.cs` | Write entry to log file (lazy creation) |
| `DbfTableDef` | `MainWindow.xaml.cs` | Model for a DBF.CSV table row |
| `DbfFieldDef` | `MainWindow.xaml.cs` | Model for a DBFCHK.CSV field row |
| `CsvFileItem` | `MainWindow.xaml.cs` | Observable item in the file list |
| `ImportAbortException` | `MainWindow.xaml.cs` | Signals a user-requested stop |
