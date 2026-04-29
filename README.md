# Kw2Pw26 — KWorks / Abraxas Legacy Data Importer

> ⚠️ **Purpose Notice**
> This application is designed **exclusively and explicitly** to import legacy data from
> **KWorks** and **Abraxas** — 1980s practice-management software products that are the same
> application sold under two different names, originally written as a **FoxPro** application.
> It is not a general-purpose CSV importer.

---

## Overview

**Kw2Pw26** is a standalone portable Windows GUI application (.NET 9 / WPF) that reads the
`.CSV` export files produced by KWorks / Abraxas and imports them into a **SQL Server 2022**
database, preserving the original schema and data as faithfully as possible.

---

## Key Features

| Feature | Detail |
|---|---|
| **Legacy-aware schema loading** | Reads `DBF.CSV` and `DBFCHK.CSV` config files exported from KWorks/Abraxas to reconstruct the original FoxPro table structure. |
| **System / Company routing** | Tables flagged `DIR=COMP` are imported into the Company database; all others go into the System database. |
| **Import ordering** | Files are ordered by the `DIR` field — system tables first, then company tables. |
| **Optional auto-ID column** | A `NewID BIGINT IDENTITY(1,1)` column can be prepended to every table for surrogate key support. |
| **Null & blank-row safety** | Completely blank rows are silently skipped; all data columns are created `NULL`-able. |
| **Connection testing** | Built-in test confirms SQL Server connectivity before any import begins. |
| **`.ini` config export** | Saves a `PcaWorks26.ini` file alongside the data for use by companion applications. |
| **Activity console** | Green-on-black scrolling log shows real-time import progress in the GUI. |
| **Timestamped log files** | Optional file logging writes `<timestamp>_Kw2Pw26.log` next to the executable. |
| **STOP support** | Import can be gracefully halted after the current file finishes. |
| **Portable single `.exe`** | Published as a self-contained, single-file executable — no installer required. |

---

## Requirements

- Windows 10 or later (x64)
- SQL Server 2022 (or compatible instance) reachable from the machine
- KWorks / Abraxas `.CSV` exports including `DBF.CSV` and `DBFCHK.CSV`

---

## Quick Start

1. Run `Kw2Pw26.exe`.
2. Enter your SQL Server **connection string** and click **TEST**.
3. Set the **System DB** and **Company DB** names.
4. Browse to the folder containing your KWorks/Abraxas `.CSV` export files.
5. Browse to the folder containing `DBF.CSV` and `DBFCHK.CSV`, then click **CHECK**.
6. Adjust import options as needed (log level, duplicate handling, ID field, etc.).
7. Click **IMPORT**.

---

## Project Structure

```
Kw2Pw26/
├── Icons/              Application icon assets
├── MainWindow.xaml     WPF UI definition
├── MainWindow.xaml.cs  All application logic (import, schema, logging)
├── Kw2Pw26.csproj      Project file (.NET 9 WPF, self-contained publish)
└── Properties/
    └── PublishProfiles/FolderProfile.pubxml   Single-file publish settings
```

---

## Building & Publishing

```powershell
# Build
dotnet build Kw2Pw26.slnx -c Release

# Publish portable single exe
dotnet publish Kw2Pw26/Kw2Pw26.csproj -c Release -r win-x64
```

---

## Dependencies

| Package | Version |
|---|---|
| `Microsoft.Data.SqlClient` | 7.0.1 |
| `CsvHelper` | 33.1.0 |

---

## License

See repository root for license details.
