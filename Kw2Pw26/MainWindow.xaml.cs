using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;

namespace Kw2Pw26;

public partial class MainWindow : Window
{
    // ── Data ──────────────────────────────────────────────────────────────
    private readonly ObservableCollection<CsvFileItem> _files = [];

    // Log path is null until the first write in a session; created lazily beside the .exe.
    private string? _logPath;

    // Cancellation source for the active import; null when no import is running.
    private CancellationTokenSource? _cts;

    /// <summary>Schema rows from DBFCHK.CSV, keyed by DBF name (upper-case).</summary>
    private Dictionary<string, List<DbfFieldDef>> _dbfFields = [];

    /// <summary>Table routing rows from DBF.CSV, keyed by DBF name (upper-case).</summary>
    private Dictionary<string, DbfTableDef> _dbfTables = [];

    public MainWindow()
    {
        InitializeComponent();
        FileList.ItemsSource = _files;
        // VIEW LOG starts disabled; enabled only once a log file has been created
        BtnViewLog.IsEnabled = false;
    }

    // ── Connection String TEST ─────────────────────────────────────────────
    private async void BtnTest_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Testing connection…");
        AppendConsole("Testing SQL Server connection…");
        try
        {
            await using var conn = new SqlConnection(TxtConnectionString.Text.Trim());
            await conn.OpenAsync();
            AppendConsole($"Connection successful. Server: {conn.DataSource}, Version: {conn.ServerVersion}");
            SetStatus("✔ Connection successful.");
        }
        catch (Exception ex)
        {
            AppendConsole($"Connection FAILED: {ex.Message}");
            SetStatus($"✘ Connection failed: {ex.Message}");
        }
    }

    // ── Folder Browse ──────────────────────────────────────────────────────
    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        // Use the modern folder picker via Win32 registry workaround (WPF has no FolderBrowserDialog)
        var dlg = new OpenFolderDialog
        {
            Title = "Select folder containing CSV files"
        };

        if (dlg.ShowDialog(this) != true) return;

        var folder = dlg.FolderName;
        TxtFolder.Text = folder;
        LoadCsvFiles(folder);
    }

    private void LoadCsvFiles(string folder)
    {
        _files.Clear();
        foreach (var file in Directory.EnumerateFiles(folder, "*.csv", SearchOption.TopDirectoryOnly)
                                      .OrderBy(f => f))
        {
            _files.Add(new CsvFileItem(Path.GetFileName(file), file));
        }
        string msg = _files.Count == 0 ? "No CSV files found in selected folder." : $"{_files.Count} CSV file(s) found.";
        SetStatus(msg);
        AppendConsole(_files.Count == 0
            ? $"CSV folder selected: {folder} — no .csv files found."
            : $"CSV folder: {folder} — {_files.Count} file(s) loaded: {string.Join(", ", _files.Select(f => f.FileName))}");
    }

    // ── File List Buttons ──────────────────────────────────────────────────
    private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _files) f.IsSelected = true;
    }

    private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
    {
        foreach (var f in _files) f.IsSelected = false;
    }

    private void BtnClearFiles_Click(object sender, RoutedEventArgs e)
    {
        _files.Clear();
        TxtFolder.Text = string.Empty;
        SetStatus("File list cleared.");
    }

    // ── .ini Folder Browse ─────────────────────────────────────────────────
    private void BtnBrowseIni_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Select folder to save PcaWorks26.ini"
        };
        if (dlg.ShowDialog(this) == true)
            TxtIniFolder.Text = dlg.FolderName;
    }

    // ── CSV Config Folder Browse + CHECK ───────────────────────────────────
    private void BtnBrowseCsvConfig_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Select folder containing DBFCHK.CSV and DBF.CSV"
        };
        if (dlg.ShowDialog(this) == true)
        {
            TxtCsvConfigFolder.Text = dlg.FolderName;
            // Reset import readiness when folder changes
            BtnImport.IsEnabled = false;
            _dbfFields.Clear();
            _dbfTables.Clear();
        }
    }

    private void BtnCheck_Click(object sender, RoutedEventArgs e)
    {
        var folder = TxtCsvConfigFolder.Text.Trim();
        if (string.IsNullOrWhiteSpace(folder))
        {
            SetStatus("Please select a CSV config folder first.");
            return;
        }

        var dbfchkPath = Path.Combine(folder, "DBFCHK.CSV");
        var dbfPath    = Path.Combine(folder, "DBF.CSV");

        var missing = new List<string>();
        if (!File.Exists(dbfchkPath)) missing.Add("DBFCHK.CSV");
        if (!File.Exists(dbfPath))    missing.Add("DBF.CSV");

        if (missing.Count > 0)
        {
            BtnImport.IsEnabled = false;
            SetStatus($"✘ Missing in selected folder: {string.Join(", ", missing)}");
            return;
        }

        try
        {
            LoadDbfConfig(dbfchkPath, dbfPath);
            BtnImport.IsEnabled = true;
            string msg = $"✔ Config loaded — {_dbfTables.Count} table(s), {_dbfFields.Values.Sum(l => l.Count)} field definition(s). IMPORT enabled.";
            SetStatus(msg);
            AppendConsole($"Config CHECK passed. DBF.CSV: {_dbfTables.Count} table routing entries. DBFCHK.CSV: {_dbfFields.Values.Sum(l => l.Count)} field definitions across {_dbfFields.Count} table(s). IMPORT enabled.");
        }
        catch (Exception ex)
        {
            BtnImport.IsEnabled = false;
            SetStatus($"✘ Error reading config CSV files: {ex.Message}");
            AppendConsole($"Config CHECK failed: {ex.Message}");
        }
    }

    /// <summary>Reads DBFCHK.CSV and DBF.CSV into in-memory dictionaries.</summary>
    private void LoadDbfConfig(string dbfchkPath, string dbfPath)
    {
        var csvCfg = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            MissingFieldFound = null,
            BadDataFound = null,
            HeaderValidated = null
        };

        // ── DBFCHK.CSV ──
        _dbfFields = [];
        using (var reader = new StreamReader(dbfchkPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
        using (var csv = new CsvReader(reader, csvCfg))
        {
            foreach (var row in csv.GetRecords<dynamic>())
            {
                var dict = (IDictionary<string, object>)row;
                string dbf   = GetField(dict, "DBF").ToUpperInvariant();
                string field = GetField(dict, "FIELD");
                string type  = GetField(dict, "TYPE").ToUpperInvariant();
                int.TryParse(GetField(dict, "LEN"), out int len);
                int.TryParse(GetField(dict, "DEC"), out int dec);
                string dir   = GetField(dict, "DIR");

                if (!_dbfFields.TryGetValue(dbf, out var list))
                {
                    list = [];
                    _dbfFields[dbf] = list;
                }
                list.Add(new DbfFieldDef(field, type, len, dec, dir));
            }
        }

        // ── DBF.CSV ──
        _dbfTables = [];
        using (var reader = new StreamReader(dbfPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
        using (var csv = new CsvReader(reader, csvCfg))
        {
            foreach (var row in csv.GetRecords<dynamic>())
            {
                var dict = (IDictionary<string, object>)row;
                string dbf = GetField(dict, "DBF").ToUpperInvariant();
                string dir = GetField(dict, "DIR").ToUpperInvariant();
                bool isCompany = dir == "COMP";
                _dbfTables[dbf] = new DbfTableDef(dbf, isCompany);
            }
        }
    }

    private static string GetField(IDictionary<string, object> row, string key)
    {
        // Case-insensitive lookup
        var pair = row.FirstOrDefault(kv => string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase));
        return pair.Value?.ToString()?.Trim() ?? string.Empty;
    }

    // ── Bottom Buttons ─────────────────────────────────────────────────────
    private void BtnClear_Click(object sender, RoutedEventArgs e)
    {
        TxtConnectionString.Text = "Server=.;Database=MyDb;Trusted_Connection=True;TrustServerCertificate=True;";
        TxtFolder.Text = string.Empty;
        _files.Clear();
        CmbDatabaseExists.SelectedIndex = 0;
        CmbDatabaseLog.SelectedIndex = 0;
        CmbTableExists.SelectedIndex = 0;
        CmbTableLog.SelectedIndex = 0;
        CmbRecordExists.SelectedIndex = 0;
        CmbRecordLog.SelectedIndex = 0;
        TxtMaxErrors.Text = "10";
        CmbImportSystemTables.SelectedIndex = 0;
        TxtSystemDbName.Text = "PcaWorks26";
        CmbImportCompanyTables.SelectedIndex = 0;
        TxtCompanyDbName.Text = "Pca2026";
        CmbSaveConfig.SelectedIndex = 0;
        TxtIniFolder.Text = string.Empty;
        TxtCsvConfigFolder.Text = string.Empty;
        _dbfFields.Clear();
        _dbfTables.Clear();
        _logPath = null;
        BtnImport.IsEnabled = false;
        BtnViewLog.IsEnabled = false;
        ChkCreateIdField.IsChecked = true;
        SetStatus("Form cleared.");
        TxtConsole.Clear();
        AppendConsole("Form cleared — ready.");
    }

    private async void BtnImport_Click(object sender, RoutedEventArgs e)
    {
        var selected = _files.Where(f => f.IsSelected).ToList();
        if (selected.Count == 0)
        {
            SetStatus("No files selected for import.");
            return;
        }

        var connStr = TxtConnectionString.Text.Trim();
        if (string.IsNullOrWhiteSpace(connStr))
        {
            SetStatus("Please enter a connection string.");
            return;
        }

        int maxErrors = ParseMaxErrors();
        var options = BuildOptions();

        // Save .ini config before import if requested
        if (CmbSaveConfig.SelectedIndex == 0)
        {
            var iniFolder = TxtIniFolder.Text.Trim();
            if (string.IsNullOrWhiteSpace(iniFolder))
            {
                SetStatus("Please choose a folder for PcaWorks26.ini, or set 'Save DB config?' to No.");
                return;
            }
            try
            {
                SaveIniFile(iniFolder, connStr, TxtSystemDbName.Text.Trim(), TxtCompanyDbName.Text.Trim());
            }
            catch (Exception ex)
            {
                SetStatus($"Failed to save .ini file: {ex.Message}");
                return;
            }
        }

        BtnImport.IsEnabled = false;
        BtnStop.IsEnabled = true;
        // Fresh timestamped log file for each import run
        _logPath = null;
        BtnViewLog.IsEnabled = false;
        SetStatus("Importing…");
        AppendConsole("─────────────────────────────────────────");
        AppendConsole($"Import started. {selected.Count} file(s) selected.");
        AppendConsole($"System DB: {TxtSystemDbName.Text.Trim()}  |  Company DB: {TxtCompanyDbName.Text.Trim()}");

        _cts = new CancellationTokenSource();
        try
        {
            await Task.Run(() => RunImport(connStr, selected, options, maxErrors, _cts.Token));
            if (_cts.IsCancellationRequested)
            {
                SetStatus("Import stopped by user.");
                AppendConsole("Import stopped by user.");
            }
            else
            {
                SetStatus("Import complete. See log for details.");
                AppendConsole("Import complete.");
            }
            AppendConsole("─────────────────────────────────────────");
        }
        catch (Exception ex)
        {
            SetStatus($"Import aborted: {ex.Message}");
            AppendConsole($"Import ABORTED: {ex.Message}");
            AppendConsole("─────────────────────────────────────────");
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            BtnImport.IsEnabled = true;
            BtnStop.IsEnabled = false;
        }
    }

    private void BtnStop_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        BtnStop.IsEnabled = false;
        SetStatus("Stop requested — finishing current file…");
        AppendConsole("Stop requested — will halt after current file.");
    }

    private void BtnViewLog_Click(object sender, RoutedEventArgs e)
    {
        // Use current session's log, or fall back to the most recent log in the exe folder.
        string? path = _logPath;

        if (path is null || !File.Exists(path))
        {
            path = Directory
                .EnumerateFiles(AppContext.BaseDirectory, "*_Kw2Pw26.log")
                .OrderByDescending(f => f)
                .FirstOrDefault();
        }

        if (path is null || !File.Exists(path))
        {
            SetStatus("No log file found.");
            return;
        }

        System.Diagnostics.Process.Start(
            new System.Diagnostics.ProcessStartInfo("notepad.exe", path) { UseShellExecute = true });
    }

    private void BtnExit_Click(object sender, RoutedEventArgs e) => Close();

    // ── Import Logic ───────────────────────────────────────────────────────
    private void RunImport(string connStr, List<CsvFileItem> files, ImportOptions opts, int maxErrors, CancellationToken ct)
    {
        int totalErrors = 0;

        // Parse base connection string once; we will switch the Database= part per file
        var baseBuilder = new SqlConnectionStringBuilder(connStr);

        // Sort: system tables first (DIR != COMP), then company tables, alphabetically within each group.
        // Files not found in DBF.CSV are treated as system tables.
        var ordered = files
            .OrderBy(f =>
            {
                string key = Path.GetFileNameWithoutExtension(f.FileName).ToUpperInvariant();
                return _dbfTables.TryGetValue(key, out var def) && def.IsCompany ? 1 : 0;
            })
            .ThenBy(f => f.FileName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        AppendLog($"[INFO] Import order ({ordered.Count} file(s)): " +
                  string.Join(", ", ordered.Select(f => f.FileName)));
        AppendConsole($"Import order ({ordered.Count} file(s)): {string.Join(", ", ordered.Select(f => f.FileName))}");

        foreach (var file in ordered)
        {
            // Check for stop request between files — as soon as the current file finishes.
            if (ct.IsCancellationRequested) return;

            try
            {
                // Determine target database from DBF.CSV routing; fall back based on user options
                string tableKey = Path.GetFileNameWithoutExtension(file.FileName).ToUpperInvariant();
                bool isCompany = _dbfTables.TryGetValue(tableKey, out var tableDef) && tableDef.IsCompany;

                if (isCompany && !opts.ImportCompanyTables)
                {
                    AppendLog($"[SKIP] {file.FileName} → company table, Import Company Tables = No.");
                    AppendConsole($"SKIP {file.FileName} — company table, Import Company Tables = No.");
                    continue;
                }
                if (!isCompany && !opts.ImportSystemTables)
                {
                    AppendLog($"[SKIP] {file.FileName} → system table, Import System Tables = No.");
                    AppendConsole($"SKIP {file.FileName} — system table, Import System Tables = No.");
                    continue;
                }

                string targetDb = isCompany ? opts.CompanyDbName : opts.SystemDbName;
                AppendConsole($"Processing {file.FileName} → [{targetDb}] ({(isCompany ? "company" : "system")})");

                // Connect to master first so we can create the target DB if it doesn't exist yet.
                baseBuilder.InitialCatalog = "master";
                using (var masterConn = new SqlConnection(baseBuilder.ConnectionString))
                {
                    masterConn.Open();
                    EnsureDatabase(masterConn, targetDb, opts);
                }

                // Now connect directly to the target database for table/import work.
                baseBuilder.InitialCatalog = targetDb;
                using var conn = new SqlConnection(baseBuilder.ConnectionString);
                conn.Open();

                _dbfFields.TryGetValue(tableKey, out var fieldDefs);
                ImportFile(conn, file.FullPath, fieldDefs, opts, ref totalErrors, maxErrors);
            }
            catch (ImportAbortException)
            {
                AppendLog($"[ABORT] Max errors ({maxErrors}) reached. Import stopped.");
                AppendConsole($"ABORT — max errors ({maxErrors}) reached. Import stopped.");
                return;
            }
            catch (Exception ex)
            {
                AppendLog($"[ERROR] {file.FileName}: {ex.Message}");
                AppendConsole($"ERROR processing {file.FileName}: {ex.Message}");
                totalErrors++;
                if (totalErrors >= maxErrors)
                {
                    AppendLog($"[ABORT] Max errors ({maxErrors}) reached. Import stopped.");
                    AppendConsole($"ABORT — max errors ({maxErrors}) reached. Import stopped.");
                    return;
                }
            }
        }
    }

    private void EnsureDatabase(SqlConnection conn, string dbName, ImportOptions opts)
    {
        // conn is open against master; safe to check and create the target DB.
        bool exists = DatabaseExists(conn, dbName);
        if (exists) return;

        switch (opts.DatabaseExists)
        {
            case ExistsAction.ErrorAndExit:
                throw new InvalidOperationException($"Database '{dbName}' does not exist.");
            case ExistsAction.Ask:
                if (!AskUser($"Database '{dbName}' does not exist. Create it?"))
                    throw new OperationCanceledException("User cancelled.");
                CreateDatabase(conn, dbName);
                break;
            case ExistsAction.Overwrite:
                CreateDatabase(conn, dbName);
                break;
            // Skip: proceed — SQL Server will error if objects are missing, caught upstream
        }
    }

    private void ImportFile(SqlConnection conn, string csvPath,
                             List<DbfFieldDef>? fieldDefs, ImportOptions opts,
                             ref int totalErrors, int maxErrors)
    {
        string tableName = Path.GetFileNameWithoutExtension(csvPath);
        AppendLog($"[INFO] Processing file: {Path.GetFileName(csvPath)} → [{conn.Database}].[{tableName}]");
        AppendConsole($"  Opening {Path.GetFileName(csvPath)} for table [{tableName}] in [{conn.Database}]");

        var csvCfg = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            MissingFieldFound = null,
            BadDataFound = null
        };

        using var reader = new StreamReader(csvPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        using var csv = new CsvReader(reader, csvCfg);

        csv.Read();
        csv.ReadHeader();
        string[] headers = csv.HeaderRecord!;
        AppendConsole($"  Columns ({headers.Length}): {string.Join(", ", headers)}");

        bool tableExists = TableExists(conn, tableName);

        if (tableExists)
        {
            switch (opts.TableExists)
            {
                case ExistsAction.ErrorAndExit:
                    throw new InvalidOperationException($"Table [{tableName}] already exists.");
                case ExistsAction.Skip:
                    AppendLog($"[SKIP] Table [{tableName}] exists — skipping file.");
                    AppendConsole($"  SKIP — table [{tableName}] already exists.");
                    return;
                case ExistsAction.Ask:
                    if (!AskUser($"Table [{tableName}] already exists. Drop and recreate?"))
                    {
                        AppendLog($"[SKIP] Table [{tableName}] — user skipped.");
                        AppendConsole($"  SKIP — user chose not to overwrite [{tableName}].");
                        return;
                    }
                    DropTable(conn, tableName);
                    AppendConsole($"  Dropped existing table [{tableName}]. Re-creating…");
                    CreateTableFromSchema(conn, tableName, headers, fieldDefs, opts.CreateIdField);
                    break;
                case ExistsAction.Overwrite:
                    DropTable(conn, tableName);
                    AppendConsole($"  Dropped existing table [{tableName}]. Re-creating…");
                    CreateTableFromSchema(conn, tableName, headers, fieldDefs, opts.CreateIdField);
                    break;
            }
        }
        else
        {
            AppendConsole($"  Creating table [{tableName}]…");
            CreateTableFromSchema(conn, tableName, headers, fieldDefs, opts.CreateIdField);
            AppendConsole($"  Table [{tableName}] created.");
        }

        const int batchSize = 2000;
        int rowNum = 0;
        int batchStart = 1;

        while (csv.Read())
        {
            rowNum++;
            try
            {
                var values = headers.Select(h => csv.GetField(h)).ToList();

                // Skip entirely blank rows — every column is null or whitespace.
                if (values.All(v => string.IsNullOrWhiteSpace(v)))
                    continue;

                InsertRow(conn, tableName, headers, values!, opts, ref totalErrors, maxErrors);
            }
            catch (ImportAbortException) { throw; }
            catch (Exception ex)
            {
                LogEntry(opts.RecordLog, $"[ERROR] Row {rowNum} in {Path.GetFileName(csvPath)}: {ex.Message}");
                AppendConsole($"  ERROR row {rowNum}: {ex.Message}");
                totalErrors++;
                if (totalErrors >= maxErrors)
                    throw new ImportAbortException();
            }

            // Report progress at each batch boundary
            if (rowNum % batchSize == 0)
            {
                AppendConsole($"  Records {batchStart:N0}–{rowNum:N0} inserted…");
                batchStart = rowNum + 1;
            }
        }

        // Report any remaining rows that didn't fill a full batch
        if (rowNum % batchSize != 0 && rowNum >= batchStart)
            AppendConsole($"  Records {batchStart:N0}–{rowNum:N0} inserted.");

        AppendLog($"[INFO] Finished [{tableName}]: {rowNum} row(s) processed.");
        AppendConsole($"  Done — [{tableName}]: {rowNum:N0} total row(s) imported.");
    }

    private void InsertRow(SqlConnection conn, string tableName, string[] headers,
                           List<string> values, ImportOptions opts, ref int totalErrors, int maxErrors)
    {
        string colList = string.Join(", ", headers.Select(EscapeId));
        string paramList = string.Join(", ", headers.Select((_, i) => $"@p{i}"));

        using var cmd = new SqlCommand($"INSERT INTO {EscapeId(tableName)} ({colList}) VALUES ({paramList})", conn);
        for (int i = 0; i < headers.Length; i++)
        {
            // Treat empty / whitespace CSV values as NULL so numeric/date columns don't throw a conversion error.
            object paramVal = string.IsNullOrWhiteSpace(values[i]) ? DBNull.Value : (object)values[i];
            cmd.Parameters.AddWithValue($"@p{i}", paramVal);
        }

        try
        {
            cmd.ExecuteNonQuery();
            LogEntry(opts.RecordLog == LogLevel.All ? opts.RecordLog : LogLevel.None,
                     $"[ROW] Inserted into [{tableName}]");
        }
        catch (SqlException ex) when (ex.Number == 2627 || ex.Number == 2601) // duplicate key
        {
            switch (opts.RecordExists)
            {
                case ExistsAction.ErrorAndExit:
                    throw;
                case ExistsAction.Skip:
                    LogEntry(opts.RecordLog, $"[SKIP] Duplicate row in [{tableName}]");
                    return;
                case ExistsAction.Ask:
                    if (!AskUser($"Duplicate record in [{tableName}]. Overwrite?"))
                    {
                        LogEntry(opts.RecordLog, $"[SKIP] Duplicate row in [{tableName}] — user skipped.");
                        return;
                    }
                    goto case ExistsAction.Overwrite;
                case ExistsAction.Overwrite:
                {
                    // Build UPDATE ... WHERE (all cols) — simple approach
                    var setClauses = headers.Select((h, i) => $"{EscapeId(h)} = @p{i}");
                    var whereClauses = headers.Select((h, i) => $"{EscapeId(h)} = @p{i}");
                    using var upd = new SqlCommand(
                        $"UPDATE {EscapeId(tableName)} SET {string.Join(", ", setClauses)} WHERE {string.Join(" AND ", whereClauses)}",
                        conn);
                    for (int i = 0; i < headers.Length; i++)
                    {
                        object paramVal = string.IsNullOrWhiteSpace(values[i]) ? DBNull.Value : (object)values[i];
                        upd.Parameters.AddWithValue($"@p{i}", paramVal);
                    }
                    upd.ExecuteNonQuery();
                    LogEntry(opts.RecordLog, $"[UPDATE] Overwrote duplicate in [{tableName}]");
                    break;
                }
            }
        }
    }

    // ── SQL Helpers ────────────────────────────────────────────────────────
    private static bool DatabaseExists(SqlConnection conn, string dbName)
    {
        using var cmd = new SqlCommand("SELECT COUNT(1) FROM sys.databases WHERE name = @n", conn);
        cmd.Parameters.AddWithValue("@n", dbName);
        return (int)cmd.ExecuteScalar()! > 0;
    }

    private static void CreateDatabase(SqlConnection conn, string dbName)
    {
        using var cmd = new SqlCommand($"CREATE DATABASE {EscapeId(dbName)}", conn);
        cmd.ExecuteNonQuery();
    }

    private static bool TableExists(SqlConnection conn, string tableName)
    {
        using var cmd = new SqlCommand(
            "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @n AND TABLE_TYPE = 'BASE TABLE'", conn);
        cmd.Parameters.AddWithValue("@n", tableName);
        return (int)cmd.ExecuteScalar()! > 0;
    }

    /// <summary>
    /// Creates the table using typed column definitions from DBFCHK.CSV when available;
    /// falls back to NVARCHAR(MAX) for any column not in the schema.
    /// All columns are nullable. Optionally prepends a NewID UNIQUEIDENTIFIER column.
    /// </summary>
    private static void CreateTableFromSchema(SqlConnection conn, string tableName,
                                               string[] headers, List<DbfFieldDef>? fieldDefs,
                                               bool createIdField)
    {
        var lookup = fieldDefs?
            .ToDictionary(f => f.FieldName, f => f, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, DbfFieldDef>(StringComparer.OrdinalIgnoreCase);

        var colDefs = new List<string>();

        if (createIdField)
            colDefs.Add("[NewID] UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID()");

        colDefs.AddRange(headers.Select(h =>
        {
            string sqlType = lookup.TryGetValue(h, out var def)
                ? MapSqlType(def)
                : "NVARCHAR(MAX)";
            return $"{EscapeId(h)} {sqlType} NULL";
        }));

        using var cmd = new SqlCommand(
            $"CREATE TABLE {EscapeId(tableName)} ({string.Join(", ", colDefs)})", conn);
        cmd.ExecuteNonQuery();
    }

    /// <summary>Maps a DBFCHK field definition to a SQL Server data type string.</summary>
    private static string MapSqlType(DbfFieldDef def)
    {
        if (def.Type == "C")
        {
            int len = def.Length > 0 ? def.Length : 255;
            return $"NVARCHAR({len})";
        }

        if (def.Type == "N")
        {
            if (IsMoney(def.FieldName))
                return "MONEY";

            if (def.Decimals > 0)
            {
                int precision = def.Length > 0 ? def.Length : 18;
                return $"DECIMAL({precision},{def.Decimals})";
            }

            // Integer — choose smallest fitting type
            return def.Length switch
            {
                <= 4  => "SMALLINT",
                <= 9  => "INT",
                <= 18 => "BIGINT",
                _     => "DECIMAL(38,0)"
            };
        }

        // Unknown type — safe fallback
        return "NVARCHAR(MAX)";
    }

    private static readonly string[] _moneyKeywords =
        ["AMT", "AMOUNT", "PRICE", "COST", "BALANCE", "TOTAL"];

    /// <summary>Returns true if the field name suggests a monetary value.</summary>
    private static bool IsMoney(string fieldName)
    {
        string upper = fieldName.ToUpperInvariant();
        return _moneyKeywords.Any(k => upper.Contains(k));
    }

    private static void DropTable(SqlConnection conn, string tableName)
    {
        using var cmd = new SqlCommand($"DROP TABLE {EscapeId(tableName)}", conn);
        cmd.ExecuteNonQuery();
    }

    private static string EscapeId(string name) => $"[{name.Replace("]", "]]")}]";

    // ── .ini Config Save ───────────────────────────────────────────────────
    /// <summary>
    /// Creates/overwrites PcaWorks26.ini in the chosen folder with database connection details.
    /// </summary>
    private void SaveIniFile(string folder, string connectionString, string systemDbName, string companyDbName)
    {
        Directory.CreateDirectory(folder);
        string iniPath = Path.Combine(folder, "PcaWorks26.ini");

        var sb = new StringBuilder();
        sb.AppendLine("[Database]");
        sb.AppendLine($"ConnectionString={connectionString}");
        sb.AppendLine($"SystemDatabase={systemDbName}");
        sb.AppendLine($"CompanyDatabase={companyDbName}");
        sb.AppendLine();
        sb.AppendLine($"; Generated by Kw2Pw26 on {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

        File.WriteAllText(iniPath, sb.ToString(), Encoding.UTF8);
        AppendLog($"[INFO] Saved config to {iniPath}");
        AppendConsole($"Saved PcaWorks26.ini → {iniPath}");
        SetStatus($"Config saved → {iniPath}");
    }

    // ── Logging ────────────────────────────────────────────────────────────

    /// <summary>
    /// Resolves (and lazily creates) the session log path: &lt;exe folder&gt;\yyyyMMdd_HHmmss_Kw2Pw26.log.
    /// The file itself is not created until the first write.
    /// </summary>
    private string EnsureLogPath()
    {
        if (_logPath is null)
        {
            string exeDir = AppContext.BaseDirectory;
            string stamp  = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            _logPath = Path.Combine(exeDir, $"{stamp}_Kw2Pw26.log");
        }
        return _logPath;
    }

    /// <summary>Appends a timestamped line to the session log, creating the file on first write.</summary>
    private void AppendLog(string message)
    {
        string path = EnsureLogPath();
        string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}";
        File.AppendAllText(path, line + Environment.NewLine);

        // Enable VIEW LOG the moment the file exists
        Dispatcher.InvokeAsync(() => BtnViewLog.IsEnabled = true);
    }

    /// <summary>
    /// Writes to the log only when the entry's severity matches the configured level.
    /// Errors are always written regardless of level.
    /// </summary>
    private void LogEntry(LogLevel level, string message)
    {
        bool isError = message.StartsWith("[ERROR", StringComparison.Ordinal);

        if (level == LogLevel.None && !isError) return;
        if (level == LogLevel.Errors && !isError) return;
        // LogLevel.All: write everything

        AppendLog(message);
    }

    // ── Helpers ────────────────────────────────────────────────────────────
    private void SetStatus(string msg) =>
        Dispatcher.Invoke(() => TxtStatus.Text = msg);

    /// <summary>Appends a timestamped line to the green activity console and auto-scrolls.</summary>
    private void AppendConsole(string message)
    {
        string line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        Dispatcher.InvokeAsync(() =>
        {
            TxtConsole.AppendText(line + "\n");
            // Defer scroll until after WPF has completed its layout pass for the new text.
            TxtConsole.Dispatcher.InvokeAsync(
                () => ConsoleScroller.ScrollToBottom(),
                System.Windows.Threading.DispatcherPriority.Background);
        });
    }

    private bool AskUser(string question)
    {
        bool result = false;
        Dispatcher.Invoke(() =>
        {
            result = MessageBox.Show(question, "Import — Action Required",
                                     MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
        });
        return result;
    }

    private int ParseMaxErrors()
    {
        if (int.TryParse(TxtMaxErrors.Text.Trim(), out int v) && v > 0) return v;
        return 10;
    }

    private ImportOptions BuildOptions() => new(
        DatabaseExists:       ComboToAction(CmbDatabaseExists),
        DatabaseLog:          ComboToLog(CmbDatabaseLog),
        TableExists:          ComboToAction(CmbTableExists),
        TableLog:             ComboToLog(CmbTableLog),
        RecordExists:         ComboToAction(CmbRecordExists),
        RecordLog:            ComboToLog(CmbRecordLog),
        ImportSystemTables:   CmbImportSystemTables.SelectedIndex == 0,
        ImportCompanyTables:  CmbImportCompanyTables.SelectedIndex == 0,
        SystemDbName:         TxtSystemDbName.Text.Trim(),
        CompanyDbName:        TxtCompanyDbName.Text.Trim(),
        CreateIdField:        ChkCreateIdField.IsChecked == true);

    private static ExistsAction ComboToAction(ComboBox cmb) => cmb.SelectedIndex switch
    {
        1 => ExistsAction.Overwrite,
        2 => ExistsAction.Skip,
        3 => ExistsAction.Ask,
        _ => ExistsAction.ErrorAndExit
    };

    private static LogLevel ComboToLog(ComboBox cmb) => cmb.SelectedIndex switch
    {
        1 => LogLevel.All,
        2 => LogLevel.None,
        _ => LogLevel.Errors
    };
}

// ── Supporting types ───────────────────────────────────────────────────────

public class CsvFileItem(string fileName, string fullPath) : System.ComponentModel.INotifyPropertyChanged
{
    private bool _isSelected = true;

    public string FileName { get; } = fileName;
    public string FullPath { get; } = fullPath;

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; PropertyChanged?.Invoke(this, new(nameof(IsSelected))); }
    }

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
}

public enum ExistsAction { ErrorAndExit, Overwrite, Skip, Ask }
public enum LogLevel { Errors, All, None }

public record ImportOptions(
    ExistsAction DatabaseExists, LogLevel DatabaseLog,
    ExistsAction TableExists,    LogLevel TableLog,
    ExistsAction RecordExists,   LogLevel RecordLog,
    bool ImportSystemTables,
    bool ImportCompanyTables,
    string SystemDbName,
    string CompanyDbName,
    bool CreateIdField);

public class ImportAbortException : Exception { }

/// <summary>One row from DBFCHK.CSV — describes a single column in a CSV/DBF file.</summary>
public record DbfFieldDef(string FieldName, string Type, int Length, int Decimals, string Dir);

/// <summary>One row from DBF.CSV — routes a CSV file to system or company database.</summary>
public record DbfTableDef(string DbfName, bool IsCompany);
