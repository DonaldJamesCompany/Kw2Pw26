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
    private string _logPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Kw2Pw26", "import.log");

    public MainWindow()
    {
        InitializeComponent();
        FileList.ItemsSource = _files;
        Directory.CreateDirectory(Path.GetDirectoryName(_logPath)!);
    }

    // ── Connection String TEST ─────────────────────────────────────────────
    private async void BtnTest_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Testing connection…");
        try
        {
            await using var conn = new SqlConnection(TxtConnectionString.Text.Trim());
            await conn.OpenAsync();
            SetStatus("✔ Connection successful.");
        }
        catch (Exception ex)
        {
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
        SetStatus(_files.Count == 0 ? "No CSV files found in selected folder." : $"{_files.Count} CSV file(s) found.");
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
        SetStatus("Form cleared.");
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

        BtnImport.IsEnabled = false;
        SetStatus("Importing…");

        try
        {
            await Task.Run(() => RunImport(connStr, selected, options, maxErrors));
            SetStatus("Import complete. See log for details.");
        }
        catch (Exception ex)
        {
            SetStatus($"Import aborted: {ex.Message}");
        }
        finally
        {
            BtnImport.IsEnabled = true;
        }
    }

    private void BtnViewLog_Click(object sender, RoutedEventArgs e)
    {
        if (!File.Exists(_logPath))
        {
            SetStatus("Log file does not exist yet.");
            return;
        }
        System.Diagnostics.Process.Start("notepad.exe", _logPath);
    }

    private void BtnExit_Click(object sender, RoutedEventArgs e) => Close();

    // ── Import Logic ───────────────────────────────────────────────────────
    private void RunImport(string connStr, List<CsvFileItem> files, ImportOptions opts, int maxErrors)
    {
        int totalErrors = 0;

        using var conn = new SqlConnection(connStr);
        conn.Open();

        // Check / create database
        string dbName = conn.Database;
        if (!DatabaseExists(conn, dbName))
        {
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
                // Skip: proceed anyway (connection already points to the db)
            }
        }

        foreach (var file in files)
        {
            try
            {
                ImportFile(conn, file.FullPath, opts, ref totalErrors, maxErrors);
            }
            catch (ImportAbortException)
            {
                AppendLog($"[ABORT] Max errors ({maxErrors}) reached. Import stopped.");
                return;
            }
            catch (Exception ex)
            {
                AppendLog($"[ERROR] {file.FileName}: {ex.Message}");
                totalErrors++;
                if (totalErrors >= maxErrors)
                {
                    AppendLog($"[ABORT] Max errors ({maxErrors}) reached. Import stopped.");
                    return;
                }
            }
        }
    }

    private void ImportFile(SqlConnection conn, string csvPath, ImportOptions opts,
                             ref int totalErrors, int maxErrors)
    {
        string tableName = Path.GetFileNameWithoutExtension(csvPath);
        AppendLog($"[INFO] Processing file: {Path.GetFileName(csvPath)} → table [{tableName}]");

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            MissingFieldFound = null,
            BadDataFound = null
        };

        using var reader = new StreamReader(csvPath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        using var csv = new CsvReader(reader, config);

        csv.Read();
        csv.ReadHeader();
        string[] headers = csv.HeaderRecord!;

        bool tableExists = TableExists(conn, tableName);

        if (tableExists)
        {
            switch (opts.TableExists)
            {
                case ExistsAction.ErrorAndExit:
                    throw new InvalidOperationException($"Table [{tableName}] already exists.");
                case ExistsAction.Skip:
                    AppendLog($"[SKIP] Table [{tableName}] exists — skipping file.");
                    return;
                case ExistsAction.Ask:
                    if (!AskUser($"Table [{tableName}] already exists. Drop and recreate?"))
                    {
                        AppendLog($"[SKIP] Table [{tableName}] — user skipped.");
                        return;
                    }
                    DropTable(conn, tableName);
                    CreateTable(conn, tableName, headers);
                    break;
                case ExistsAction.Overwrite:
                    DropTable(conn, tableName);
                    CreateTable(conn, tableName, headers);
                    break;
            }
        }
        else
        {
            CreateTable(conn, tableName, headers);
        }

        int rowNum = 0;
        while (csv.Read())
        {
            rowNum++;
            try
            {
                var values = headers.Select(h => csv.GetField(h)).ToList();
                InsertRow(conn, tableName, headers, values!, opts, ref totalErrors, maxErrors);
            }
            catch (ImportAbortException) { throw; }
            catch (Exception ex)
            {
                LogEntry(opts.RecordLog, $"[ERROR] Row {rowNum} in {Path.GetFileName(csvPath)}: {ex.Message}");
                totalErrors++;
                if (totalErrors >= maxErrors)
                    throw new ImportAbortException();
            }
        }

        AppendLog($"[INFO] Finished [{tableName}]: {rowNum} row(s) processed.");
    }

    private void InsertRow(SqlConnection conn, string tableName, string[] headers,
                           List<string> values, ImportOptions opts, ref int totalErrors, int maxErrors)
    {
        string colList = string.Join(", ", headers.Select(EscapeId));
        string paramList = string.Join(", ", headers.Select((_, i) => $"@p{i}"));

        using var cmd = new SqlCommand($"INSERT INTO {EscapeId(tableName)} ({colList}) VALUES ({paramList})", conn);
        for (int i = 0; i < headers.Length; i++)
            cmd.Parameters.AddWithValue($"@p{i}", (object?)values[i] ?? DBNull.Value);

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
                        upd.Parameters.AddWithValue($"@p{i}", (object?)values[i] ?? DBNull.Value);
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

    private static void CreateTable(SqlConnection conn, string tableName, string[] headers)
    {
        string cols = string.Join(", ", headers.Select(h => $"{EscapeId(h)} NVARCHAR(MAX)"));
        using var cmd = new SqlCommand($"CREATE TABLE {EscapeId(tableName)} ({cols})", conn);
        cmd.ExecuteNonQuery();
    }

    private static void DropTable(SqlConnection conn, string tableName)
    {
        using var cmd = new SqlCommand($"DROP TABLE {EscapeId(tableName)}", conn);
        cmd.ExecuteNonQuery();
    }

    private static string EscapeId(string name) => $"[{name.Replace("]", "]]")}]";

    // ── Logging ────────────────────────────────────────────────────────────
    private void AppendLog(string message)
    {
        string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  {message}";
        File.AppendAllText(_logPath, line + Environment.NewLine);
    }

    private void LogEntry(LogLevel level, string message)
    {
        if (level == LogLevel.None) return;
        if (level == LogLevel.Errors && !message.StartsWith("[ERROR")) return;
        AppendLog(message);
    }

    // ── Helpers ────────────────────────────────────────────────────────────
    private void SetStatus(string msg) =>
        Dispatcher.Invoke(() => TxtStatus.Text = msg);

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
        DatabaseExists: ComboToAction(CmbDatabaseExists),
        DatabaseLog:    ComboToLog(CmbDatabaseLog),
        TableExists:    ComboToAction(CmbTableExists),
        TableLog:       ComboToLog(CmbTableLog),
        RecordExists:   ComboToAction(CmbRecordExists),
        RecordLog:      ComboToLog(CmbRecordLog));

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
    ExistsAction RecordExists,   LogLevel RecordLog);

public class ImportAbortException : Exception { }
