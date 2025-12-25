using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System.Windows;

namespace WpfApp1
{
    public partial class MainWindow : System.Windows.Window
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";
        private const string DbPath = @"C:\sqLite\data_qc.db";

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog;

        public MainWindow()
        {
            InitializeComponent();
            ImportExcelToSQLite();
        }

        private void ImportExcelToSQLite()
        {
            try
            {
                EnsureDatabaseFolderExists();

                using var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                CreateBusbarTable(connection);

                using var transaction = connection.BeginTransaction();

                TraverseFoldersAndImport(connection, transaction);

                transaction.Commit();

                ShowFinalReport();
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"ERROR FATAL:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Import Gagal",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        private void EnsureDatabaseFolderExists()
        {
            string folder = System.IO.Path.GetDirectoryName(DbPath);

            if (!string.IsNullOrEmpty(folder) && !System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }
        }

        private void CreateBusbarTable(Microsoft.Data.Sqlite.SqliteConnection connection)
        {
            using var cmd = connection.CreateCommand();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Busbar;

                CREATE TABLE IF NOT EXISTS Busbar (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT,
                    Thickness_mm REAL,
                    Width_mm REAL,
                    Length INTEGER,
                    Radius REAL,
                    Chamber_mm REAL,
                    Electric_IACS REAL,
                    Weight REAL,
                    Elongation REAL,
                    Tensile REAL,
                    Bend_test TEXT,
                    Spectro_Cu REAL,
                    Oxygen REAL
                );
            ";

            cmd.ExecuteNonQuery();
        }

        private void TraverseFoldersAndImport(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            ResetCounters();

            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

            foreach (string yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                string year = new System.IO.DirectoryInfo(yearDir).Name.Trim();

                foreach (string monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    string rawMonth = new System.IO.DirectoryInfo(monthDir).Name.Trim();
                    string normalizedMonth = NormalizeMonthFolder(rawMonth);

                    foreach (string file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);

                        if (fileName.StartsWith("~$"))
                            continue;

                        _totalFilesFound++;

                        ProcessSingleExcelFile(connection, transaction, file, year, normalizedMonth);
                    }
                }
            }
        }

        private void ProcessSingleExcelFile(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string filePath,
            string year,
            string month)
        {
            try
            {
                using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);

                var sheet = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w.Name.Trim().Equals(
                            "YLB 50",
                            System.StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    AppendDebug($"SKIP: Sheet 'YLB 50' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                    return;
                }

                int row = 3;

                while (true)
                {
                    string sizeValue = sheet.Cell(row, "C").GetString();

                    if (string.IsNullOrWhiteSpace(sizeValue))
                        break;

                    string cleanedSize = CleanSizeText(sizeValue);


                    // 1. Kelompok 2 Angka Belakang Koma
                    double rawThickness = ParseCustomDecimal(sheet.Cell(row, "G").GetString());
                    double valThickness = System.Math.Round(rawThickness, 2);

                    double rawWidth = ParseCustomDecimal(sheet.Cell(row, "I").GetString());
                    double valWidth = System.Math.Round(rawWidth, 2);

                    double rawRadius = ParseCustomDecimal(sheet.Cell(row, "J").GetString());
                    double valRadius = System.Math.Round(rawRadius, 2);

                    double rawChamber = ParseCustomDecimal(sheet.Cell(row, "L").GetString());
                    double valChamber = System.Math.Round(rawChamber, 2);

                    double rawElectric = ParseCustomDecimal(sheet.Cell(row, "U").GetString());
                    double valElectric = System.Math.Round(rawElectric, 2);

                    double rawOxygen = ParseCustomDecimal(sheet.Cell(row, "X").GetString());
                    double valOxygen = System.Math.Round(rawOxygen, 2);

                    // 2. Kelompok Tanpa Batasan (Asli)
                    double valSpectro = ParseCustomDecimal(sheet.Cell(row, "Y").GetString());
                    double valResistivity = ParseCustomDecimal(sheet.Cell(row, "T").GetString());

                    // 3. Kelompok Bilangan Bulat (0 Koma)
                    double rawLength = ParseCustomDecimal(sheet.Cell(row, "K").GetString());
                    double valLength = System.Math.Round(rawLength, 0);

                    // 4. Kelompok Merge/Rata-rata (Tetap dibulatkan 2 angka)
                    double rawElongation = GetMergedOrAverageValue(sheet, row, "R");
                    double valElongation = System.Math.Round(rawElongation, 2);

                    double rawTensile = GetMergedOrAverageValue(sheet, row, "Q");
                    double valTensile = System.Math.Round(rawTensile, 2);

                    string valBendTest = sheet.Cell(row, "W").GetString();

                    InsertBusbarRow(
                        connection, transaction,
                        cleanedSize, year, month,
                        valThickness, valWidth, valLength, valRadius, valChamber,
                        valElectric, valResistivity, valElongation, valTensile,
                        valBendTest, valSpectro, valOxygen
                    );

                    _totalRowsInserted++;

                    row += 2;
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE: {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private double GetMergedOrAverageValue(ClosedXML.Excel.IXLWorksheet sheet, int startRow, string columnLetter)
        {
            var cellFirst = sheet.Cell(startRow, columnLetter);

            if (cellFirst.IsMerged())
            {
                return ParseCustomDecimal(cellFirst.GetString());
            }
            else
            {
                var val1 = ParseCustomDecimal(cellFirst.GetString());
                var val2 = ParseCustomDecimal(sheet.Cell(startRow + 1, columnLetter).GetString());

                if (val1 == 0) return val2;
                if (val2 == 0) return val1;

                return (val1 + val2) / 2.0;
            }
        }

        private double ParseCustomDecimal(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
                return 0.0;

            string cleanInput = rawInput.Replace(",", ".").Trim();

            if (double.TryParse(
                cleanInput,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double result))
            {
                return result;
            }

            return 0.0;
        }

        private string NormalizeMonthFolder(string rawMonth)
        {
            if (string.IsNullOrWhiteSpace(rawMonth))
                return string.Empty;

            for (int i = 0; i < rawMonth.Length; i++)
            {
                if (!char.IsDigit(rawMonth[i]))
                    continue;

                int start = i;

                while (i < rawMonth.Length && char.IsDigit(rawMonth[i]))
                    i++;

                string numberText = rawMonth.Substring(start, i - start);

                if (!int.TryParse(numberText, out int monthNumber))
                    continue;

                if (monthNumber < 1 || monthNumber > 12)
                    continue;

                return new System.Globalization.DateTimeFormatInfo()
                    .GetMonthName(monthNumber);
            }

            return string.Empty;
        }

        private string CleanSizeText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            string text = raw.ToUpper();

            for (int i = 0; i < text.Length; i++)
            {
                if (!char.IsDigit(text[i]))
                    continue;

                int start = i;

                while (i < text.Length && char.IsDigit(text[i]))
                    i++;

                if (i >= text.Length || text[i] != 'X')
                    continue;

                i++;

                if (i >= text.Length || !char.IsDigit(text[i]))
                    continue;

                while (i < text.Length && char.IsDigit(text[i]))
                    i++;

                string result = text.Substring(start, i - start);
                return result.Trim();
            }

            return string.Empty;
        }

        private void InsertBusbarRow(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string size, string year, string month,
            double thickness, double width, double length, double radius, double chamber,
            double electric, double resistivity, double elongation, double tensile,
            string bendTest, double spectro, double oxygen)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                INSERT INTO Busbar (
                    Size_mm, Year_folder, Month_folder, 
                    Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                    Electric_IACS, Weight, Elongation, Tensile,
                    Bend_test, Spectro_Cu, Oxygen
                )
                VALUES (
                    @Size, @Year, @Month, 
                    @Thickness, @Width, @Length, @Radius, @Chamber,
                    @Electric, @Resistivity, @Elongation, @Tensile,
                    @Bend, @Spectro, @Oxygen
                );
            ";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year.Trim());
            cmd.Parameters.AddWithValue("@Month", month.Trim());

            cmd.Parameters.AddWithValue("@Thickness", thickness);
            cmd.Parameters.AddWithValue("@Width", width);
            cmd.Parameters.AddWithValue("@Length", length);
            cmd.Parameters.AddWithValue("@Radius", radius);
            cmd.Parameters.AddWithValue("@Chamber", chamber);

            cmd.Parameters.AddWithValue("@Electric", electric);
            cmd.Parameters.AddWithValue("@Resistivity", resistivity);
            cmd.Parameters.AddWithValue("@Elongation", elongation);
            cmd.Parameters.AddWithValue("@Tensile", tensile);

            object bendVal = string.IsNullOrEmpty(bendTest) ? System.DBNull.Value : bendTest;
            cmd.Parameters.AddWithValue("@Bend", bendVal);

            cmd.Parameters.AddWithValue("@Spectro", spectro);
            cmd.Parameters.AddWithValue("@Oxygen", oxygen);

            cmd.ExecuteNonQuery();
        }

        private void ResetCounters()
        {
            _totalFilesFound = 0;
            _totalRowsInserted = 0;
            _debugLog = "";
        }

        private void AppendDebug(string message)
        {
            if (_debugLog.Length < 1000)
            {
                _debugLog += message + System.Environment.NewLine;
            }
        }

        private void ShowFinalReport()
        {
            System.Windows.MessageBox.Show(
                $"IMPORT SELESAI\n\n" +
                $"File ditemukan : {_totalFilesFound}\n" +
                $"Baris disimpan : {_totalRowsInserted}\n\n" +
                $"Debug Log:\n{_debugLog}",
                "Laporan Import",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }
    }
}