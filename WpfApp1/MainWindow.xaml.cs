using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using System.Linq;

namespace WpfApp1
{
    public class TljRecord
    {
        public System.DateTime Date { get; set; }
        public string BatchNo { get; set; }
        public string Size { get; set; }
        // Tambahan info untuk debugging agar tahu data dari file mana
        public string SourceFile { get; set; }
    }

    public partial class MainWindow : System.Windows.Window
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";
        private const string DbPath = @"C:\sqLite\data_qc.db";

        private int _totalFilesFound;
        private int _totalRowsInserted;
        private string _debugLog;

        // Variabel Global untuk menampung SEMUA data TLJ dari seluruh folder
        private System.Collections.Generic.List<TljRecord> _globalTljRecords;

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
                ResetCounters();

                // FASE 1: BACA SEMUA DATA TLJ DARI SEMUA FILE (GLOBAL HARVESTING)
                _globalTljRecords = new System.Collections.Generic.List<TljRecord>();
                HarvestAllTljData();

                if (_globalTljRecords.Count == 0)
                {
                    System.Windows.MessageBox.Show("Tidak ada data TLJ ditemukan di seluruh folder.", "Peringatan");
                }

                // FASE 2: PROSES DATA YLB DAN COCOKKAN DENGAN GLOBAL TLJ
                using var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={DbPath}");
                connection.Open();

                CreateBusbarTable(connection);

                using var transaction = connection.BeginTransaction();

                TraverseFoldersAndProcessYlb(connection, transaction);

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

        // --- FASE 1: MENGUMPULKAN DATA TLJ ---
        private void HarvestAllTljData()
        {
            if (!System.IO.Directory.Exists(ExcelRootFolder)) return;

            foreach (string yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                foreach (string monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    foreach (string file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);
                        if (fileName.StartsWith("~$")) continue;

                        try
                        {
                            // Buka Workbook hanya untuk mengambil TLJ
                            using var workbook = new ClosedXML.Excel.XLWorkbook(file);

                            // Ambil dari Sheet 350
                            var tlj350 = LoadTljSheet(workbook, "TLJ 350", fileName);
                            _globalTljRecords.AddRange(tlj350);

                            // Ambil dari Sheet 500
                            var tlj500 = LoadTljSheet(workbook, "TLJ 500", fileName);
                            _globalTljRecords.AddRange(tlj500);
                        }
                        catch (System.Exception ex)
                        {
                            AppendDebug($"Warning Reading TLJ {fileName}: {ex.Message}");
                        }
                    }
                }
            }
        }

        private System.Collections.Generic.List<TljRecord> LoadTljSheet(ClosedXML.Excel.XLWorkbook workbook, string sheetName, string sourceFileName)
        {
            var results = new System.Collections.Generic.List<TljRecord>();
            var sheet = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals(sheetName, System.StringComparison.OrdinalIgnoreCase));

            if (sheet == null) return results;

            int r = 3;
            while (true)
            {
                string rawSize = sheet.Cell(r, "D").GetString();
                if (string.IsNullOrWhiteSpace(rawSize)) break;

                string rawBatch = sheet.Cell(r, "C").GetString();
                string rawDate = sheet.Cell(r, "B").GetString();

                if (System.DateTime.TryParse(rawDate, out System.DateTime dt))
                {
                    // Filter: Hanya masukkan data yang valid (Batch terisi)
                    if (!string.IsNullOrWhiteSpace(rawBatch))
                    {
                        results.Add(new TljRecord
                        {
                            Date = dt,
                            BatchNo = rawBatch.Trim(),
                            Size = CleanSizeText(rawSize),
                            SourceFile = sourceFileName
                        });
                    }
                }
                r++;
            }
            return results;
        }

        // --- FASE 2: MEMPROSES YLB ---
        private void TraverseFoldersAndProcessYlb(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            if (!System.IO.Directory.Exists(ExcelRootFolder)) return;

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
                        if (fileName.StartsWith("~$")) continue;

                        _totalFilesFound++;
                        ProcessSingleExcelFileForYlb(connection, transaction, file, year, normalizedMonth);
                    }
                }
            }
        }

        private void ProcessSingleExcelFileForYlb(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string filePath,
            string year,
            string month)
        {
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);

            int row = 3;
            try
            {
                var sheet_YLB = workbook.Worksheets
                    .FirstOrDefault(w => w.Name.Trim().Equals("YLB 50", System.StringComparison.OrdinalIgnoreCase));

                if (sheet_YLB == null) return;

                string currentProdDateString = string.Empty;
                System.DateTime? currentProdDateObj = null;

                int folderMonthNum = GetMonthNumber(month);
                int folderYearNum = 0;
                int.TryParse(year, out folderYearNum);

                while (true)
                {
                    string sizeValue_YLB = sheet_YLB.Cell(row, "C").GetString();
                    if (string.IsNullOrWhiteSpace(sizeValue_YLB)) break;

                    string rawDateFromCell = sheet_YLB.Cell(row, "B").GetString().Trim();
                    if (!string.IsNullOrEmpty(rawDateFromCell))
                    {
                        currentProdDateString = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        if (System.DateTime.TryParseExact(currentProdDateString, "dd/MM/yyyy",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None,
                            out System.DateTime dtResult))
                        {
                            currentProdDateObj = dtResult;
                        }
                    }

                    string cleanSize_YLB = CleanSizeText(sizeValue_YLB);
                    string foundBatchNo = "0.0";

                    if (currentProdDateObj.HasValue && !string.IsNullOrEmpty(cleanSize_YLB))
                    {
                        // Mencari di GLOBAL LIST (_globalTljRecords)
                        // Logika: Cari Size Sama, Tanggal TLJ <= Tanggal YLB, ambil tanggal paling besar

                        var candidates = _globalTljRecords
                            .Where(x => x.Size == cleanSize_YLB && x.Date <= currentProdDateObj.Value)
                            .OrderByDescending(x => x.Date)
                            .ToList();

                        if (candidates.Count > 0)
                        {
                            foundBatchNo = candidates.First().BatchNo;
                        }
                    }

                    // --- Mengambil Data Lainnya ---
                    double valThickness = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "G").GetString()), 2);
                    double valWidth = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "I").GetString()), 2);
                    double valRadius = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "J").GetString()), 2);
                    double valChamber = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "L").GetString()), 2);
                    double valElectric = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "U").GetString()), 2);
                    double valOxygen = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "X").GetString()), 2);
                    double valSpectro = ParseCustomDecimal(sheet_YLB.Cell(row, "Y").GetString());
                    double valResistivity = ParseCustomDecimal(sheet_YLB.Cell(row, "T").GetString());
                    double valLength = System.Math.Round(ParseCustomDecimal(sheet_YLB.Cell(row, "K").GetString()), 0);

                    double valElongation = System.Math.Round(GetMergedOrAverageValue(sheet_YLB, row, "R"), 2);
                    double valTensile = System.Math.Round(GetMergedOrAverageValue(sheet_YLB, row, "Q"), 2);
                    string valBendTest = sheet_YLB.Cell(row, "W").GetString();

                    InsertBusbarRow(
                        connection, transaction,
                        cleanSize_YLB, year, month, foundBatchNo, currentProdDateString,
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
                AppendDebug($"ERROR PROCESS YLB: {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        // --- Helper Methods ---

        private string StandardizeDate(string rawDate, int expectedMonth, int expectedYear)
        {
            if (string.IsNullOrWhiteSpace(rawDate)) return string.Empty;
            if (System.DateTime.TryParse(rawDate, out System.DateTime parsedDate))
            {
                if (expectedYear > 2000 && parsedDate.Year != expectedYear)
                    parsedDate = new System.DateTime(expectedYear, parsedDate.Month, parsedDate.Day);

                if (expectedMonth > 0 && parsedDate.Month != expectedMonth)
                {
                    if (parsedDate.Day <= 12)
                    {
                        int newMonth = parsedDate.Day;
                        int newDay = parsedDate.Month;
                        if (newMonth == expectedMonth)
                            parsedDate = new System.DateTime(parsedDate.Year, newMonth, newDay);
                    }
                }
                return parsedDate.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            }
            return rawDate;
        }

        private int GetMonthNumber(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName)) return 0;
            try { return System.DateTime.ParseExact(monthName, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month; }
            catch { return 0; }
        }

        private double GetMergedOrAverageValue(ClosedXML.Excel.IXLWorksheet sheet_YLB, int startRow, string columnLetter)
        {
            var cellFirst = sheet_YLB.Cell(startRow, columnLetter);
            if (cellFirst.IsMerged()) return ParseCustomDecimal(cellFirst.GetString());
            var val1 = ParseCustomDecimal(cellFirst.GetString());
            var val2 = ParseCustomDecimal(sheet_YLB.Cell(startRow + 1, columnLetter).GetString());
            if (val1 == 0) return val2;
            if (val2 == 0) return val1;
            return (val1 + val2) / 2.0;
        }

        private double ParseCustomDecimal(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return 0.0;
            string cleanInput = rawInput.Replace(",", ".").Trim();
            if (double.TryParse(cleanInput, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }

        private string NormalizeMonthFolder(string rawMonth)
        {
            if (string.IsNullOrWhiteSpace(rawMonth)) return string.Empty;
            for (int i = 0; i < rawMonth.Length; i++)
            {
                if (!char.IsDigit(rawMonth[i])) continue;
                int start = i;
                while (i < rawMonth.Length && char.IsDigit(rawMonth[i])) i++;
                string numberText = rawMonth.Substring(start, i - start);
                if (!int.TryParse(numberText, out int monthNumber)) continue;
                if (monthNumber < 1 || monthNumber > 12) continue;
                return new System.Globalization.DateTimeFormatInfo().GetMonthName(monthNumber);
            }
            return string.Empty;
        }

        private string CleanSizeText(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            string text = raw.ToUpper();
            int start = -1;
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsDigit(text[i])) { start = i; break; }
            }
            if (start == -1) return string.Empty;
            string substring = text.Substring(start);
            int xIndex = substring.IndexOf('X');
            if (xIndex == -1) return string.Empty;
            if (xIndex + 1 >= substring.Length || !char.IsDigit(substring[xIndex + 1])) return string.Empty;
            int end = xIndex + 1;
            while (end < substring.Length && char.IsDigit(substring[end])) end++;
            return substring.Substring(0, end).Trim();
        }

        private void InsertBusbarRow(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string size, string year, string month, string batchNo, string prodDate,
            double thickness, double width, double length, double radius, double chamber,
            double electric, double resistivity, double elongation, double tensile,
            string bendTest, double spectro, double oxygen)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = @"
                INSERT INTO Busbar (
                    Size_mm, Year_folder, Month_folder, Batch_no, Prod_date, 
                    Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                    Electric_IACS, Weight, Elongation, Tensile,
                    Bend_test, Spectro_Cu, Oxygen
                )
                VALUES (
                    @Size, @Year, @Month, @BatchNo, @ProdDate,
                    @Thickness, @Width, @Length, @Radius, @Chamber,
                    @Electric, @Resistivity, @Elongation, @Tensile,
                    @Bend, @Spectro, @Oxygen
                );";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year.Trim());
            cmd.Parameters.AddWithValue("@Month", month.Trim());
            cmd.Parameters.AddWithValue("@BatchNo", batchNo);
            cmd.Parameters.AddWithValue("@ProdDate", prodDate);
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
            if (_debugLog.Length < 1000) _debugLog += message + System.Environment.NewLine;
        }

        private void ShowFinalReport()
        {
            System.Windows.MessageBox.Show(
                $"IMPORT SELESAI\n\nFile ditemukan : {_totalFilesFound}\nBaris disimpan : {_totalRowsInserted}\n\nDebug Log:\n{_debugLog}",
                "Laporan Import",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information);
        }
    }
}