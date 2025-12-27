using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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

                UpdateBusbarBatchNumbers(connection, transaction);

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

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ500;

                CREATE TABLE IF NOT EXISTS TLJ500 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT
                );
            ";

            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ350;

                CREATE TABLE IF NOT EXISTS TLJ350 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT,
                    Month_folder TEXT,
                    Batch_no TEXT,
                    Prod_date TEXT,
                    Size_mm TEXT
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
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);
            int row = 3;

            // --- Process YLB Sheet ---
            try
            {
                var sheet_YLB = workbook.Worksheets
                    .FirstOrDefault(w =>
                        w.Name.Trim().Equals(
                            "YLB 50",
                            System.StringComparison.OrdinalIgnoreCase));

                if (sheet_YLB == null)
                {
                    AppendDebug($"SKIP: Sheet 'YLB 50' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    string currentProdDate = string.Empty;

                    int folderMonthNum = GetMonthNumber(month);
                    int folderYearNum = 0;
                    int.TryParse(year, out folderYearNum);

                    while (true)
                    {
                        string sizeValue_YLB = sheet_YLB.Cell(row, "C").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue_YLB))
                            break;

                        string rawDateFromCell = sheet_YLB.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_YLB = CleanSizeText(sizeValue_YLB);

                        // Kelompok 2 Angka Belakang Koma
                        double rawThickness = ParseCustomDecimal(sheet_YLB.Cell(row, "G").GetString());
                        double valThickness = System.Math.Round(rawThickness, 2);

                        double rawWidth = ParseCustomDecimal(sheet_YLB.Cell(row, "I").GetString());
                        double valWidth = System.Math.Round(rawWidth, 2);

                        double rawRadius = ParseCustomDecimal(sheet_YLB.Cell(row, "J").GetString());
                        double valRadius = System.Math.Round(rawRadius, 2);

                        double rawChamber = ParseCustomDecimal(sheet_YLB.Cell(row, "L").GetString());
                        double valChamber = System.Math.Round(rawChamber, 2);

                        double rawElectric = ParseCustomDecimal(sheet_YLB.Cell(row, "U").GetString());
                        double valElectric = System.Math.Round(rawElectric, 2);

                        double rawOxygen = ParseCustomDecimal(sheet_YLB.Cell(row, "X").GetString());
                        double valOxygen = System.Math.Round(rawOxygen, 2);

                        // Kelompok Tanpa Batasan (Asli)
                        double valSpectro = ParseCustomDecimal(sheet_YLB.Cell(row, "Y").GetString());
                        double valResistivity = ParseCustomDecimal(sheet_YLB.Cell(row, "T").GetString());

                        // Kelompok Integer
                        double rawLength = ParseCustomDecimal(sheet_YLB.Cell(row, "K").GetString());
                        double valLength = System.Math.Round(rawLength, 0);

                        // Kelompok Merge/Rata-rata
                        double rawElongation = GetMergedOrAverageValue(sheet_YLB, row, "R");
                        double valElongation = System.Math.Round(rawElongation, 2);

                        double rawTensile = GetMergedOrAverageValue(sheet_YLB, row, "Q");
                        double valTensile = System.Math.Round(rawTensile, 2);

                        string valBendTest = sheet_YLB.Cell(row, "W").GetString();

                        InsertBusbarRow(
                            connection, transaction,
                            cleanSize_YLB, year, month, currentProdDate,
                            valThickness, valWidth, valLength, valRadius, valChamber,
                            valElectric, valResistivity, valElongation, valTensile,
                            valBendTest, valSpectro, valOxygen
                        );

                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (YLB): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            // --- Process TLJ 350 Sheet ---
            row = 3;
            try
            {
                var sheet_TLJ350 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 350", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ350 == null)
                {
                    AppendDebug($"SKIP: Sheet 'TLJ350' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    string currentProdDate = string.Empty;

                    int folderMonthNum = GetMonthNumber(month);
                    int folderYearNum = 0;
                    int.TryParse(year, out folderYearNum);

                    while (true)
                    {
                        string sizeValue_TLJ350 = sheet_TLJ350.Cell(row, "D").GetString();

                        if (string.IsNullOrWhiteSpace(sizeValue_TLJ350))
                            break;

                        string rawDateFromCell = sheet_TLJ350.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_TLJ350 = CleanSizeText(sizeValue_TLJ350);

                        string batchValue = sheet_TLJ350.Cell(row, "C").GetString();

                        InsertTLJ350_Row(connection, transaction, cleanSize_TLJ350, year, month, currentProdDate, batchValue);

                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (TLJ350): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            // --- Process TLJ 500 Sheet ---
            row = 3;
            try
            {
                var sheet_TLJ500 = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals("TLJ 500", System.StringComparison.OrdinalIgnoreCase));
                if (sheet_TLJ500 == null)
                {
                    AppendDebug($"SKIP: Sheet 'TLJ500' tidak ditemukan -> {System.IO.Path.GetFileName(filePath)}");
                }
                else
                {
                    string currentProdDate = string.Empty;

                    int folderMonthNum = GetMonthNumber(month);
                    int folderYearNum = 0;
                    int.TryParse(year, out folderYearNum);

                    while (true)
                    {
                        string sizeValue_TLJ500 = sheet_TLJ500.Cell(row, "D").GetString();

                        if (string.IsNullOrWhiteSpace(sizeValue_TLJ500))
                            break;

                        string rawDateFromCell = sheet_TLJ500.Cell(row, "B").GetString().Trim();

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string cleanSize_TLJ500 = CleanSizeText(sizeValue_TLJ500);

                        string batchValue = sheet_TLJ500.Cell(row, "C").GetString();

                        InsertTLJ500_Row(connection, transaction, cleanSize_TLJ500, year, month, currentProdDate, batchValue);

                        _totalRowsInserted++;
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (TLJ500): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private string StandardizeDate(string rawDate, int expectedMonth, int expectedYear)
        {
            if (string.IsNullOrWhiteSpace(rawDate)) return string.Empty;

            if (System.DateTime.TryParse(rawDate, out System.DateTime parsedDate))
            {
                if (expectedYear > 2000 && parsedDate.Year != expectedYear)
                {
                    parsedDate = new System.DateTime(expectedYear, parsedDate.Month, parsedDate.Day);
                }

                if (expectedMonth > 0 && parsedDate.Month != expectedMonth)
                {
                    if (parsedDate.Day <= 12)
                    {
                        int newMonth = parsedDate.Day;
                        int newDay = parsedDate.Month;

                        if (newMonth == expectedMonth)
                        {
                            parsedDate = new System.DateTime(parsedDate.Year, newMonth, newDay);
                        }
                    }
                }

                return parsedDate.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            }

            return rawDate;
        }

        private int GetMonthNumber(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName)) return 0;
            try
            {
                return System.DateTime.ParseExact(monthName, "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
            }
            catch
            {
                return 0;
            }
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

        private System.String CleanSizeText(System.String raw)
        {
            if (System.String.IsNullOrWhiteSpace(raw)) return System.String.Empty;

            System.String text = raw.ToUpper();

            int start = -1;
            for (int i = 0; i < text.Length; i++)
            {
                if (System.Char.IsDigit(text[i]))
                {
                    start = i;
                    break;
                }
            }

            if (start == -1) return System.String.Empty;

            System.String substring = text.Substring(start);

            int xIndex = substring.IndexOf('X');
            if (xIndex == -1) return System.String.Empty;

            if (xIndex + 1 >= substring.Length || !System.Char.IsDigit(substring[xIndex + 1]))
                return System.String.Empty;

            int end = xIndex + 1;
            while (end < substring.Length && System.Char.IsDigit(substring[end]))
                end++;

            System.String result = substring.Substring(0, end).Trim();

            System.String remaining = substring.Substring(end).Trim();

            System.String keyword = System.String.Empty;

            System.String cleanRemaining = System.String.Empty;
            for (int i = 0; i < remaining.Length; i++)
            {
                if (System.Char.IsLetterOrDigit(remaining[i]))
                {
                    cleanRemaining += remaining[i];
                }
            }

            if (cleanRemaining.Contains("FR"))
            {
                int frIndex = cleanRemaining.IndexOf("FR");
                if (frIndex >= 0)
                {
                    keyword = "FR";
                }
            }
            else if (!System.String.IsNullOrEmpty(cleanRemaining))
            {
                for (int i = 0; i < cleanRemaining.Length; i++)
                {
                    if (cleanRemaining[i] == 'B' && i + 1 < cleanRemaining.Length &&
                        System.Char.IsDigit(cleanRemaining[i + 1]))
                    {
                        int bStart = i;
                        int bEnd = i + 1;
                        while (bEnd < cleanRemaining.Length && System.Char.IsDigit(cleanRemaining[bEnd]))
                        {
                            bEnd++;
                        }

                        if (bEnd - bStart >= 2)
                        {
                            keyword = cleanRemaining.Substring(bStart, bEnd - bStart);
                            break;
                        }
                    }
                }
            }

            if (!System.String.IsNullOrEmpty(keyword))
            {
                result = result + " " + keyword;
            }

            return result.Trim();
        }

        private void InsertBusbarRow(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string size, string year, string month, string prodDate,
            double thickness, double width, double length, double radius, double chamber,
            double electric, double resistivity, double elongation, double tensile,
            string bendTest, double spectro, double oxygen)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                INSERT INTO Busbar (
                    Size_mm, Year_folder, Month_folder, Prod_date, 
                    Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                    Electric_IACS, Weight, Elongation, Tensile,
                    Bend_test, Spectro_Cu, Oxygen
                )
                VALUES (
                    @Size, @Year, @Month, @ProdDate,
                    @Thickness, @Width, @Length, @Radius, @Chamber,
                    @Electric, @Resistivity, @Elongation, @Tensile,
                    @Bend, @Spectro, @Oxygen
                );
            ";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year.Trim());
            cmd.Parameters.AddWithValue("@Month", month.Trim());
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

        private void InsertTLJ350_Row(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string size, string year, string month, string prodDate, string batchNum)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                INSERT INTO TLJ350 (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no)
                VALUES (@Size, @Year, @Month, @ProdDate, @Batch);
            ";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year.Trim());
            cmd.Parameters.AddWithValue("@Month", month.Trim());
            cmd.Parameters.AddWithValue("@ProdDate", prodDate);
            object batchVal = string.IsNullOrEmpty(batchNum) ? System.DBNull.Value : batchNum;
            cmd.Parameters.AddWithValue("@Batch", batchVal);

            cmd.ExecuteNonQuery();
        }

        private void InsertTLJ500_Row(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string size, string year, string month, string prodDate, string batchNum)
        {
            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                INSERT INTO TLJ500 (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no)
                VALUES (@Size, @Year, @Month, @ProdDate, @Batch);
            ";

            cmd.Parameters.AddWithValue("@Size", size);
            cmd.Parameters.AddWithValue("@Year", year.Trim());
            cmd.Parameters.AddWithValue("@Month", month.Trim());
            cmd.Parameters.AddWithValue("@ProdDate", prodDate);
            object batchVal = string.IsNullOrEmpty(batchNum) ? System.DBNull.Value : batchNum;
            cmd.Parameters.AddWithValue("@Batch", batchVal);

            cmd.ExecuteNonQuery();
        }

        private void UpdateBusbarBatchNumbers(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            try
            {
                using var selectBusbarCmd = connection.CreateCommand();
                selectBusbarCmd.Transaction = transaction;
                selectBusbarCmd.CommandText = @"
                    SELECT Id, Size_mm, Prod_date 
                    FROM Busbar 
                    WHERE (Batch_no IS NULL OR Batch_no = '')
                    ORDER BY Prod_date, Id
                ";

                using var busbarReader = selectBusbarCmd.ExecuteReader();
                while (busbarReader.Read())
                {
                    int busbarId = busbarReader.GetInt32(0);
                    string size_mm = busbarReader.GetString(1);
                    string prod_date = busbarReader.GetString(2);

                    string targetTable = DetermineTLJTable(size_mm);

                    string batchNumbers = FindBatchNumbers(connection, transaction, targetTable, size_mm, prod_date);

                    if (!System.String.IsNullOrEmpty(batchNumbers))
                    {
                        UpdateBusbarBatch(connection, transaction, busbarId, batchNumbers);
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR UpdateBusbarBatchNumbers: {ex.Message}");
                throw;
            }
        }

        private string DetermineTLJTable(string size_mm)
        {
            string cleanSize = size_mm.ToUpper().Replace(" ", "");

            int xIndex = cleanSize.IndexOf('X');
            if (xIndex == -1) return "TLJ500"; 

            string beforeX = cleanSize.Substring(0, xIndex);
            string afterX = cleanSize.Substring(xIndex + 1);

            string afterXDigits = "";
            for (int i = 0; i < afterX.Length; i++)
            {
                if (System.Char.IsDigit(afterX[i]))
                {
                    afterXDigits += afterX[i];
                }
                else
                {
                    break;
                }
            }

            if (int.TryParse(beforeX, out int firstDimension) &&
                int.TryParse(afterXDigits, out int secondDimension))
            {
                if (firstDimension <= 10 && secondDimension <= 100)
                {
                    return "TLJ350";
                }
            }

            return "TLJ500";
        }

        private string FindBatchNumbers(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            string tableName,
            string size_mm,
            string targetDate)
        {
            try
            {
                System.DateTime targetDateTime;
                if (!System.DateTime.TryParseExact(
                    targetDate,
                    "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out targetDateTime))
                {
                    if (!System.DateTime.TryParse(targetDate, out targetDateTime))
                    {
                        return string.Empty;
                    }
                }

                using var cmdSameDate = connection.CreateCommand();
                cmdSameDate.Transaction = transaction;
                cmdSameDate.CommandText = $@"
                    SELECT Batch_no, Prod_date 
                    FROM {tableName} 
                    WHERE Size_mm = @Size_mm 
                      AND Prod_date = @TargetDate
                    ORDER BY Prod_date DESC
                ";

                cmdSameDate.Parameters.AddWithValue("@Size_mm", size_mm);
                cmdSameDate.Parameters.AddWithValue("@TargetDate", targetDate);

                using var readerSameDate = cmdSameDate.ExecuteReader();
                if (readerSameDate.HasRows)
                {
                    return ExtractBatchNumbers(readerSameDate);
                }

                using var cmdBeforeDate = connection.CreateCommand();
                cmdBeforeDate.Transaction = transaction;
                cmdBeforeDate.CommandText = $@"
                    SELECT Batch_no, Prod_date 
                    FROM {tableName} 
                    WHERE Size_mm = @Size_mm 
                      AND Prod_date < @TargetDate
                    ORDER BY Prod_date DESC
                    LIMIT 1
                ";

                cmdBeforeDate.Parameters.AddWithValue("@Size_mm", size_mm);
                cmdBeforeDate.Parameters.AddWithValue("@TargetDate", targetDate);

                using var readerBeforeDate = cmdBeforeDate.ExecuteReader();
                if (readerBeforeDate.HasRows)
                {
                    return ExtractBatchNumbers(readerBeforeDate);
                }

                return string.Empty;
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FindBatchNumbers for {size_mm} on {targetDate}: {ex.Message}");
                return string.Empty;
            }
        }

        private string ExtractBatchNumbers(Microsoft.Data.Sqlite.SqliteDataReader reader)
        {
            System.Collections.Generic.List<string> batchList = new System.Collections.Generic.List<string>();

            while (reader.Read())
            {
                if (!reader.IsDBNull(0))
                {
                    string batchData = reader.GetString(0);

                    if (!System.String.IsNullOrEmpty(batchData))
                    {
                        string[] batches = batchData.Split(
                            new[] { '\n', '\r' },
                            System.StringSplitOptions.RemoveEmptyEntries
                        );

                        foreach (string batch in batches)
                        {
                            string trimmedBatch = batch.Trim();
                            if (!System.String.IsNullOrEmpty(trimmedBatch))
                            {
                                batchList.Add(trimmedBatch);
                            }
                        }
                    }
                }
            }

            return System.String.Join("\n", batchList);
        }

        private void UpdateBusbarBatch(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            int busbarId,
            string batchNumbers)
        {
            using var updateCmd = connection.CreateCommand();
            updateCmd.Transaction = transaction;
            updateCmd.CommandText = @"
                UPDATE Busbar 
                SET Batch_no = @BatchNumbers 
                WHERE Id = @Id
            ";

            updateCmd.Parameters.AddWithValue("@BatchNumbers", batchNumbers);
            updateCmd.Parameters.AddWithValue("@Id", busbarId);

            updateCmd.ExecuteNonQuery();
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