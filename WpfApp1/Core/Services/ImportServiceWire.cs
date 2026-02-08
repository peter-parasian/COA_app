using System;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Core.Models;
using WpfApp1.Data.Repositories;
using WpfApp1.Shared.Helpers;
using ExcelDataReader;

namespace WpfApp1.Core.Services
{
    public class ImportServiceWire
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\WIRE";

        private WireRepository _repository;

        public event System.Action<string>? OnDebugMessage;
        public event System.Action<int, int>? OnProgress;

        public ImportServiceWire(WireRepository repository)
        {
            _repository = repository;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public int TotalFilesFound { get; private set; }
        public int TotalRowsInserted { get; private set; }
        private int _currentFileIndex = 0;

        public void Import(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            TotalFilesFound = 0;
            TotalRowsInserted = 0;
            _currentFileIndex = 0;

            CountTotalFiles();

            OnProgress?.Invoke(0, TotalFilesFound);

            _repository.CreateWireTables(connection);
            _repository.ClearCurrentTable(connection, transaction);

            TraverseFoldersAndImport(connection, transaction);

            _repository.FlushAll(connection, transaction);
        }

        private void AppendDebug(string message)
        {
            if (OnDebugMessage != null) OnDebugMessage.Invoke(message);
        }

        private void CountTotalFiles()
        {
            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

            foreach (string yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                foreach (string monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    foreach (string file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);
                        if (!fileName.StartsWith("~$"))
                        {
                            TotalFilesFound++;
                        }
                    }
                }
            }
        }

        private void TraverseFoldersAndImport(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            if (!System.IO.Directory.Exists(ExcelRootFolder)) return;

            var filesToProcess = new System.Collections.Generic.List<string>();

            foreach (string yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                string year = new System.IO.DirectoryInfo(yearDir).Name.Trim();

                foreach (string monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    string rawMonth = new System.IO.DirectoryInfo(monthDir).Name.Trim();
                    string normalizedMonth = DateHelper.NormalizeMonthFolder(rawMonth);

                    foreach (string file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        string fileName = System.IO.Path.GetFileName(file);

                        if (fileName.StartsWith("~$"))
                            continue;

                        filesToProcess.Add(file + "|" + year + "|" + normalizedMonth);
                    }
                }
            }

            var localWireData = new System.Collections.Generic.List<WireRecord>(500);

            foreach (string fileItem in filesToProcess)
            {
                try
                {
                    string[] parts = fileItem.Split(new[] { '|' }, 3);
                    string filePath = parts[0];
                    string year = parts[1];
                    string month = parts[2];

                    localWireData.Clear();

                    ProcessSingleExcelFileToMemory(
                        filePath,
                        year,
                        month,
                        localWireData
                    );

                    foreach (var record in localWireData)
                    {
                        _repository.InsertIntoCurrent(connection, transaction, record);
                        TotalRowsInserted++;
                    }

                    _currentFileIndex++;
                    OnProgress?.Invoke(_currentFileIndex, TotalFilesFound);

                    localWireData.Clear();

                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                }
                catch (System.Exception ex)
                {
                    AppendDebug($"ERROR SERIAL: {System.IO.Path.GetFileName(fileItem)} -> {ex.Message}");
                }
            }
        }

        private void ProcessSingleExcelFileToMemory(
            string filePath,
            string year,
            string month,
            System.Collections.Generic.List<WireRecord> wireList)
        {
            try
            {
                using var stream = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                using var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var result = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataReader.ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                var validSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Wire 1.20", "Wire 1.24", "Wire 1.38", "Wire 2.60", "Wire 1.50", "Wire 1.60"
                };

                foreach (System.Data.DataTable table in result.Tables)
                {
                    string sheetName = table.TableName.Trim();
                    bool isValid = validSheets.Contains(sheetName) || sheetName.StartsWith("Wire", StringComparison.OrdinalIgnoreCase);

                    if (isValid)
                    {
                        ProcessWireSheetToMemory(table, sheetName, year, month, wireList);
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (READ): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private void ProcessWireSheetToMemory(
            System.Data.DataTable sheet,
            string sheetName,
            string year,
            string month,
            System.Collections.Generic.List<WireRecord> list)
        {
            string size = sheetName.Replace("Wire ", "").Trim();

            int rowIndex = 4;
            int rowCount = sheet.Rows.Count;
            string currentProdDate = string.Empty;
            string currentCustomer = string.Empty;
            int folderMonthNum = DateHelper.GetMonthNumber(month);
            int.TryParse(year, out int folderYearNum);

            while (rowIndex < rowCount)
            {
                string GetStr(int colIdx, int targetRowIndex)
                {
                    if (targetRowIndex >= rowCount || colIdx >= sheet.Columns.Count) return "";
                    object val = sheet.Rows[targetRowIndex][colIdx];
                    return val?.ToString()?.Trim() ?? "";
                }

                double GetDualValue(int colIdx)
                {
                    string s1 = GetStr(colIdx, rowIndex);
                    string s2 = GetStr(colIdx, rowIndex + 1);

                    double v1 = StringHelper.ParseCustomDecimal(s1);
                    double v2 = StringHelper.ParseCustomDecimal(s2);

                    if (v1 > 0 && v2 > 0) return (v1 + v2) / 2.0;
                    if (v1 > 0) return v1;
                    return v2;
                }

                string rawDate1 = GetStr(1, rowIndex);
                string rawDate2 = GetStr(1, rowIndex + 1);
                string effectiveDate = !string.IsNullOrEmpty(rawDate1) ? rawDate1 : rawDate2;

                if (!string.IsNullOrEmpty(effectiveDate) && !effectiveDate.Equals("Date", StringComparison.OrdinalIgnoreCase))
                {
                    string standardized = DateHelper.StandardizeDate(effectiveDate, folderMonthNum, folderYearNum);
                    if (!string.IsNullOrEmpty(standardized))
                    {
                        currentProdDate = standardized;
                    }
                }

                string rawLot1 = GetStr(2, rowIndex);
                string rawLot2 = GetStr(2, rowIndex + 1);
                string effectiveLot = rawLot1;

                if (!string.IsNullOrEmpty(rawLot2))
                {
                    effectiveLot = string.IsNullOrEmpty(effectiveLot) ? rawLot2 : effectiveLot + " " + rawLot2;
                }

                if (string.IsNullOrWhiteSpace(effectiveLot) || effectiveLot.Contains("Lot Copper", StringComparison.OrdinalIgnoreCase))
                {
                    rowIndex += 2;
                    continue;
                }

                effectiveLot = effectiveLot.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ").Trim();

                string[] lotParts = effectiveLot.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (lotParts.Length >= 2)
                {
                    rowIndex += 2;
                    continue;
                }

                string rawCust1 = GetStr(3, rowIndex);
                string rawCust2 = GetStr(3, rowIndex + 1);
                string effectiveCust = !string.IsNullOrEmpty(rawCust1) ? rawCust1 : rawCust2;

                if (!string.IsNullOrEmpty(effectiveCust))
                {
                    currentCustomer = effectiveCust;
                }

                WireRecord record = new WireRecord();
                record.Size = size;
                record.Date = currentProdDate;
                record.Lot = effectiveLot.Trim();
                record.Customer = currentCustomer;

                record.Diameter = System.Math.Round(GetDualValue(6), 2);
                record.Yield = System.Math.Round(GetDualValue(7), 2);
                record.Tensile = System.Math.Round(GetDualValue(8), 2);
                record.Elongation = System.Math.Round(GetDualValue(9), 2);
                record.IACS = System.Math.Round(GetDualValue(10), 2);

                list.Add(record);

                rowIndex += 2;
            }
        }
    }
}