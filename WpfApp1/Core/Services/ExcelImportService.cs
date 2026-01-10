using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Core.Models;
using WpfApp1.Data.Repositories;
using WpfApp1.Shared.Helpers;
using ExcelDataReader;

namespace WpfApp1.Core.Services
{
    public class ExcelImportService
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";

        private BusbarRepository _repository;

        public event System.Action<string>? OnDebugMessage;

        public event System.Action<int, int>? OnProgress;

        public ExcelImportService(BusbarRepository repository)
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

            _repository.CreateBusbarTable(connection);

            TraverseFoldersAndImport(connection, transaction);

            _repository.FlushAll(connection, transaction);
            _repository.UpdateBusbarBatchNumbers(connection, transaction);
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
            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

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

            var concurrentBusbarData = new System.Collections.Concurrent.ConcurrentBag<BusbarRecord>();
            var concurrentTLJ350Data = new System.Collections.Concurrent.ConcurrentBag<TLJRecord>();
            var concurrentTLJ500Data = new System.Collections.Concurrent.ConcurrentBag<TLJRecord>();

            var parallelOptions = new System.Threading.Tasks.ParallelOptions
            {
                MaxDegreeOfParallelism = System.Environment.ProcessorCount
            };

            System.Threading.Tasks.Parallel.ForEach(filesToProcess, parallelOptions, (fileItem) =>
            {
                try
                {
                    string[] parts = fileItem.Split(new[] { '|' }, 3);
                    string filePath = parts[0];
                    string year = parts[1];
                    string month = parts[2];

                    ProcessSingleExcelFileToMemory(
                        filePath,
                        year,
                        month,
                        concurrentBusbarData,
                        concurrentTLJ350Data,
                        concurrentTLJ500Data
                    );

                    int currentIndex = System.Threading.Interlocked.Increment(ref _currentFileIndex);
                    OnProgress?.Invoke(currentIndex, TotalFilesFound);
                }
                catch (System.Exception ex)
                {
                    AppendDebug($"ERROR PARALLEL: {System.IO.Path.GetFileName(fileItem)} -> {ex.Message}");
                }
            });

            int rowsInsertedCount = 0;

            foreach (var record in concurrentBusbarData)
            {
                _repository.InsertBusbarRow(connection, transaction, record);
                rowsInsertedCount++;
            }

            foreach (var record in concurrentTLJ350Data)
            {
                _repository.InsertTLJ350_Row(connection, transaction, record);
                rowsInsertedCount++;
            }

            foreach (var record in concurrentTLJ500Data)
            {
                _repository.InsertTLJ500_Row(connection, transaction, record);
                rowsInsertedCount++;
            }

            TotalRowsInserted = rowsInsertedCount;
        }

        private void ProcessSingleExcelFileToMemory(
            string filePath,
            string year,
            string month,
            System.Collections.Concurrent.ConcurrentBag<BusbarRecord> busbarBag,
            System.Collections.Concurrent.ConcurrentBag<TLJRecord> tlj350Bag,
            System.Collections.Concurrent.ConcurrentBag<TLJRecord> tlj500Bag)
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

                System.Data.DataTable? tableYLB = null;
                foreach (System.Data.DataTable table in result.Tables)
                {
                    if (table.TableName.Trim().Equals("YLB 50", System.StringComparison.OrdinalIgnoreCase))
                    {
                        tableYLB = table;
                        break;
                    }
                }

                if (tableYLB != null)
                {
                    int rowIndex = 2;
                    int rowCount = tableYLB.Rows.Count;
                    string currentProdDate = string.Empty;
                    int folderMonthNum = DateHelper.GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (rowIndex < rowCount)
                    {
                        object rawSize = tableYLB.Rows[rowIndex][2];
                        string sizeValue_YLB = rawSize != null ? rawSize.ToString() ?? "" : "";

                        if (string.IsNullOrWhiteSpace(sizeValue_YLB)) break;

                        object rawDateObj = tableYLB.Rows[rowIndex][1];
                        string rawDateFromCell = rawDateObj != null ? rawDateObj.ToString()?.Trim() ?? "" : "";

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = DateHelper.StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        BusbarRecord record = new BusbarRecord();
                        record.Size = StringHelper.CleanSizeText(sizeValue_YLB);
                        record.Year = year;
                        record.Month = month;
                        record.ProdDate = currentProdDate;

                        string GetStr(int colIdx)
                        {
                            if (colIdx >= tableYLB.Columns.Count) return "";
                            object val = tableYLB.Rows[rowIndex][colIdx];
                            return val != null ? val.ToString() ?? "" : "";
                        }

                        record.Thickness = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(6)), 2);
                        record.Width = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(8)), 2);
                        record.Radius = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(9)), 2);
                        record.Chamber = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(11)), 2);
                        record.Electric = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(20)), 2);
                        record.Oxygen = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(23)), 2);

                        record.Spectro = StringHelper.ParseCustomDecimal(GetStr(24));
                        record.Resistivity = StringHelper.ParseCustomDecimal(GetStr(19));

                        record.Length = (int)System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(10)), 0);

                        object rawT1 = tableYLB.Rows[rowIndex][16];
                        object rawE1 = tableYLB.Rows[rowIndex][17];
                        double valT1 = StringHelper.ParseCustomDecimal(rawT1 != null ? rawT1.ToString() : "");
                        double valE1 = StringHelper.ParseCustomDecimal(rawE1 != null ? rawE1.ToString() : "");

                        double valT2 = 0;
                        double valE2 = 0;
                        if (rowIndex + 1 < rowCount)
                        {
                            object rawT2 = tableYLB.Rows[rowIndex + 1][16];
                            object rawE2 = tableYLB.Rows[rowIndex + 1][17];
                            valT2 = StringHelper.ParseCustomDecimal(rawT2 != null ? rawT2.ToString() : "");
                            valE2 = StringHelper.ParseCustomDecimal(rawE2 != null ? rawE2.ToString() : "");
                        }

                        var calcResult = MathHelper.CalculateTensileAndElongation(valT1, valT2, valE1, valE2);

                        record.Tensile = calcResult.Tensile;
                        record.Elongation = calcResult.Elongation;

                        record.BendTest = GetStr(22);

                        busbarBag.Add(record);

                        rowIndex += 2;
                    }
                }

                ProcessTLJSheetToMemory(result, "TLJ 350", year, month, tlj350Bag);
                ProcessTLJSheetToMemory(result, "TLJ 500", year, month, tlj500Bag);
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (READ): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }
        }

        private void ProcessTLJSheetToMemory(
            System.Data.DataSet dataSet,
            string sheetName,
            string year,
            string month,
            System.Collections.Concurrent.ConcurrentBag<TLJRecord> bag)
        {
            System.Data.DataTable? sheet = null;
            foreach (System.Data.DataTable table in dataSet.Tables)
            {
                if (table.TableName.Trim().Equals(sheetName, System.StringComparison.OrdinalIgnoreCase))
                {
                    sheet = table;
                    break;
                }
            }

            if (sheet != null)
            {
                int rowIndex = 2;
                int rowCount = sheet.Rows.Count;
                string currentProdDate = string.Empty;
                int folderMonthNum = DateHelper.GetMonthNumber(month);
                int.TryParse(year, out int folderYearNum);

                while (rowIndex < rowCount)
                {
                    if (3 >= sheet.Columns.Count) break;

                    object rawSize = sheet.Rows[rowIndex][3];
                    string sizeValue = rawSize != null ? rawSize.ToString() ?? "" : "";

                    if (string.IsNullOrWhiteSpace(sizeValue)) break;

                    object rawDate = sheet.Rows[rowIndex][1];
                    string rawDateStr = rawDate != null ? rawDate.ToString()?.Trim() ?? "" : "";

                    if (!string.IsNullOrEmpty(rawDateStr))
                    {
                        currentProdDate = DateHelper.StandardizeDate(rawDateStr, folderMonthNum, folderYearNum);
                    }

                    object rawBatch = sheet.Rows[rowIndex][2];
                    string batchValue = rawBatch != null ? rawBatch.ToString() ?? "" : "";

                    TLJRecord record = new TLJRecord
                    {
                        Size = StringHelper.CleanSizeText(sizeValue),
                        Year = year,
                        Month = month,
                        ProdDate = currentProdDate,
                        BatchNo = batchValue
                    };

                    bag.Add(record);
                    rowIndex += 2;
                }
            }
        }
    }
}