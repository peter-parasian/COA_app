using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using WpfApp1.Core.Models;
using WpfApp1.Data.Repositories;
using WpfApp1.Shared.Helpers;

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

            System.Threading.Tasks.Parallel.ForEach(filesToProcess, (fileItem) =>
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
            using var workbook = new ClosedXML.Excel.XLWorkbook(filePath);

            try
            {
                var sheet_YLB = workbook.Worksheets
                    .FirstOrDefault(w => w.Name.Trim().Equals("YLB 50", System.StringComparison.OrdinalIgnoreCase));

                if (sheet_YLB != null)
                {
                    int row = 3;
                    string currentProdDate = string.Empty;
                    int folderMonthNum = DateHelper.GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (true)
                    {
                        string sizeValue_YLB = sheet_YLB.Cell(row, "C").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue_YLB)) break;

                        string rawDateFromCell = sheet_YLB.Cell(row, "B").GetString().Trim();
                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = DateHelper.StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        BusbarRecord record = new BusbarRecord();
                        record.Size = StringHelper.CleanSizeText(sizeValue_YLB);
                        record.Year = year;
                        record.Month = month;
                        record.ProdDate = currentProdDate;

                        record.Thickness = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "G").GetString()), 2);
                        record.Width = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "I").GetString()), 2);
                        record.Radius = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "J").GetString()), 2);
                        record.Chamber = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "L").GetString()), 2);
                        record.Electric = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "U").GetString()), 2);
                        record.Oxygen = System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "X").GetString()), 2);

                        record.Spectro = StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "Y").GetString());
                        record.Resistivity = StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "T").GetString());

                        record.Length = (int)System.Math.Round(StringHelper.ParseCustomDecimal(sheet_YLB.Cell(row, "K").GetString()), 0);

                        record.Elongation = System.Math.Round(MathHelper.GetMergedOrAverageValue(sheet_YLB, row, "R"), 2);
                        record.Tensile = System.Math.Round(MathHelper.GetMergedOrAverageValue(sheet_YLB, row, "Q"), 2);

                        record.BendTest = sheet_YLB.Cell(row, "W").GetString();

                        busbarBag.Add(record);

                        row += 2;
                    }
                }
            }
            catch (System.Exception ex)
            {
                AppendDebug($"ERROR FILE (YLB): {System.IO.Path.GetFileName(filePath)} -> {ex.Message}");
            }

            ProcessTLJSheetToMemory(workbook, "TLJ 350", year, month, tlj350Bag);
            ProcessTLJSheetToMemory(workbook, "TLJ 500", year, month, tlj500Bag);
        }

        private void ProcessTLJSheetToMemory(
            ClosedXML.Excel.XLWorkbook workbook,
            string sheetName,
            string year,
            string month,
            System.Collections.Concurrent.ConcurrentBag<TLJRecord> bag)
        {
            int row = 3;
            try
            {
                var sheet = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals(sheetName, System.StringComparison.OrdinalIgnoreCase));
                if (sheet != null)
                {
                    string currentProdDate = string.Empty;
                    int folderMonthNum = DateHelper.GetMonthNumber(month);
                    int.TryParse(year, out int folderYearNum);

                    while (true)
                    {
                        string sizeValue = sheet.Cell(row, "D").GetString();
                        if (string.IsNullOrWhiteSpace(sizeValue)) break;

                        string rawDate = sheet.Cell(row, "B").GetString().Trim();
                        if (!string.IsNullOrEmpty(rawDate))
                        {
                            currentProdDate = DateHelper.StandardizeDate(rawDate, folderMonthNum, folderYearNum);
                        }

                        TLJRecord record = new TLJRecord
                        {
                            Size = StringHelper.CleanSizeText(sizeValue),
                            Year = year,
                            Month = month,
                            ProdDate = currentProdDate,
                            BatchNo = sheet.Cell(row, "C").GetString()
                        };

                        bag.Add(record);
                        row += 2;
                    }
                }
            }
            catch (System.Exception ex) { AppendDebug($"ERROR FILE ({sheetName}): {ex.Message}"); }
        }
    }
}