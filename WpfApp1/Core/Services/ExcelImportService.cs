using System;
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

            var localBusbarData = new System.Collections.Generic.List<BusbarRecord>(500);
            var localTLJ350Data = new System.Collections.Generic.List<TLJRecord>(100);
            var localTLJ500Data = new System.Collections.Generic.List<TLJRecord>(100);

            foreach (string fileItem in filesToProcess)
            {
                try
                {
                    string[] parts = fileItem.Split(new[] { '|' }, 3);
                    string filePath = parts[0];
                    string year = parts[1];
                    string month = parts[2];

                    localBusbarData.Clear();
                    localTLJ350Data.Clear();
                    localTLJ500Data.Clear();

                    ProcessSingleExcelFileToMemory(
                        filePath,
                        year,
                        month,
                        localBusbarData,
                        localTLJ350Data,
                        localTLJ500Data
                    );

                    foreach (var record in localBusbarData)
                    {
                        _repository.InsertBusbarRow(connection, transaction, record);
                        TotalRowsInserted++;
                    }

                    foreach (var record in localTLJ350Data)
                    {
                        _repository.InsertTLJ350_Row(connection, transaction, record);
                        TotalRowsInserted++;
                    }

                    foreach (var record in localTLJ500Data)
                    {
                        _repository.InsertTLJ500_Row(connection, transaction, record);
                        TotalRowsInserted++;
                    }

                    _currentFileIndex++;
                    OnProgress?.Invoke(_currentFileIndex, TotalFilesFound);

                    localBusbarData.Clear();
                    localTLJ350Data.Clear();
                    localTLJ500Data.Clear();

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
            System.Collections.Generic.List<BusbarRecord> busbarList,
            System.Collections.Generic.List<TLJRecord> tlj350List,
            System.Collections.Generic.List<TLJRecord> tlj500List)
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

                        if (string.IsNullOrWhiteSpace(sizeValue_YLB))
                        {
                            rowIndex += 2;
                            continue;
                        }

                        object rawDateObj = tableYLB.Rows[rowIndex][1];
                        string rawDateFromCell = rawDateObj != null ? rawDateObj.ToString()?.Trim() ?? "" : "";

                        if (!string.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = DateHelper.StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        string GetStr(int colIdx, int targetRowIndex)
                        {
                            if (colIdx >= tableYLB.Columns.Count) return "";
                            object val = tableYLB.Rows[targetRowIndex][colIdx];
                            return val?.ToString() ?? "";
                        }

                        BusbarRecord record = new BusbarRecord();
                        record.Size = StringHelper.CleanSizeText(sizeValue_YLB);
                        record.Year = year;
                        record.Month = month;
                        record.ProdDate = currentProdDate;

                        record.Thickness = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(6, rowIndex)), 2);
                        record.Width = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(8, rowIndex)), 2);
                        record.Radius = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(9, rowIndex)), 2);
                        record.Chamber = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(11, rowIndex)), 2);
                        record.Electric = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(20, rowIndex)), 2);
                        record.Oxygen = System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(23, rowIndex)), 2);

                        record.Spectro = StringHelper.ParseCustomDecimal(GetStr(24, rowIndex));
                        record.Resistivity = StringHelper.ParseCustomDecimal(GetStr(19, rowIndex));

                        record.Length = (int)System.Math.Round(StringHelper.ParseCustomDecimal(GetStr(10, rowIndex)), 0);

                        double valH1 = StringHelper.ParseCustomDecimal(GetStr(18, rowIndex));
                        double valH2 = 0.0;

                        if (rowIndex + 1 < rowCount)
                        {
                            valH2 = StringHelper.ParseCustomDecimal(GetStr(18, rowIndex + 1));
                        }

                        double maxHrf = System.Math.Max(valH1, valH2);
                        double convertedHv = 0.0;

                        if (maxHrf > 0.0)
                        {
                            try
                            {
                                convertedHv = HrfToHvConverter.Convert(maxHrf);
                            }
                            catch
                            {
                                convertedHv = 0.0;
                            }
                        }
                        record.Hardness = System.Math.Round(convertedHv, 2);

                        double valT1 = StringHelper.ParseCustomDecimal(GetStr(16, rowIndex));
                        double valE1 = StringHelper.ParseCustomDecimal(GetStr(17, rowIndex));

                        double valT2 = 0;
                        double valE2 = 0;
                        if (rowIndex + 1 < rowCount)
                        {
                            valT2 = StringHelper.ParseCustomDecimal(GetStr(16, rowIndex + 1));
                            valE2 = StringHelper.ParseCustomDecimal(GetStr(17, rowIndex + 1));
                        }

                        var calcResult = MathHelper.CalculateTensileAndElongation(valT1, valT2, valE1, valE2);

                        record.Tensile = calcResult.Tensile;
                        record.Elongation = calcResult.Elongation;
                        record.BendTest = GetStr(22, rowIndex);

                        busbarList.Add(record);

                        rowIndex += 2;
                    }
                }

                ProcessTLJSheetToMemory(result, "TLJ 350", year, month, tlj350List);
                ProcessTLJSheetToMemory(result, "TLJ 500", year, month, tlj500List);
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
            System.Collections.Generic.List<TLJRecord> list)
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
                    string sizeValue = rawSize?.ToString() ?? "";

                    if (string.IsNullOrWhiteSpace(sizeValue))
                    {
                        rowIndex += 2;
                        continue;
                    }

                    object rawDate = sheet.Rows[rowIndex][1];
                    string rawDateStr = rawDate?.ToString()?.Trim() ?? "";

                    if (!string.IsNullOrEmpty(rawDateStr))
                    {
                        currentProdDate = DateHelper.StandardizeDate(rawDateStr, folderMonthNum, folderYearNum);
                    }

                    object rawBatch = sheet.Rows[rowIndex][2];
                    string batchValue = rawBatch?.ToString() ?? "";

                    TLJRecord record = new TLJRecord
                    {
                        Size = StringHelper.CleanSizeText(sizeValue),
                        Year = year,
                        Month = month,
                        ProdDate = currentProdDate,
                        BatchNo = batchValue
                    };

                    list.Add(record);
                    rowIndex += 2;
                }
            }
        }
    }
}