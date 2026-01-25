using ExcelDataReader;

namespace WpfApp1.Core.Services
{
    public class ExcelImportService
    {
        private const System.String ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\H\OPERATOR\COPPER BUSBAR & STRIP";

        private WpfApp1.Data.Repositories.BusbarRepository _repository;

        public event System.Action<System.String>? OnDebugMessage;

        public event System.Action<System.Int32, System.Int32>? OnProgress;

        public ExcelImportService(WpfApp1.Data.Repositories.BusbarRepository repository)
        {
            _repository = repository;

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public System.Int32 TotalFilesFound { get; private set; }
        public System.Int32 TotalRowsInserted { get; private set; }
        private System.Int32 _currentFileIndex = 0;

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

        private void AppendDebug(System.String message)
        {
            if (OnDebugMessage != null) OnDebugMessage.Invoke(message);
        }

        private void CountTotalFiles()
        {
            if (!System.IO.Directory.Exists(ExcelRootFolder))
            {
                throw new System.IO.DirectoryNotFoundException($"Folder root Excel tidak ditemukan: {ExcelRootFolder}");
            }

            foreach (System.String yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                foreach (System.String monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    foreach (System.String file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        System.String fileName = System.IO.Path.GetFileName(file);
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

            System.Collections.Generic.List<System.String> filesToProcess = new System.Collections.Generic.List<System.String>();

            foreach (System.String yearDir in System.IO.Directory.GetDirectories(ExcelRootFolder))
            {
                System.String year = new System.IO.DirectoryInfo(yearDir).Name.Trim();

                foreach (System.String monthDir in System.IO.Directory.GetDirectories(yearDir))
                {
                    System.String rawMonth = new System.IO.DirectoryInfo(monthDir).Name.Trim();
                    System.String normalizedMonth = WpfApp1.Shared.Helpers.DateHelper.NormalizeMonthFolder(rawMonth);

                    foreach (System.String file in System.IO.Directory.GetFiles(monthDir, "*.xlsx"))
                    {
                        System.String fileName = System.IO.Path.GetFileName(file);

                        if (fileName.StartsWith("~$"))
                        {
                            continue;
                        }

                        filesToProcess.Add(file + "|" + year + "|" + normalizedMonth);
                    }
                }
            }

            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.BusbarRecord> concurrentBusbarData =
                new System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.BusbarRecord>();
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord> concurrentTLJ350Data =
                new System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord>();
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord> concurrentTLJ500Data =
                new System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord>();

            System.Threading.Tasks.ParallelOptions parallelOptions = new System.Threading.Tasks.ParallelOptions
            {
                MaxDegreeOfParallelism = System.Math.Max(1, System.Environment.ProcessorCount - 1)
            };

            System.Threading.Tasks.Parallel.ForEach(filesToProcess, parallelOptions, (fileItem) =>
            {
                try
                {
                    System.String[] parts = fileItem.Split(new[] { '|' }, 3);
                    System.String filePath = parts[0];
                    System.String year = parts[1];
                    System.String month = parts[2];

                    ProcessSingleExcelFileToMemory(
                        filePath,
                        year,
                        month,
                        concurrentBusbarData,
                        concurrentTLJ350Data,
                        concurrentTLJ500Data
                    );

                    System.Int32 currentIndex = System.Threading.Interlocked.Increment(ref _currentFileIndex);
                    OnProgress?.Invoke(currentIndex, TotalFilesFound);
                }
                catch (System.Exception ex)
                {
                    AppendDebug($"ERROR PARALLEL: {System.IO.Path.GetFileName(fileItem)} -> {ex.Message}");
                }
            });

            System.Int32 rowsInsertedCount = 0;

            foreach (WpfApp1.Core.Models.BusbarRecord record in concurrentBusbarData)
            {
                _repository.InsertBusbarRow(connection, transaction, record);
                rowsInsertedCount++;
            }

            foreach (WpfApp1.Core.Models.TLJRecord record in concurrentTLJ350Data)
            {
                _repository.InsertTLJ350_Row(connection, transaction, record);
                rowsInsertedCount++;
            }

            foreach (WpfApp1.Core.Models.TLJRecord record in concurrentTLJ500Data)
            {
                _repository.InsertTLJ500_Row(connection, transaction, record);
                rowsInsertedCount++;
            }

            TotalRowsInserted = rowsInsertedCount;
        }

        private void ProcessSingleExcelFileToMemory(
            System.String filePath,
            System.String year,
            System.String month,
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.BusbarRecord> busbarBag,
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord> tlj350Bag,
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord> tlj500Bag)
        {
            System.IO.FileStream? stream = null;
            ExcelDataReader.IExcelDataReader? reader = null;

            try
            {
                stream = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);

                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                System.Data.DataSet result = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration()
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
                    System.Int32 rowIndex = 2;
                    System.Int32 rowCount = tableYLB.Rows.Count;
                    System.String currentProdDate = System.String.Empty;
                    System.Int32 folderMonthNum = WpfApp1.Shared.Helpers.DateHelper.GetMonthNumber(month);
                    System.Int32.TryParse(year, out int folderYearNum);

                    while (rowIndex < rowCount)
                    {
                        System.Object rawSize = tableYLB.Rows[rowIndex][2];
                        System.String sizeValue_YLB = rawSize != null ? rawSize.ToString() ?? "" : "";

                        if (System.String.IsNullOrWhiteSpace(sizeValue_YLB))
                        {
                            break;
                        }

                        System.Object rawDateObj = tableYLB.Rows[rowIndex][1];
                        System.String rawDateFromCell = rawDateObj != null ? rawDateObj.ToString()?.Trim() ?? "" : "";

                        if (!System.String.IsNullOrEmpty(rawDateFromCell))
                        {
                            currentProdDate = WpfApp1.Shared.Helpers.DateHelper.StandardizeDate(rawDateFromCell, folderMonthNum, folderYearNum);
                        }

                        WpfApp1.Core.Models.BusbarRecord record = new WpfApp1.Core.Models.BusbarRecord();
                        record.Size = WpfApp1.Shared.Helpers.StringHelper.CleanSizeText(sizeValue_YLB);
                        record.Year = year;
                        record.Month = month;
                        record.ProdDate = currentProdDate;

                        System.String GetStr(System.Int32 colIdx)
                        {
                            if (colIdx >= tableYLB.Columns.Count) return "";
                            System.Object val = tableYLB.Rows[rowIndex][colIdx];
                            return val != null ? val.ToString() ?? "" : "";
                        }

                        record.Thickness = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(6)), 2);
                        record.Width = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(8)), 2);
                        record.Radius = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(9)), 2);
                        record.Chamber = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(11)), 2);
                        record.Electric = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(20)), 2);
                        record.Oxygen = System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(23)), 2);

                        record.Spectro = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(24));
                        record.Resistivity = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(19));

                        record.Length = (int)System.Math.Round(WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(GetStr(10)), 0);

                        System.Object rawT1 = tableYLB.Rows[rowIndex][16];
                        System.Object rawE1 = tableYLB.Rows[rowIndex][17];
                        System.Double valT1 = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(rawT1 != null ? rawT1.ToString() : "");
                        System.Double valE1 = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(rawE1 != null ? rawE1.ToString() : "");

                        System.Double valT2 = 0;
                        System.Double valE2 = 0;
                        if (rowIndex + 1 < rowCount)
                        {
                            System.Object rawT2 = tableYLB.Rows[rowIndex + 1][16];
                            System.Object rawE2 = tableYLB.Rows[rowIndex + 1][17];
                            valT2 = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(rawT2 != null ? rawT2.ToString() : "");
                            valE2 = WpfApp1.Shared.Helpers.StringHelper.ParseCustomDecimal(rawE2 != null ? rawE2.ToString() : "");
                        }

                        (System.Double Tensile, System.Double Elongation) calcResult = WpfApp1.Shared.Helpers.MathHelper.CalculateTensileAndElongation(valT1, valT2, valE1, valE2);

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
            finally
            {
                if (reader != null)
                {
                    reader.Dispose();
                }
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }

        private void ProcessTLJSheetToMemory(
            System.Data.DataSet dataSet,
            System.String sheetName,
            System.String year,
            System.String month,
            System.Collections.Concurrent.ConcurrentBag<WpfApp1.Core.Models.TLJRecord> bag)
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
                System.Int32 rowIndex = 2;
                System.Int32 rowCount = sheet.Rows.Count;
                System.String currentProdDate = System.String.Empty;
                System.Int32 folderMonthNum = WpfApp1.Shared.Helpers.DateHelper.GetMonthNumber(month);
                System.Int32.TryParse(year, out int folderYearNum);

                while (rowIndex < rowCount)
                {
                    if (3 >= sheet.Columns.Count)
                    {
                        break;
                    }

                    System.Object rawSize = sheet.Rows[rowIndex][3];
                    System.String sizeValue = rawSize != null ? rawSize.ToString() ?? "" : "";

                    if (System.String.IsNullOrWhiteSpace(sizeValue))
                    {
                        break;
                    }

                    System.Object rawDate = sheet.Rows[rowIndex][1];
                    System.String rawDateStr = rawDate != null ? rawDate.ToString()?.Trim() ?? "" : "";

                    if (!System.String.IsNullOrEmpty(rawDateStr))
                    {
                        currentProdDate = WpfApp1.Shared.Helpers.DateHelper.StandardizeDate(rawDateStr, folderMonthNum, folderYearNum);
                    }

                    System.Object rawBatch = sheet.Rows[rowIndex][2];
                    System.String batchValue = rawBatch != null ? rawBatch.ToString() ?? "" : "";

                    WpfApp1.Core.Models.TLJRecord record = new WpfApp1.Core.Models.TLJRecord
                    {
                        Size = WpfApp1.Shared.Helpers.StringHelper.CleanSizeText(sizeValue),
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