using System;
using System.Collections.Generic;
using WpfApp1.Core.Models;
using WpfApp1.Shared.Helpers;
using WpfApp1.Data.Database;

namespace WpfApp1.Data.Repositories
{
    public class BusbarRepository
    {
        //private const int BATCH_SIZE = 1000;
        private const int BATCH_SIZE = 5000;

        private System.Collections.Generic.List<BusbarRecord> _busbarBatchBuffer = new System.Collections.Generic.List<BusbarRecord>(BATCH_SIZE);
        private System.Collections.Generic.List<TLJRecord> _tlj350BatchBuffer = new System.Collections.Generic.List<TLJRecord>(BATCH_SIZE);
        private System.Collections.Generic.List<TLJRecord> _tlj500BatchBuffer = new System.Collections.Generic.List<TLJRecord>(BATCH_SIZE);

        private readonly SqliteContext _dbContext;

        public BusbarRepository(SqliteContext dbContext)
        {
            _dbContext = dbContext;
        }

        public void CreateBusbarTable(Microsoft.Data.Sqlite.SqliteConnection connection)
        {
            _busbarBatchBuffer.Clear();
            _tlj350BatchBuffer.Clear();
            _tlj500BatchBuffer.Clear();

            using var cmd = connection.CreateCommand();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Busbar;
                CREATE TABLE IF NOT EXISTS Busbar (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT,
                    Thickness_mm REAL, Width_mm REAL, Length INTEGER, Radius REAL, Chamber_mm REAL,
                    Electric_IACS REAL, Weight REAL, Elongation REAL, Tensile REAL, Bend_test TEXT,
                    Spectro_Cu REAL, Oxygen REAL
                );
                CREATE INDEX IF NOT EXISTS IDX_Busbar_LookUp ON Busbar(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ500;
                CREATE TABLE IF NOT EXISTS TLJ500 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT
                );
                CREATE INDEX IF NOT EXISTS IDX_TLJ500_LookUp ON TLJ500(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS TLJ350;
                CREATE TABLE IF NOT EXISTS TLJ350 (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Year_folder TEXT, Month_folder TEXT, Batch_no TEXT, Prod_date TEXT, Size_mm TEXT
                );
                CREATE INDEX IF NOT EXISTS IDX_TLJ350_LookUp ON TLJ350(Size_mm, Prod_date);
            ";
            cmd.ExecuteNonQuery();
        }

        public void InsertBusbarRow(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, BusbarRecord record)
        {
            _busbarBatchBuffer.Add(record);
            if (_busbarBatchBuffer.Count >= BATCH_SIZE)
            {
                FlushBusbarBatch(connection, transaction);
            }
        }

        public void InsertTLJ350_Row(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, TLJRecord record)
        {
            _tlj350BatchBuffer.Add(record);
            if (_tlj350BatchBuffer.Count >= BATCH_SIZE)
            {
                FlushTLJBatch(connection, transaction, "TLJ350", _tlj350BatchBuffer);
            }
        }

        public void InsertTLJ500_Row(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, TLJRecord record)
        {
            _tlj500BatchBuffer.Add(record);
            if (_tlj500BatchBuffer.Count >= BATCH_SIZE)
            {
                FlushTLJBatch(connection, transaction, "TLJ500", _tlj500BatchBuffer);
            }
        }

        public void FlushAll(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            FlushBusbarBatch(connection, transaction);
            FlushTLJBatch(connection, transaction, "TLJ350", _tlj350BatchBuffer);
            FlushTLJBatch(connection, transaction, "TLJ500", _tlj500BatchBuffer);
        }

        private void FlushBusbarBatch(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            if (_busbarBatchBuffer.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder(BATCH_SIZE * 250); //200
            sqlBuilder.Append(@"INSERT INTO Busbar (
                Size_mm, Year_folder, Month_folder, Prod_date, 
                Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                Electric_IACS, Weight, Elongation, Tensile,
                Bend_test, Spectro_Cu, Oxygen
            ) VALUES ");

            for (int i = 0; i < _busbarBatchBuffer.Count; i++)
            {
                if (i > 0) sqlBuilder.Append(",");
                sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @t{i}, @wd{i}, @l{i}, @r{i}, @c{i}, @e{i}, @wt{i}, @el{i}, @tn{i}, @bt{i}, @sp{i}, @ox{i})");

                var item = _busbarBatchBuffer[i];
                cmd.Parameters.AddWithValue($"@s{i}", item.Size);
                cmd.Parameters.AddWithValue($"@y{i}", item.Year.Trim());
                cmd.Parameters.AddWithValue($"@m{i}", item.Month.Trim());
                cmd.Parameters.AddWithValue($"@d{i}", item.ProdDate);
                cmd.Parameters.AddWithValue($"@t{i}", item.Thickness);
                cmd.Parameters.AddWithValue($"@wd{i}", item.Width);
                cmd.Parameters.AddWithValue($"@l{i}", item.Length);
                cmd.Parameters.AddWithValue($"@r{i}", item.Radius);
                cmd.Parameters.AddWithValue($"@c{i}", item.Chamber);
                cmd.Parameters.AddWithValue($"@wt{i}", item.Resistivity);
                cmd.Parameters.AddWithValue($"@e{i}", item.Electric);
                cmd.Parameters.AddWithValue($"@el{i}", item.Elongation);
                cmd.Parameters.AddWithValue($"@tn{i}", item.Tensile);
                cmd.Parameters.AddWithValue($"@bt{i}", string.IsNullOrEmpty(item.BendTest) ? (object)System.DBNull.Value : item.BendTest);
                cmd.Parameters.AddWithValue($"@sp{i}", item.Spectro);
                cmd.Parameters.AddWithValue($"@ox{i}", item.Oxygen);
            }

            cmd.CommandText = sqlBuilder.ToString();
            cmd.ExecuteNonQuery();
            _busbarBatchBuffer.Clear();
        }

        private void FlushTLJBatch(
           Microsoft.Data.Sqlite.SqliteConnection connection,
           Microsoft.Data.Sqlite.SqliteTransaction transaction,
           string tableName,
           System.Collections.Generic.List<TLJRecord> buffer)
        {
            if (buffer.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder(buffer.Count * 150);
            sqlBuilder.Append($"INSERT INTO {tableName} (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no) VALUES ");

            for (int i = 0; i < buffer.Count; i++)
            {
                if (i > 0) sqlBuilder.Append(",");
                sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @b{i})");

                var item = buffer[i];
                cmd.Parameters.AddWithValue($"@s{i}", item.Size);
                cmd.Parameters.AddWithValue($"@y{i}", item.Year.Trim());
                cmd.Parameters.AddWithValue($"@m{i}", item.Month.Trim());
                cmd.Parameters.AddWithValue($"@d{i}", item.ProdDate);
                cmd.Parameters.AddWithValue($"@b{i}", string.IsNullOrEmpty(item.BatchNo) ? (object)System.DBNull.Value : item.BatchNo);
            }

            cmd.CommandText = sqlBuilder.ToString();
            cmd.ExecuteNonQuery();
            buffer.Clear();
        }

        public void UpdateBusbarBatchNumbers(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            try
            {
                var cache350 = LoadTLJCache(connection, transaction, "TLJ350");
                var cache500 = LoadTLJCache(connection, transaction, "TLJ500");

                using var selectBusbarCmd = connection.CreateCommand();
                selectBusbarCmd.Transaction = transaction;
                selectBusbarCmd.CommandText = @"
                    SELECT Id, Size_mm, Prod_date 
                    FROM Busbar 
                    WHERE (Batch_no IS NULL OR Batch_no = '')
                    ORDER BY Prod_date ASC
                ";

                var updateBatch = new System.Collections.Generic.List<(int Id, string Batch)>(5000);
                var usedBatchNumbers = new System.Collections.Generic.HashSet<string>();

                using (var reader = selectBusbarCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        int id = reader.GetInt32(0);
                        string size = reader.GetString(1);
                        string dateStr = reader.GetString(2);

                        string targetTable = StringHelper.DetermineTLJTable(size);
                        var targetCache = (targetTable == "TLJ350") ? cache350 : cache500;

                        string batchNo = FindClosestAvailableBatch(targetCache, size, dateStr, usedBatchNumbers);

                        if (!string.IsNullOrEmpty(batchNo))
                        {
                            updateBatch.Add((id, batchNo));
                            usedBatchNumbers.Add($"{size}|{batchNo}");
                        }

                        if (updateBatch.Count >= 5000)
                        {
                            ExecuteBulkUpdate(connection, transaction, updateBatch);
                            updateBatch.Clear();
                        }
                    }
                }

                if (updateBatch.Count > 0)
                {
                    ExecuteBulkUpdate(connection, transaction, updateBatch);
                }
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        private void ExecuteBulkUpdate(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction transaction,
            System.Collections.Generic.List<(int Id, string Batch)> updates)
        {
            if (updates.Count == 0) return;

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = "CREATE TEMP TABLE IF NOT EXISTS TempBusbarUpdates (Id INTEGER PRIMARY KEY, Batch_no TEXT)";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM TempBusbarUpdates";
            cmd.ExecuteNonQuery();

            var sqlBuilder = new System.Text.StringBuilder();
            sqlBuilder.Append("INSERT INTO TempBusbarUpdates (Id, Batch_no) VALUES ");

            for (int i = 0; i < updates.Count; i++)
            {
                if (i > 0) sqlBuilder.Append(",");
                sqlBuilder.Append($"(@id{i}, @b{i})");
                cmd.Parameters.AddWithValue($"@id{i}", updates[i].Id);
                cmd.Parameters.AddWithValue($"@b{i}", updates[i].Batch);
            }

            cmd.CommandText = sqlBuilder.ToString();
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                UPDATE Busbar
                SET Batch_no = (SELECT Batch_no FROM TempBusbarUpdates WHERE TempBusbarUpdates.Id = Busbar.Id)
                WHERE Id IN (SELECT Id FROM TempBusbarUpdates);
            ";
            cmd.Parameters.Clear();
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DROP TABLE TempBusbarUpdates";
            cmd.ExecuteNonQuery();
        }

        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>> LoadTLJCache(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction trans,
            string tableName)
        {
            var cache = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>>(500);

            using var cmd = connection.CreateCommand();
            cmd.Transaction = trans;
            cmd.CommandText = $"SELECT Size_mm, Prod_date, Batch_no FROM {tableName} WHERE Batch_no IS NOT NULL AND Batch_no != ''";

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                string size = reader.GetString(0);
                string dateStr = reader.GetString(1);
                string batch = reader.GetString(2);

                if (!System.DateTime.TryParseExact(dateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime dt))
                {
                    continue;
                }

                if (!cache.ContainsKey(size))
                {
                    cache[size] = new System.Collections.Generic.List<TLJRecord>(100);
                }

                cache[size].Add(new TLJRecord
                {
                    Size = size,
                    ProdDate = dateStr,
                    ParsedDate = dt,
                    BatchNo = batch
                });
            }

            foreach (var key in cache.Keys)
            {
                cache[key].Sort((a, b) => a.ParsedDate.CompareTo(b.ParsedDate));
            }

            return cache;
        }

        private string FindClosestAvailableBatch(
            System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<TLJRecord>> cache,
            string size,
            string targetDateStr,
            System.Collections.Generic.HashSet<string> usedBatchNumbers)
        {
            if (!cache.ContainsKey(size)) return string.Empty;

            if (!System.DateTime.TryParseExact(targetDateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime targetDate))
            {
                return string.Empty;
            }

            var list = cache[size];

            TLJRecord? bestCandidate = null;
            System.TimeSpan smallestGap = System.TimeSpan.MaxValue;

            foreach (var rec in list)
            {
                if (rec.ParsedDate > targetDate)
                {
                    break;
                }

                string batchKey = $"{size}|{rec.BatchNo}";
                if (usedBatchNumbers.Contains(batchKey))
                {
                    continue;
                }

                System.TimeSpan gap = targetDate - rec.ParsedDate;

                if (gap < smallestGap)
                {
                    smallestGap = gap;
                    bestCandidate = rec;
                }
            }

            if (bestCandidate.HasValue)
            {
                return StringHelper.ProcessRawBatchString(bestCandidate.Value.BatchNo);
            }

            return string.Empty;
        }

        public System.Collections.Generic.List<BusbarSearchItem> SearchBusbarRecords(string year, string month, string prodDate)
        {
            var results = new System.Collections.Generic.List<BusbarSearchItem>();

            using var conn = _dbContext.CreateConnection();

            using var command = conn.CreateCommand();

            command.CommandText = @"
                SELECT * FROM Busbar 
                WHERE Year_folder = @year 
                  AND Month_folder = @month 
                  AND Prod_date = @prodDate";

            command.Parameters.AddWithValue("@year", year);
            command.Parameters.AddWithValue("@month", month);
            command.Parameters.AddWithValue("@prodDate", prodDate);

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                BusbarRecord fullRecord = new BusbarRecord();

                fullRecord.Id = reader.GetInt32(reader.GetOrdinal("Id"));

                fullRecord.Size = reader["Size_mm"] as string ?? string.Empty;
                fullRecord.Year = reader["Year_folder"] as string ?? string.Empty;
                fullRecord.Month = reader["Month_folder"] as string ?? string.Empty;
                fullRecord.ProdDate = reader["Prod_date"] as string ?? string.Empty;
                fullRecord.BendTest = reader["Bend_test"] as string ?? string.Empty;
                fullRecord.BatchNo = reader["Batch_no"] as string ?? string.Empty;

                fullRecord.Thickness = ParseDoubleSafe(reader["Thickness_mm"]);
                fullRecord.Width = ParseDoubleSafe(reader["Width_mm"]);
                fullRecord.Length = ParseIntSafe(reader["Length"]);
                fullRecord.Radius = ParseDoubleSafe(reader["Radius"]);
                fullRecord.Chamber = ParseDoubleSafe(reader["Chamber_mm"]);
                fullRecord.Electric = ParseDoubleSafe(reader["Electric_IACS"]);
                fullRecord.Resistivity = ParseDoubleSafe(reader["Weight"]);
                fullRecord.Elongation = ParseDoubleSafe(reader["Elongation"]);
                fullRecord.Tensile = ParseDoubleSafe(reader["Tensile"]);
                fullRecord.Spectro = ParseDoubleSafe(reader["Spectro_Cu"]);
                fullRecord.Oxygen = ParseDoubleSafe(reader["Oxygen"]);

                // FILTER DINONAKTIFKAN (DI-KOMENTARI) SESUAI REQUEST
                /*
                if (!IsRecordComplete(fullRecord))
                {
                    continue;
                }

                if (!fullRecord.BendTest.Equals("OK", System.StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                */

                results.Add(new BusbarSearchItem
                {
                    No = fullRecord.Id,
                    Specification = fullRecord.Size,
                    DateProd = fullRecord.ProdDate,
                    FullRecord = fullRecord
                });
            }

            return results;
        }

        private bool IsRecordComplete(BusbarRecord record)
        {
            if (string.IsNullOrWhiteSpace(record.BatchNo)) return false;
            if (string.IsNullOrWhiteSpace(record.Size)) return false;
            if (string.IsNullOrWhiteSpace(record.BendTest)) return false;
            if (string.IsNullOrWhiteSpace(record.Year)) return false;
            if (string.IsNullOrWhiteSpace(record.Month)) return false;
            if (string.IsNullOrWhiteSpace(record.ProdDate)) return false;

            if (record.Length <= 0) return false;
            if (record.Thickness <= 0) return false;
            if (record.Width <= 0) return false;
            if (record.Radius <= 0) return false;
            if (record.Chamber <= 0) return false;
            if (record.Electric <= 0) return false;
            if (record.Resistivity <= 0) return false;
            if (record.Elongation <= 0) return false;
            if (record.Tensile <= 0) return false;
            if (record.Spectro <= 0) return false;
            if (record.Oxygen <= 0) return false;

            return true;
        }

        private double ParseDoubleSafe(object value)
        {
            if (value == null || value == System.DBNull.Value) return 0.0;
            if (value is double d) return d;
            if (value is float f) return (double)f;
            if (value is int i) return (double)i;
            if (value is long l) return (double)l;
            if (value is string s && double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }

        private int ParseIntSafe(object value)
        {
            if (value == null || value == System.DBNull.Value) return 0;
            if (value is long l) return (int)l;
            if (value is int i) return i;
            if (value is string s && int.TryParse(s, out int result)) return result;
            if (value is double d) return (int)d;
            if (value is float f) return (int)f;
            return 0;
        }

        public System.Collections.Generic.List<string> GetAvailableYears()
        {
            var years = new System.Collections.Generic.List<string>();

            try
            {
                using var conn = _dbContext.CreateConnection();
                using var command = conn.CreateCommand();

                command.CommandText = "SELECT DISTINCT Year_folder FROM Busbar ORDER BY Year_folder DESC";

                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    years.Add(reader.GetString(0));
                }
            }
            catch (System.Exception)
            {
            }

            return years;
        }
    }
}