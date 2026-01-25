namespace WpfApp1.Data.Repositories
{
    public class BusbarRepository
    {
        private const System.Int32 BATCH_SIZE = 5000;

        private System.Collections.Generic.List<WpfApp1.Core.Models.BusbarRecord> _busbarBatchBuffer =
            new System.Collections.Generic.List<WpfApp1.Core.Models.BusbarRecord>(BATCH_SIZE);
        private System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord> _tlj350BatchBuffer =
            new System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>(BATCH_SIZE);
        private System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord> _tlj500BatchBuffer =
            new System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>(BATCH_SIZE);

        private readonly WpfApp1.Data.Database.SqliteContext _dbContext;

        public BusbarRepository(WpfApp1.Data.Database.SqliteContext dbContext)
        {
            _dbContext = dbContext;
        }

        public void CreateBusbarTable(Microsoft.Data.Sqlite.SqliteConnection connection)
        {
            _busbarBatchBuffer.Clear();
            _tlj350BatchBuffer.Clear();
            _tlj500BatchBuffer.Clear();

            Microsoft.Data.Sqlite.SqliteCommand? cmd = null;
            try
            {
                cmd = connection.CreateCommand();

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
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
        }

        public void InsertBusbarRow(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, WpfApp1.Core.Models.BusbarRecord record)
        {
            _busbarBatchBuffer.Add(record);
            if (_busbarBatchBuffer.Count >= BATCH_SIZE)
            {
                FlushBusbarBatch(connection, transaction);
            }
        }

        public void InsertTLJ350_Row(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, WpfApp1.Core.Models.TLJRecord record)
        {
            _tlj350BatchBuffer.Add(record);
            if (_tlj350BatchBuffer.Count >= BATCH_SIZE)
            {
                FlushTLJBatch(connection, transaction, "TLJ350", _tlj350BatchBuffer);
            }
        }

        public void InsertTLJ500_Row(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction, WpfApp1.Core.Models.TLJRecord record)
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

            Microsoft.Data.Sqlite.SqliteCommand? cmd = null;
            try
            {
                cmd = connection.CreateCommand();
                cmd.Transaction = transaction;

                System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder(BATCH_SIZE * 250);
                sqlBuilder.Append(@"INSERT INTO Busbar (
                Size_mm, Year_folder, Month_folder, Prod_date, 
                Thickness_mm, Width_mm, Length, Radius, Chamber_mm,
                Electric_IACS, Weight, Elongation, Tensile,
                Bend_test, Spectro_Cu, Oxygen
            ) VALUES ");

                for (System.Int32 i = 0; i < _busbarBatchBuffer.Count; i++)
                {
                    if (i > 0) sqlBuilder.Append(",");
                    sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @t{i}, @wd{i}, @l{i}, @r{i}, @c{i}, @e{i}, @wt{i}, @el{i}, @tn{i}, @bt{i}, @sp{i}, @ox{i})");

                    WpfApp1.Core.Models.BusbarRecord item = _busbarBatchBuffer[i];
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
                    cmd.Parameters.AddWithValue($"@bt{i}", System.String.IsNullOrEmpty(item.BendTest) ? (System.Object)System.DBNull.Value : item.BendTest);
                    cmd.Parameters.AddWithValue($"@sp{i}", item.Spectro);
                    cmd.Parameters.AddWithValue($"@ox{i}", item.Oxygen);
                }

                cmd.CommandText = sqlBuilder.ToString();
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                _busbarBatchBuffer.Clear();
            }
        }

        private void FlushTLJBatch(
           Microsoft.Data.Sqlite.SqliteConnection connection,
           Microsoft.Data.Sqlite.SqliteTransaction transaction,
           System.String tableName,
           System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord> buffer)
        {
            if (buffer.Count == 0) return;

            Microsoft.Data.Sqlite.SqliteCommand? cmd = null;
            try
            {
                cmd = connection.CreateCommand();
                cmd.Transaction = transaction;

                System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder(buffer.Count * 150);
                sqlBuilder.Append($"INSERT INTO {tableName} (Size_mm, Year_folder, Month_folder, Prod_date, Batch_no) VALUES ");

                for (System.Int32 i = 0; i < buffer.Count; i++)
                {
                    if (i > 0) sqlBuilder.Append(",");
                    sqlBuilder.Append($"(@s{i}, @y{i}, @m{i}, @d{i}, @b{i})");

                    WpfApp1.Core.Models.TLJRecord item = buffer[i];
                    cmd.Parameters.AddWithValue($"@s{i}", item.Size);
                    cmd.Parameters.AddWithValue($"@y{i}", item.Year.Trim());
                    cmd.Parameters.AddWithValue($"@m{i}", item.Month.Trim());
                    cmd.Parameters.AddWithValue($"@d{i}", item.ProdDate);
                    cmd.Parameters.AddWithValue($"@b{i}", System.String.IsNullOrEmpty(item.BatchNo) ? (System.Object)System.DBNull.Value : item.BatchNo);
                }

                cmd.CommandText = sqlBuilder.ToString();
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
                buffer.Clear();
            }
        }

        public void UpdateBusbarBatchNumbers(Microsoft.Data.Sqlite.SqliteConnection connection, Microsoft.Data.Sqlite.SqliteTransaction transaction)
        {
            try
            {
                System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> cache350 = LoadTLJCache(connection, transaction, "TLJ350");
                System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> cache500 = LoadTLJCache(connection, transaction, "TLJ500");

                Microsoft.Data.Sqlite.SqliteCommand? selectBusbarCmd = null;
                Microsoft.Data.Sqlite.SqliteDataReader? reader = null;

                try
                {
                    selectBusbarCmd = connection.CreateCommand();
                    selectBusbarCmd.Transaction = transaction;
                    selectBusbarCmd.CommandText = @"
                    SELECT Id, Size_mm, Prod_date 
                    FROM Busbar 
                    WHERE (Batch_no IS NULL OR Batch_no = '')
                    ORDER BY Prod_date ASC
                ";

                    System.Collections.Generic.List<(System.Int32 Id, System.String Batch)> updateBatch =
                        new System.Collections.Generic.List<(System.Int32, System.String)>(5000);
                    System.Collections.Generic.HashSet<System.String> usedBatchNumbers =
                        new System.Collections.Generic.HashSet<System.String>();

                    reader = selectBusbarCmd.ExecuteReader();
                    while (reader.Read())
                    {
                        System.Int32 id = reader.GetInt32(0);
                        System.String size = reader.GetString(1);
                        System.String dateStr = reader.GetString(2);

                        System.String targetTable = WpfApp1.Shared.Helpers.StringHelper.DetermineTLJTable(size);
                        System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> targetCache =
                            (targetTable == "TLJ350") ? cache350 : cache500;

                        System.String batchNo = FindClosestAvailableBatch(targetCache, size, dateStr, usedBatchNumbers);

                        if (!System.String.IsNullOrEmpty(batchNo))
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

                    if (updateBatch.Count > 0)
                    {
                        ExecuteBulkUpdate(connection, transaction, updateBatch);
                    }
                }
                finally
                {
                    if (reader != null) reader.Dispose();
                    if (selectBusbarCmd != null) selectBusbarCmd.Dispose();
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
            System.Collections.Generic.List<(System.Int32 Id, System.String Batch)> updates)
        {
            if (updates.Count == 0) return;

            Microsoft.Data.Sqlite.SqliteCommand? cmd = null;
            try
            {
                cmd = connection.CreateCommand();
                cmd.Transaction = transaction;

                cmd.CommandText = "CREATE TEMP TABLE IF NOT EXISTS TempBusbarUpdates (Id INTEGER PRIMARY KEY, Batch_no TEXT)";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "DELETE FROM TempBusbarUpdates";
                cmd.ExecuteNonQuery();

                System.Text.StringBuilder sqlBuilder = new System.Text.StringBuilder();
                sqlBuilder.Append("INSERT INTO TempBusbarUpdates (Id, Batch_no) VALUES ");

                for (System.Int32 i = 0; i < updates.Count; i++)
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
            finally
            {
                if (cmd != null)
                {
                    cmd.Dispose();
                }
            }
        }

        private System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> LoadTLJCache(
            Microsoft.Data.Sqlite.SqliteConnection connection,
            Microsoft.Data.Sqlite.SqliteTransaction trans,
            System.String tableName)
        {
            System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> cache =
                new System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>>(500);

            Microsoft.Data.Sqlite.SqliteCommand? cmd = null;
            Microsoft.Data.Sqlite.SqliteDataReader? reader = null;

            try
            {
                cmd = connection.CreateCommand();
                cmd.Transaction = trans;
                cmd.CommandText = $"SELECT Size_mm, Prod_date, Batch_no FROM {tableName} WHERE Batch_no IS NOT NULL AND Batch_no != ''";

                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    System.String size = reader.GetString(0);
                    System.String dateStr = reader.GetString(1);
                    System.String batch = reader.GetString(2);

                    if (!System.DateTime.TryParseExact(dateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime dt))
                    {
                        continue;
                    }

                    if (!cache.ContainsKey(size))
                    {
                        cache[size] = new System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>(100);
                    }

                    cache[size].Add(new WpfApp1.Core.Models.TLJRecord
                    {
                        Size = size,
                        ProdDate = dateStr,
                        ParsedDate = dt,
                        BatchNo = batch
                    });
                }
            }
            finally
            {
                if (reader != null) reader.Dispose();
                if (cmd != null) cmd.Dispose();
            }

            foreach (System.String key in cache.Keys)
            {
                cache[key].Sort((a, b) => a.ParsedDate.CompareTo(b.ParsedDate));
            }

            return cache;
        }

        private System.String FindClosestAvailableBatch(
            System.Collections.Generic.Dictionary<System.String, System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord>> cache,
            System.String size,
            System.String targetDateStr,
            System.Collections.Generic.HashSet<System.String> usedBatchNumbers)
        {
            if (!cache.ContainsKey(size)) return System.String.Empty;

            if (!System.DateTime.TryParseExact(targetDateStr, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out System.DateTime targetDate))
            {
                return System.String.Empty;
            }

            System.Collections.Generic.List<WpfApp1.Core.Models.TLJRecord> list = cache[size];

            WpfApp1.Core.Models.TLJRecord? bestCandidate = null;
            System.TimeSpan smallestGap = System.TimeSpan.MaxValue;

            foreach (WpfApp1.Core.Models.TLJRecord rec in list)
            {
                if (rec.ParsedDate > targetDate)
                {
                    break;
                }

                System.String batchKey = $"{size}|{rec.BatchNo}";
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
                return WpfApp1.Shared.Helpers.StringHelper.ProcessRawBatchString(bestCandidate.Value.BatchNo);
            }

            return System.String.Empty;
        }

        public System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem> SearchBusbarRecords(System.String year, System.String month, System.String prodDate)
        {
            System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem> results = new System.Collections.Generic.List<WpfApp1.Core.Models.BusbarSearchItem>();

            Microsoft.Data.Sqlite.SqliteConnection? conn = null;
            Microsoft.Data.Sqlite.SqliteCommand? command = null;
            Microsoft.Data.Sqlite.SqliteDataReader? reader = null;

            try
            {
                conn = _dbContext.CreateConnection();
                command = conn.CreateCommand();

                command.CommandText = @"
                SELECT * FROM Busbar 
                WHERE Year_folder = @year 
                  AND Month_folder = @month 
                  AND Prod_date = @prodDate";

                command.Parameters.AddWithValue("@year", year);
                command.Parameters.AddWithValue("@month", month);
                command.Parameters.AddWithValue("@prodDate", prodDate);

                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    WpfApp1.Core.Models.BusbarRecord fullRecord = new WpfApp1.Core.Models.BusbarRecord();

                    fullRecord.Id = reader.GetInt32(reader.GetOrdinal("Id"));

                    fullRecord.Size = reader["Size_mm"] as System.String ?? System.String.Empty;
                    fullRecord.Year = reader["Year_folder"] as System.String ?? System.String.Empty;
                    fullRecord.Month = reader["Month_folder"] as System.String ?? System.String.Empty;
                    fullRecord.ProdDate = reader["Prod_date"] as System.String ?? System.String.Empty;
                    fullRecord.BendTest = reader["Bend_test"] as System.String ?? System.String.Empty;
                    fullRecord.BatchNo = reader["Batch_no"] as System.String ?? System.String.Empty;

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

                    results.Add(new WpfApp1.Core.Models.BusbarSearchItem
                    {
                        No = fullRecord.Id,
                        Specification = fullRecord.Size,
                        DateProd = fullRecord.ProdDate,
                        FullRecord = fullRecord
                    });
                }
            }
            finally
            {
                if (reader != null) reader.Dispose();
                if (command != null) command.Dispose();
                if (conn != null) conn.Dispose();
            }

            return results;
        }

        private System.Double ParseDoubleSafe(System.Object value)
        {
            if (value == null || value == System.DBNull.Value) return 0.0;
            if (value is System.Double d) return d;
            if (value is System.Single f) return (System.Double)f;
            if (value is System.Int32 i) return (System.Double)i;
            if (value is System.Int64 l) return (System.Double)l;
            if (value is System.String s && System.Double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }

        private System.Int32 ParseIntSafe(System.Object value)
        {
            if (value == null || value == System.DBNull.Value) return 0;
            if (value is System.Int64 l) return (System.Int32)l;
            if (value is System.Int32 i) return i;
            if (value is System.String s && System.Int32.TryParse(s, out int result)) return result;
            if (value is System.Double d) return (System.Int32)d;
            if (value is System.Single f) return (System.Int32)f;
            return 0;
        }

        public System.Collections.Generic.List<System.String> GetAvailableYears()
        {
            System.Collections.Generic.List<System.String> years = new System.Collections.Generic.List<System.String>();

            Microsoft.Data.Sqlite.SqliteConnection? conn = null;
            Microsoft.Data.Sqlite.SqliteCommand? command = null;
            Microsoft.Data.Sqlite.SqliteDataReader? reader = null;

            try
            {
                conn = _dbContext.CreateConnection();
                command = conn.CreateCommand();

                command.CommandText = "SELECT DISTINCT Year_folder FROM Busbar ORDER BY Year_folder DESC";

                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    years.Add(reader.GetString(0));
                }
            }
            catch (System.Exception)
            {
            }
            finally
            {
                if (reader != null) reader.Dispose();
                if (command != null) command.Dispose();
                if (conn != null) conn.Dispose();
            }

            return years;
        }
    }
}