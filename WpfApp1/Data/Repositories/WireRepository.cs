using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Data.Sqlite;
using WpfApp1.Core.Models;
using WpfApp1.Data.Database;

namespace WpfApp1.Data.Repositories
{
    public class WireRepository
    {
        private const int BATCH_SIZE = 100;

        private readonly SqliteContext _dbContext;

        private List<WireRecord> _currentBuffer = new List<WireRecord>(BATCH_SIZE);
        private List<WireRecord> _exportBuffer = new List<WireRecord>(BATCH_SIZE);

        public WireRepository(SqliteContext dbContext)
        {
            if (dbContext == null)
            {
                throw new ArgumentNullException(nameof(dbContext));
            }

            _dbContext = dbContext;
        }

        public void CreateWireTables(SqliteConnection connection)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            _currentBuffer.Clear();
            _exportBuffer.Clear();

            using var cmd = connection.CreateCommand();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Wire_Current;
                CREATE TABLE IF NOT EXISTS Wire_Current (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Size TEXT,
                    Date TEXT,
                    Lot TEXT,
                    Customer TEXT,
                    Diameter REAL,
                    Yield REAL,
                    Tensile REAL,
                    Elongation REAL,
                    IACS REAL
                );
            ";
            cmd.ExecuteNonQuery();

            cmd.CommandText = @"
                DROP TABLE IF EXISTS Wire_Export;
                CREATE TABLE IF NOT EXISTS Wire_Export (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Size TEXT,
                    Date TEXT,
                    Lot TEXT,
                    Customer TEXT,
                    Diameter REAL,
                    Yield REAL,
                    Tensile REAL,
                    Elongation REAL,
                    IACS REAL,
                    TimeExport TEXT
                );
                CREATE UNIQUE INDEX IF NOT EXISTS IDX_Unique_Export ON Wire_Export(Size, Date, Lot, Customer);
            ";
            cmd.ExecuteNonQuery();
        }

        public void ClearCurrentTable(SqliteConnection connection, SqliteTransaction transaction)
        {
            if (connection == null)
            {
                throw new ArgumentNullException(nameof(connection));
            }

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = "DELETE FROM Wire_Current;";
            cmd.ExecuteNonQuery();
        }

        public void InsertIntoCurrent(SqliteConnection connection, SqliteTransaction transaction, WireRecord record)
        {
            _currentBuffer.Add(record);
            if (_currentBuffer.Count >= BATCH_SIZE)
            {
                FlushCurrentBatch(connection, transaction);
            }
        }

        public void InsertIntoExport(SqliteConnection connection, SqliteTransaction transaction, WireRecord record)
        {
            _exportBuffer.Add(record);
            if (_exportBuffer.Count >= BATCH_SIZE)
            {
                FlushExportBatch(connection, transaction);
            }
        }

        public void FlushAll(SqliteConnection connection, SqliteTransaction transaction)
        {
            FlushCurrentBatch(connection, transaction);
            FlushExportBatch(connection, transaction);
        }

        private void FlushCurrentBatch(SqliteConnection connection, SqliteTransaction transaction)
        {
            if (_currentBuffer.Count == 0)
            {
                return;
            }

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            StringBuilder sqlBuilder = new StringBuilder(_currentBuffer.Count * 100);
            sqlBuilder.Append(@"INSERT INTO Wire_Current (
                Size, Date, Lot, Customer, Diameter, Yield, Tensile, Elongation, IACS
            ) VALUES ");

            for (int i = 0; i < _currentBuffer.Count; i++)
            {
                if (i > 0)
                {
                    sqlBuilder.Append(",");
                }
                sqlBuilder.Append($"(@size{i}, @date{i}, @lot{i}, @cust{i}, @dia{i}, @yield{i}, @tens{i}, @elong{i}, @iacs{i})");

                WireRecord item = _currentBuffer[i];
                cmd.Parameters.AddWithValue($"@size{i}", item.Size ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@date{i}", item.Date ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@lot{i}", item.Lot ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@cust{i}", item.Customer ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@dia{i}", item.Diameter);
                cmd.Parameters.AddWithValue($"@yield{i}", item.Yield);
                cmd.Parameters.AddWithValue($"@tens{i}", item.Tensile);
                cmd.Parameters.AddWithValue($"@elong{i}", item.Elongation);
                cmd.Parameters.AddWithValue($"@iacs{i}", item.IACS);
            }

            cmd.CommandText = sqlBuilder.ToString();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (SqliteException ex)
            {
                if (ex.SqliteErrorCode != 19)
                {
                    throw;
                }
            }
            _currentBuffer.Clear();
        }

        private void FlushExportBatch(SqliteConnection connection, SqliteTransaction transaction)
        {
            if (_exportBuffer.Count == 0)
            {
                return;
            }

            using var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            StringBuilder sqlBuilder = new StringBuilder(_exportBuffer.Count * 120);
            sqlBuilder.Append(@"INSERT INTO Wire_Export (
                Size, Date, Lot, Customer, Diameter, Yield, Tensile, Elongation, IACS, TimeExport
            ) VALUES ");

            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            for (int i = 0; i < _exportBuffer.Count; i++)
            {
                if (i > 0)
                {
                    sqlBuilder.Append(",");
                }
                sqlBuilder.Append($"(@size{i}, @date{i}, @lot{i}, @cust{i}, @dia{i}, @yield{i}, @tens{i}, @elong{i}, @iacs{i}, @time{i})");

                WireRecord item = _exportBuffer[i];
                cmd.Parameters.AddWithValue($"@size{i}", item.Size ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@date{i}", item.Date ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@lot{i}", item.Lot ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@cust{i}", item.Customer ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue($"@dia{i}", item.Diameter);
                cmd.Parameters.AddWithValue($"@yield{i}", item.Yield);
                cmd.Parameters.AddWithValue($"@tens{i}", item.Tensile);
                cmd.Parameters.AddWithValue($"@elong{i}", item.Elongation);
                cmd.Parameters.AddWithValue($"@iacs{i}", item.IACS);
                cmd.Parameters.AddWithValue($"@time{i}", currentTime);
            }

            cmd.CommandText = sqlBuilder.ToString();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (SqliteException ex)
            {
                if (ex.SqliteErrorCode != 19)
                {
                    throw;
                }
            }
            _exportBuffer.Clear();
        }

        public List<WireSearchItem> SearchWireRecords(string size, string customer, string prodDate)
        {
            var results = new List<WireSearchItem>();

            using var conn = _dbContext.CreateConnection();

            using var command = conn.CreateCommand();

            command.CommandText = @"
                SELECT wc.*
                FROM Wire_Current wc
                WHERE wc.Size = @size 
                  AND wc.Customer = @customer 
                  AND wc.Date = @prodDate
                  AND NOT EXISTS (
                      SELECT 1 FROM Wire_Export we
                      WHERE we.Size = wc.Size
                        AND we.Date = wc.Date
                        AND we.Lot = wc.Lot
                        AND we.Customer = wc.Customer
                  )
                ORDER BY wc.Id ASC";

            command.Parameters.AddWithValue("@size", size);
            command.Parameters.AddWithValue("@customer", customer);
            command.Parameters.AddWithValue("@prodDate", prodDate);

            using var reader = command.ExecuteReader();
            int no = 1;
            while (reader.Read())
            {
                WireRecord fullRecord = new WireRecord
                {
                    Size = reader["Size"] as string ?? string.Empty,
                    Date = reader["Date"] as string ?? string.Empty,
                    Lot = reader["Lot"] as string ?? string.Empty,
                    Customer = reader["Customer"] as string ?? string.Empty,
                    Diameter = ParseDoubleSafe(reader["Diameter"]),
                    Yield = ParseDoubleSafe(reader["Yield"]),
                    Tensile = ParseDoubleSafe(reader["Tensile"]),
                    Elongation = ParseDoubleSafe(reader["Elongation"]),
                    IACS = ParseDoubleSafe(reader["IACS"])
                };

                results.Add(new WireSearchItem
                {
                    No = no++,
                    Specification = fullRecord.Size,
                    CustomerName = fullRecord.Customer,
                    DateProd = fullRecord.Date,
                    FullRecord = fullRecord
                });
            }

            return results;
        }

        private double ParseDoubleSafe(object value)
        {
            if (value == null || value == DBNull.Value) return 0.0;
            if (value is double d) return d;
            if (value is float f) return (double)f;
            if (value is int i) return (double)i;
            if (value is long l) return (double)l;
            if (value is string s && double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result)) return result;
            return 0.0;
        }
    }
}