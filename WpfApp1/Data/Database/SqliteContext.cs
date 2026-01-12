using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WpfApp1.Data.Database
{
    public class SqliteContext
    {
        private readonly string _dbPath;

        public SqliteContext()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            string appFolder = Path.Combine(localAppData, "Database_QC");

            _dbPath = Path.Combine(appFolder, "data_qc.db");
        }
        public void EnsureDatabaseFolderExists()
        {
            string? folder = Path.GetDirectoryName(_dbPath);

            if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        public Microsoft.Data.Sqlite.SqliteConnection CreateConnection()
        {
            EnsureDatabaseFolderExists();
            var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={_dbPath}");
            conn.Open();

            using (var pragmaCmd = conn.CreateCommand())
            {
                pragmaCmd.CommandText = @"
            PRAGMA synchronous = OFF;
            PRAGMA journal_mode = WAL;
            PRAGMA temp_store = MEMORY;
            PRAGMA locking_mode = EXCLUSIVE;
            PRAGMA cache_size = -128000;
            PRAGMA page_size = 8192;
            PRAGMA mmap_size = 30000000000;
            PRAGMA automatic_index = OFF;
            PRAGMA query_only = OFF;
        ";
                pragmaCmd.ExecuteNonQuery();
            }
            return conn;
        }

    }
}