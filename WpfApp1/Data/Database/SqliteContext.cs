using System;
using System.Collections.Generic;
using System.Text;

namespace WpfApp1.Data.Database
{
    public class SqliteContext
    {
        private const string DbPath = @"C:\sqLite\data_qc.db";

        public void EnsureDatabaseFolderExists()
        {
            string? folder = System.IO.Path.GetDirectoryName(DbPath);
            if (!string.IsNullOrEmpty(folder) && !System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }
        }

        public Microsoft.Data.Sqlite.SqliteConnection CreateConnection()
        {
            var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={DbPath}");
            conn.Open();

            using (var pragmaCmd = conn.CreateCommand())
            {
                pragmaCmd.CommandText = @"
                    PRAGMA synchronous = OFF;
                    PRAGMA journal_mode = MEMORY;
                    PRAGMA temp_store = MEMORY;
                    PRAGMA locking_mode = EXCLUSIVE;
                    PRAGMA cache_size = -64000;
                    PRAGMA page_size = 4096;
                    PRAGMA mmap_size = 30000000000;
                ";
                pragmaCmd.ExecuteNonQuery();
            }
            return conn;
        }
    }
}