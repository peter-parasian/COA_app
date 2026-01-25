namespace WpfApp1.Data.Database
{
    public class SqliteContext
    {
        private readonly System.String _dbPath;

        public SqliteContext()
        {
            System.String localAppData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);

            System.String appFolder = System.IO.Path.Combine(localAppData, "Database_QC");

            _dbPath = System.IO.Path.Combine(appFolder, "data_qc.db");
        }
        public void EnsureDatabaseFolderExists()
        {
            System.String? folder = System.IO.Path.GetDirectoryName(_dbPath);

            if (!System.String.IsNullOrEmpty(folder) && !System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }
        }

        public Microsoft.Data.Sqlite.SqliteConnection CreateConnection()
        {
            EnsureDatabaseFolderExists();
            Microsoft.Data.Sqlite.SqliteConnection conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={_dbPath}");
            conn.Open();

            Microsoft.Data.Sqlite.SqliteCommand? pragmaCmd = null;
            try
            {
                pragmaCmd = conn.CreateCommand();
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
            finally
            {
                if (pragmaCmd != null)
                {
                    pragmaCmd.Dispose();
                }
            }
            return conn;
        }

    }
}