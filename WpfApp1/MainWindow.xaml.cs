using System.Windows;
using Microsoft.Data.Sqlite; 

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            SimpanDataKeSQLite();
        }

        private void SimpanDataKeSQLite()
        {
            string connectionString = @"Data Source=C:\Users\mrrx\data_aplikasi.db";
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();

                var createTableCmd = connection.CreateCommand();
                createTableCmd.CommandText =
                @"
            CREATE TABLE IF NOT EXISTS Pengguna (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Nama TEXT NOT NULL,
                Email TEXT
            );
        ";
                createTableCmd.ExecuteNonQuery();

                var insertCmd = connection.CreateCommand();
                insertCmd.CommandText = "INSERT INTO Pengguna (Nama, Email) VALUES ($nama, $email)";
                insertCmd.Parameters.AddWithValue("$nama", "Budi Santoso");
                insertCmd.Parameters.AddWithValue("$email", "budi@example.com");

                insertCmd.ExecuteNonQuery();

                MessageBox.Show("Tabel dipastikan ada dan data berhasil dimasukkan!");
            }
        }
    }
}