using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private const string ExcelRootFolder = @"C:\Users\mrrx\Documents\My Web Sites\EXCEL";
        private const string DbPath = @"C:\Users\mrrx\data_qc.db";

        public MainWindow()
        {
            InitializeComponent();
            ImportExcelToSQLite();
        }

        private void ImportExcelToSQLite()
        {
            using var connection = new SqliteConnection($"Data Source={DbPath}");
            connection.Open();

            CreateTable(connection);

            var excelFiles = Directory
                .EnumerateFiles(ExcelRootFolder, "*.xlsx", SearchOption.AllDirectories)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"));

            using var transaction = connection.BeginTransaction();

            foreach (var file in excelFiles)
            {
                ParseSingleExcel(file, connection);
            }

            transaction.Commit();

            MessageBox.Show("Import Excel ke SQLite selesai");
        }

        private void CreateTable(SqliteConnection connection)
        {
            var cmd = connection.CreateCommand();
            cmd.CommandText =
            @"
            CREATE TABLE IF NOT EXISTS QC_Data (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                SourceFile TEXT,
                SheetName TEXT,
                BatchNumber TEXT,
                Tanggal TEXT,
                TS TEXT,
                BD TEXT,
                Cu TEXT,
                Electricity TEXT
            );
            ";
            cmd.ExecuteNonQuery();
        }

        private void ParseSingleExcel(string filePath, SqliteConnection connection)
        {
            using var workbook = new XLWorkbook(filePath);

            foreach (var sheet in workbook.Worksheets)
            {
                var rows = sheet.RowsUsed().ToList();
                if (rows.Count < 2)
                    continue;

                var header = rows[0]
                    .Cells()
                    .Select(c => c.GetString().Trim())
                    .ToList();

                foreach (var row in rows.Skip(1))
                {
                    InsertRow(connection, filePath, sheet.Name, header, row);
                }
            }
        }

        private void InsertRow(
            SqliteConnection connection,
            string sourceFile,
            string sheetName,
            List<string> header,
            IXLRow row)
        {
            string? GetValue(string colName)
            {
                int idx = header.IndexOf(colName);
                return idx >= 0 ? row.Cell(idx + 1).GetString() : null;
            }

            object DbValue(string? value)
            {
                return string.IsNullOrWhiteSpace(value)
                    ? DBNull.Value
                    : value;
            }

            var cmd = connection.CreateCommand();
            cmd.CommandText =
            @"
    INSERT INTO QC_Data
    (SourceFile, SheetName, BatchNumber, Tanggal, TS, BD, Cu, Electricity)
    VALUES
    ($file, $sheet, $batch, $tanggal, $ts, $bd, $cu, $elec);
    ";

            cmd.Parameters.AddWithValue("$file", DbValue(sourceFile));
            cmd.Parameters.AddWithValue("$sheet", DbValue(sheetName));
            cmd.Parameters.AddWithValue("$batch", DbValue(GetValue("Batch Number")));
            cmd.Parameters.AddWithValue("$tanggal", DbValue(GetValue("Tanggal")));
            cmd.Parameters.AddWithValue("$ts", DbValue(GetValue("TS")));
            cmd.Parameters.AddWithValue("$bd", DbValue(GetValue("BD")));
            cmd.Parameters.AddWithValue("$cu", DbValue(GetValue("Cu")));
            cmd.Parameters.AddWithValue("$elec", DbValue(GetValue("Electricity")));

            cmd.ExecuteNonQuery();
        }
    }
}