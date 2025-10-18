using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using PERI.SK.Domain.Interfaces;
using System.Text.Json;

namespace PERI.SK.Infrastructure.Data
{
    public class ExcelQueries : IDataQueries
    {
        readonly SqliteConnection _connection;

        public ExcelQueries()
        {
            _connection = new SqliteConnection("Data Source=:memory:");
            _connection.Open();
        }

        public async Task<bool> CanConnect(string connectionString)
        {
            var path = Path.GetFullPath(connectionString);

            return File.Exists(path);
        }

        public async Task<string> GetData(string connectionString, string? query = null)
        {
            using var command = _connection.CreateCommand();
            command.CommandText = query;

            using var reader = await command.ExecuteReaderAsync();

            var results = new List<Dictionary<string, object>>();

            while (await reader.ReadAsync())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[reader.GetName(i)] = await reader.IsDBNullAsync(i) ? null : reader.GetValue(i);
                }
                results.Add(row);
            }

            return JsonSerializer.Serialize(results);
        }

        public async Task<string> GetFields(string connectionString, string? schema = null, string? table = null)
        {
            return await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(connectionString);
                var worksheet = workbook.Worksheet(1);  // first worksheet

                var firstRow = worksheet.Row(1);
                int lastColumn = firstRow.LastCellUsed().Address.ColumnNumber;

                var headers = new string[lastColumn];
                for (int col = 1; col <= lastColumn; col++)
                {
                    headers[col - 1] = worksheet.Cell(1, col).GetString();
                }

                return string.Join(", ", headers);
            });
        }

        /// <summary>
        /// Creates SQLite db and table
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="tableName"></param>
        public void CreateInMemoryData(string excelPath, string tableName)
        {
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(1);

            // Get column names from first row
            var firstRow = worksheet.Row(1);
            var columns = new List<string>();
            foreach (var cell in firstRow.CellsUsed())
                columns.Add(cell.GetString());

            // Helper to escape SQLite identifiers safely
            string EscapeSqlIdentifier(string name) => $"[{name.Replace("]", "]]")}]";
            // Check if table exists
            using (var checkCmd = _connection.CreateCommand())
            {
                checkCmd.CommandText = $"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName.Replace("'", "''")}'";
                using var reader = checkCmd.ExecuteReader();
                if (reader.Read())
                {
                    // Table already exists, return early
                    return;
                }
            }

            // Create table SQL (all TEXT), with safe column names
            string createTableSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", columns.ConvertAll(c => $"{EscapeSqlIdentifier(c)} TEXT"))});";

            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = createTableSql;
                cmd.ExecuteNonQuery();
            }

            int lastRow = worksheet.LastRowUsed().RowNumber();

            for (int rowNum = 2; rowNum <= lastRow; rowNum++)
            {
                var row = worksheet.Row(rowNum);
                var values = new List<string>();

                for (int colIndex = 1; colIndex <= columns.Count; colIndex++)
                {
                    var val = row.Cell(colIndex).GetString().Replace("'", "''");
                    values.Add($"'{val}'");
                }

                string insertSql = $"INSERT INTO [{tableName}] ({string.Join(", ", columns.ConvertAll(c => EscapeSqlIdentifier(c)))}) VALUES ({string.Join(", ", values)});";

                using var insertCmd = _connection.CreateCommand();
                insertCmd.CommandText = insertSql;
                insertCmd.ExecuteNonQuery();
            }

            Console.WriteLine($"Inserted {lastRow - 1} rows into table '{tableName}'.");
        }
    }
}
