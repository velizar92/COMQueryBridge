using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace COMQueryBridge
{
    [ComVisible(true)]
    [Guid("C53A5A7D-7313-40C4-A628-0C0D4256EFD4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("COMQueryBridge.SQLQueryBridge")]
    public class SQLQueryBridge : ISQLQueryBridge
    {
        private SqlConnection _connection;
        private string _connectionString;

        public void Connect()
        {
            if (_connection != null && _connection.State == ConnectionState.Open)
            {
                return;
            }

            _connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            _connection = new SqlConnection(_connectionString);
            _connection.Open();
        }

        public void Disconnect()
        {
            if (_connection != null && _connection.State != ConnectionState.Closed)
            {
                _connection.Close();
                _connection.Dispose();
                _connection = null;
            }
        }

        public int ExecuteNonQuery(string sqlQuery)
        {
            CheckConnection();

            EnsureNotNullOrWhiteSpace(sqlQuery, nameof(sqlQuery));

            using (var cmd = new SqlCommand(sqlQuery, _connection))
            {
                return cmd.ExecuteNonQuery();
            }
        }

        public string ExecuteQueryAsJson(string sqlQuery)
        {
            CheckConnection();

            EnsureNotNullOrWhiteSpace(sqlQuery, nameof(sqlQuery));

            using (var cmd = new SqlCommand(sqlQuery, _connection))
            using (var reader = cmd.ExecuteReader())
            {
                var table = new DataTable();

                table.Load(reader);

                if (table.Rows.Count == 0)
                {
                    return "[]";
                }

                return JsonSerializer.Serialize(table);
            }
        }

        public string ExecuteScalar(string sqlQuery)
        {
            CheckConnection();

            EnsureNotNullOrWhiteSpace(sqlQuery, nameof(sqlQuery));

            using (var cmd = new SqlCommand(sqlQuery, _connection))
            {
                var result = cmd.ExecuteScalar();

                if (result == null || result == DBNull.Value)
                {
                    return null;
                }

                return result?.ToString();
            }
        }

        public void ExportJsonToCsv(string jsonData, string filePath)
        {
            if (string.IsNullOrWhiteSpace(jsonData))
            {
                throw new ArgumentException("JSON data cannot be null or empty.", nameof(jsonData));
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
            }

            DataTable table = JsonSerializer.Deserialize<DataTable>(jsonData)
                ?? throw new InvalidOperationException("Failed to deserialize JSON to DataTable.");

            using (var writer = new StreamWriter(filePath))
            {
                var columnNames = table.Columns.Cast<DataColumn>()
                    .Select(col => $"\"{col.ColumnName}\"");
                writer.WriteLine(string.Join(",", columnNames));

                foreach (DataRow row in table.Rows)
                {
                    var fields = row.ItemArray.Select(field =>
                        $"\"{field?.ToString()?.Replace("\"", "\"\"")}\""
                    );

                    writer.WriteLine(string.Join(",", fields));
                }
            }
        }


        private void CheckConnection()
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                throw new InvalidOperationException("Database connection is not established. Call Connect() first.");
            }
        }

        private void EnsureNotNullOrWhiteSpace(string value, string paramName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException($"Parameter '{paramName}' cannot be null or whitespace.", paramName);
            }
        }
    }
}
