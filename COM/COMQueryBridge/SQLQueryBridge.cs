using Microsoft.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Runtime.InteropServices;

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
            throw new NotImplementedException();
        }

        public string ExecuteScalar(string sqlQuery)
        {           
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

        public void ExportResults(string jsonData, string filePath, string format)
        {
            throw new NotImplementedException();
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
