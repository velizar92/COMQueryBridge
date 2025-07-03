using System.Runtime.InteropServices;

namespace COMQueryBridge
{
    /// <summary>
    /// Interface for executing SQL queries and exporting results via COM interop.
    /// </summary>
    [ComVisible(true)]
    [Guid("CEE0F91F-AF2B-4557-9454-509B58C28717")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISQLQueryBridge
    {
        /// <summary>
        /// Establishes a connection to the database. The connection string comes from the application's configuration file.
        /// </summary>
        void Connect();

        /// <summary>
        /// Closes the current database connection, if any.
        /// </summary>
        void Disconnect();

        /// <summary>
        /// Executes a SQL query that returns a single scalar value.
        /// </summary>
        string ExecuteScalar(string sqlQuery);

        /// <summary>
        /// Executes a SQL query that does not return any data (e.g., INSERT, UPDATE, DELETE).
        /// </summary>
        /// <param name="sqlQuery">The SQL query to execute.</param>
        /// <returns>The number of rows affected.</returns>
        int ExecuteNonQuery(string sqlQuery);

        /// <summary>
        /// Executes a SQL query and returns the result as a JSON string.
        /// </summary>
        /// <param name="sqlQuery">The SQL SELECT query to execute.</param>
        /// <returns>The query result in JSON format.</returns>
        string ExecuteQueryAsJson(string sqlQuery);

        /// <summary>
        /// Exports query results provided as JSON to a file in the specified format.
        /// </summary>
        /// <param name="jsonData">The JSON-formatted data to export.</param>
        /// <param name="filePath">The path to the output file.</param>
        /// <param name="format">The export format (e.g., "csv" or "xlsx").</param>
        void ExportResults(string jsonData, string filePath, string format);
    }
}
