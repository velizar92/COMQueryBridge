# COMQueryBridge

A COM-visible, provider-agnostic SQL query bridge for .NET applications that allows executing SQL queries and exporting results in CSV format. Designed for easy integration with any COM-compatible client, supporting SQL Server and SQLite database providers configured via the app configuration.

## Features
- Connects to any database using ADO.NET provider factories (SQL Server, SQLite, etc.). 

- Execute scalar queries, non-query commands, and read queries.

- Returns query results serialized as JSON for easy consumption.

- Export query results from JSON to CSV files.

- COM-visible with ProgID for seamless interop with COM clients.

## Installation

This project includes an **Inno Setup** installer for easy installation and COM registration.

### Steps to install

1. Run the installer (`COMQueryBridgeSetup.exe`).
2. The installer will copy the files, register the COM DLL, and create shortcuts if applicable.
3. To uninstall, use **Add or Remove Programs** or run the uninstaller from the installation folder.

## Configuration

- The installer includes a sample configuration file `appsettings.example.json`.
- Before using the COM component, copy and rename it to `appsettings.json` in the installation folder.
- Edit `appsettings.json` to set your database connection string and provider name.
- This approach helps keep sensitive data out of the installer package.

### Building the installer from source

If you want to build the installer yourself:

- Download and install [Inno Setup](https://jrsoftware.org/isinfo.php).
- Open the provided `build_installer_script.iss` script in Inno Setup Compiler.
- Adjust paths if necessary.
- Compile the installer.

The installer handles registering and unregistering the COM DLL automatically.

---

*This installer simplifies deployment by registering the COM server and managing files properly on Windows.*


## Getting Started
Prerequisites:
- .NET with Windows COM support

- Relevant database provider NuGet packages installed (e.g., Microsoft.Data.SqlClient for SQL Server, Microsoft.Data.Sqlite for SQLite)

- Properly configured appsettings.json file

## Configuration

For SQL Server: 
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=your_server;Database=your_db;User Id=your_user;Password=your_password;"
  },
  "ProviderName": "Microsoft.Data.SqlClient"
}
```
For SQLite:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Data Source=path_to_your_database.sqlite;Version=3;"
  },
  "ProviderName": "System.Data.SQLite"
}
```

VB Script example (usage):
```vbscript
' Create the COM object
Dim sqlBridge
Set sqlBridge = CreateObject("COMQueryBridge.SQLQueryBridge")

' Connect to the database (connection string and provider from config)
sqlBridge.Connect()

' Insert a new record
Dim insertSql
insertSql = "INSERT INTO YourTable (Column1, Column2) VALUES ('Value1', 'Value2')"
Dim rowsInserted
rowsInserted = sqlBridge.ExecuteNonQuery(insertSql)
MsgBox "Rows inserted: " & rowsInserted

' Update existing records
Dim updateSql
updateSql = "UPDATE YourTable SET Column2 = 'NewValue' WHERE Column1 = 'Value1'"
Dim rowsUpdated
rowsUpdated = sqlBridge.ExecuteNonQuery(updateSql)
MsgBox "Rows updated: " & rowsUpdated

' Execute a scalar query
Dim count
count = sqlBridge.ExecuteScalar("SELECT COUNT(*) FROM YourTable")
MsgBox "Total rows: " & count

' Execute a query and get results as JSON
Dim jsonResult
jsonResult = sqlBridge.ExecuteQueryAsJson("SELECT * FROM YourTable")
MsgBox "Query result in JSON:" & vbCrLf & jsonResult

' Export the JSON result to CSV file
sqlBridge.ExportJsonToCsv jsonResult, "C:\temp\output.csv"

' Disconnect from the database
sqlBridge.Disconnect()
```

# COMQueryBridge API

## Class: SQLQueryBridge
This COM-visible class provides a generic database query bridge supporting multiple database providers via ADO.NET. It allows connecting, executing queries, and exporting data in CSV format.

## Methods
### Connect()
Description:
- Opens a database connection using the connection string and provider name from the application configuration file.

Usage:
- Call this before executing any queries.

Exceptions:
- Throws InvalidOperationException if configuration is missing or connection cannot be established.

### Disconnect()
Description:
- Closes and disposes the open database connection if any.

Usage:
- Call this to cleanly close the connection when done.

### ExecuteNonQuery(string sqlQuery) : int
Description:
- Executes a SQL command that does not return rows, such as INSERT, UPDATE, or DELETE.

Parameters:
- sqlQuery – The SQL command to execute (non-empty string).

Returns:
- The number of rows affected by the command.

Exceptions:
- Throws ArgumentException if sqlQuery is null or whitespace. Throws InvalidOperationException if not connected.

### ExecuteQueryAsJson(string sqlQuery) : string
Description:
- Executes a SQL SELECT query and returns the result as a JSON array string.

Parameters:
- sqlQuery – The SQL query to execute (non-empty string).

Returns:
- JSON serialized string of the query results as an array of objects. Returns "[]" if no rows found.

Exceptions:
- Throws ArgumentException if sqlQuery is null or whitespace. Throws InvalidOperationException if not connected.

### ExecuteScalar(string sqlQuery) : string
Description:
- Executes a SQL query and returns the value of the first column of the first row in the result set.

Parameters:
- sqlQuery – The SQL query to execute (non-empty string).

Returns:
- The scalar value as a string, or null if the result is empty or DBNull.

Exceptions:
- Throws ArgumentException if sqlQuery is null or whitespace. Throws InvalidOperationException if not connected.

### ExportJsonToCsv(string jsonData, string filePath)
Description:
- Exports JSON-formatted query result data into a CSV file.

Parameters:

- jsonData – JSON string representing a DataTable (non-empty).

- filePath – The path of the CSV file to create (non-empty).

Usage:
- Typically called with JSON returned from ExecuteQueryAsJson.

Exceptions:
- Throws ArgumentException if any parameter is null or whitespace. Throws InvalidOperationException if JSON cannot be deserialized.

## Notes
- Ensure the provider NuGet package corresponding to your database is installed and referenced.

- Connection strings and provider names must be correctly configured for your environment.

- This component is designed to be used by COM clients; ensure proper registration for COM interop.

## License
MIT License

## Contributing
Feel free to open issues or submit pull requests.

## Contact
Velizar Gerasimov - velizar9209@gmail.com

Project Link: https://github.com/velizar92/COMQueryBridge


