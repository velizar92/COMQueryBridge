using System.Runtime.InteropServices;

namespace COMQueryBridge
{
    [ComVisible(true)]
    [Guid("C53A5A7D-7313-40C4-A628-0C0D4256EFD4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("COMQueryBridge.SQLQueryBridge")]
    public class SQLQueryBridge : ISQLQueryBridge
    {
        public void Connect(string connectionString)
        {
            throw new NotImplementedException();
        }

        public void Disconnect()
        {
            throw new NotImplementedException();
        }

        public int ExecuteNonQuery(string sqlQuery)
        {
            throw new NotImplementedException();
        }

        public string ExecuteQueryAsJson(string sqlQuery)
        {
            throw new NotImplementedException();
        }

        public string ExecuteScalar(string sqlQuery)
        {
            throw new NotImplementedException();
        }

        public void ExportResults()
        {
            throw new NotImplementedException();
        }
    }
}
