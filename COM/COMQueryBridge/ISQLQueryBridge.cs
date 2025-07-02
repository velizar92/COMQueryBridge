using System.Runtime.InteropServices;

namespace COMQueryBridge
{
    [ComVisible(true)]
    [Guid("CEE0F91F-AF2B-4557-9454-509B58C28717")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISQLQueryBridge
    {
        void Connect(string connectionString);

        void Disconnect();

        string ExecuteScalar(string sqlQuery);

        int ExecuteNonQuery(string sqlQuery);

        string ExecuteQueryAsJson(string sqlQuery);

        void ExportResults();
    }
}
