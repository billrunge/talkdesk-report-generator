using log4net;
using System.Data.SqlClient;

namespace WpfTalkdeskReportGenerator
{
    public interface IDatabase
    {
        SqlConnection GetConnection();
    }

    public class Database : IDatabase
    {
        private readonly string _dataSource = "IL1IPRFLCTDB005";
        private readonly string _database = "Talkdesk";
        private readonly int _timeout = 30;
        private readonly string _connectionString;
        private readonly ILog _log;

        public Database(ILog log)
        {
            _log = log;
           _connectionString = $@"Server={_dataSource}; 
                                  Database={_database}; 
                                  Integrated Security=True; 
                                  Connection Timeout={_timeout};";
            if (_log.IsDebugEnabled)
            {
                _log.Debug($"Database.Database - Connection String = { _connectionString }");
            }

        }

        public SqlConnection GetConnection()
        {
            if (_log.IsDebugEnabled)
            {
                _log.Debug("Database.GetConnection - Sql Connection Requested");
            }
            return new SqlConnection(_connectionString);
        }
    }
}
