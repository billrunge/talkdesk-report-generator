using System;
using System.Data.SqlClient;

namespace TalkdeskReportGenerator
{
    public class Database
    {
        private readonly SqlConnection _connection;
        private readonly string _dataSource = "IL1IPRFLCTDB005";
        private readonly string _database = "Talkdesk";
        private readonly int _timeout = 30;


        public Database()
        {
            _connection = new SqlConnection($"Server={_dataSource};" +
                                            $"Database={_database};" +
                                            $"Integrated Security=True;" +
                                            $"Connection Timeout={_timeout};");
        }

        public SqlConnection GetConnection()
        {
            try
            {
                _connection.Open();
            }
            catch (Exception e)
            {
                //need better error handling
                Console.WriteLine(e.ToString());
            }

            return _connection;
        }
    }
}
