using System;
using System.Collections.Generic;
using System.Data.SqlClient;


namespace ConsoleTalkdeskReportGenerator
{
    internal class GetAgents
    {
        private IDatabase _database;

        public GetAgents(IDatabase database)
        {
            _database = database;
        }

        public void GetAgentsList()
        {
            List<Agent> agentList = new List<Agent>();
            SqlConnection connection = _database.OpenConnection();

            string sql = @"
                SELECT DISTINCT( [UserName] ), 
                               [UserID] 
                FROM   [UserStatus] WITH(NOLOCK) 
                ORDER  BY [UserName] ASC";

            SqlCommand command = new SqlCommand(sql, connection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Console.WriteLine(reader["UserName"].ToString());
                Console.WriteLine(reader["UserID"].ToString());
            }

            _database.CloseConnection();
        }
    }
}
