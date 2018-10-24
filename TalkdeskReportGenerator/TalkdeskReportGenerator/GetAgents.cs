using System;
using System.Collections.Generic;
using System.Data.SqlClient;


namespace TalkdeskReportGenerator
{
    internal class GetAgents
    {
        private IDatabase _database;

        public GetAgents(IDatabase database)
        {
            _database = database;
        }

        public List<Agent> GetAgentsList()
        {
            List<Agent> agentList = 
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
        }
    }
}
