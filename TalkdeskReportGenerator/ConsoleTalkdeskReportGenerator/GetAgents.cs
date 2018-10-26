using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ConsoleTalkdeskReportGenerator
{
    class GetAgents
    {
        private IDatabase _database;

        public GetAgents(IDatabase database)
        {
            _database = database;
        }

        public List<Agent> GetAgentsList()
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
                string name = reader["UserName"].ToString();
                string userId = reader["UserID"].ToString();
                
                if (userId == null)
                {
                    throw new NullReferenceException("UserID returned from database was null.");
                }
                Agent agent = new Agent()
                {
                    Name = name,
                    UserId = userId
                };
                agentList.Add(agent);
            }

            _database.CloseConnection();
            return agentList;
        }
    }
}
