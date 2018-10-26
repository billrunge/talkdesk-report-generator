using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ConsoleTalkdeskReportGenerator
{
    class GetAgentStatuses
    {
        private IDatabase _database;

        public GetAgentStatuses(IDatabase database)
        {
            _database = database;
        }

        public List<AgentStatus> GetAgentStatusesList(string userId, DateTime statusStart, DateTime statusEnd)
        {
            List<AgentStatus> agentStatuses = new List<AgentStatus>();

            SqlConnection connection = _database.OpenConnection();

            string sql = @"
                SELECT SUM([StatusTime]) AS [StatusTime], 
                       [StatusLabel] 
                FROM   [UserStatus] WITH(NOLOCK) 
                WHERE  [UserID] = @UserID 
                       AND [StatusStart] > @StatusStart 
                       and [StatusEnd] < @StatusEnd 
                GROUP  BY [StatusLabel]";

            SqlParameter userIdParam = new SqlParameter("@UserID", SqlDbType.NVarChar)
            {
                Value = userId
            };

            SqlParameter statusStartParam = new SqlParameter("@StatusStart", SqlDbType.DateTime)
            {
                Value = statusStart
            };

            SqlParameter statusEndParam = new SqlParameter("@StatusEnd", SqlDbType.DateTime)
            {
                Value = statusEnd
            };

            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.Add(userIdParam);
            command.Parameters.Add(statusStartParam);
            command.Parameters.Add(statusEndParam);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                string statusLabel;

                if (!int.TryParse(reader["StatusTime"].ToString(), out int statusTime))
                {
                    throw new Exception("Unable to cast status time in seconds returned from database to int");
                };

                statusLabel = reader["StatusLabel"].ToString();

                AgentStatus agentStatus = new AgentStatus()
                {
                    StatusLabel = statusLabel,
                    StatusTime = statusTime
                };

                agentStatuses.Add(agentStatus);
            }
            return agentStatuses;
        }
    }
}
