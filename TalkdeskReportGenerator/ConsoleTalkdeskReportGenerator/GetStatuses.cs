using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace ConsoleTalkdeskReportGenerator
{
    internal interface IGetStatuses
    {
        List<Status> GetStatusesList(string userId, DateTime statusStart, DateTime statusEnd);
        string GetUserIdFromName(string name);
    }

    internal class GetStatuses : IGetStatuses
    {
        private IDatabase _database;

        public GetStatuses(IDatabase database)
        {
            _database = database;
        }

        public List<Status> GetStatusesList(string userId, DateTime statusStart, DateTime statusEnd)
        {
            List<Status> statuses = new List<Status>();

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

                Status status = new Status()
                {
                    StatusLabel = statusLabel,
                    StatusTime = statusTime
                };

                statuses.Add(status);
            }
            connection.Close();
            return statuses;
        }

        public string GetUserIdFromName(string name)
        {
            SqlConnection connection = _database.OpenConnection();
            string userId = "";

            string sql = @"
                SELECT TOP 1 [UserID]
                FROM [UserStatus] WITH(NOLOCK)
                WHERE [UserName] = @UserName";

            SqlParameter userNameParam = new SqlParameter("@UserName", SqlDbType.NVarChar)
            {
                Value = name
            };

            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.Add(userNameParam);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                userId = reader["UserID"].ToString();
            }

            connection.Close();
            return userId;
        }
    }
}
