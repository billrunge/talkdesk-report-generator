using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace WpfTalkdeskReportGenerator
{
    internal interface IGetStatuses
    {
        Task<string> GetUserIdFromNameAsync(string name);
        Task<List<Status>> GetStatusesListAsync(string userId, DateTime statusStart, DateTime statusEnd, int offset);
    }

    internal class GetStatuses : IGetStatuses
    {
        private IDatabase _database;
        private readonly ILog _log;

        public GetStatuses(IDatabase database, ILog log)
        {
            _database = database;
            _log = log;
        }

        public async Task<List<Status>> GetStatusesListAsync(string userId, DateTime statusStart, DateTime statusEnd, int offset)
        {
            List<Status> statuses = new List<Status>();

            TimeSpan utcOffset = await Task.Run(() => TimeSpan.FromHours(offset));

            using (SqlConnection connection = _database.GetConnection())
            {
                if (_log.IsDebugEnabled)
                {
                    _log.Debug("GetStatuses.GetStatusesListAsync - Opening SQL connection");
                }

                await connection.OpenAsync();

                string sql = @"
                SELECT Sum(CASE 
                             WHEN ( @StatusStart <= [StatusStart] AND @StatusEnd >= [StatusEnd] ) 
                                    THEN [StatusTime] 
                             WHEN ( @StatusStart <= [StatusStart] AND @StatusEnd < [StatusEnd] ) 
                                    THEN Datediff(SECOND, [StatusStart], @StatusEnd) 
                             WHEN ( @StatusStart > [StatusStart] AND @StatusEnd <= [StatusEnd] ) 
                                    THEN Datediff(SECOND, @StatusStart, @StatusEnd) 
                             WHEN ( @StatusStart > [StatusStart] AND @StatusEnd > [StatusEnd] ) 
                                    THEN Datediff(SECOND, @StatusStart, [StatusEnd])  
                           END) AS [StatusTime], 
                       [StatusLabel] 
                FROM   [UserStatus] WITH(NOLOCK) 
                WHERE  [UserID] = @UserID 
                       AND [StatusEnd] > @StatusStart 
                       AND [StatusStart] < @StatusEnd 
                GROUP  BY [StatusLabel]";

                SqlParameter userIdParam = new SqlParameter("@UserID", SqlDbType.NVarChar)
                {
                    Value = userId
                };

                SqlParameter statusStartParam = new SqlParameter("@StatusStart", SqlDbType.DateTime2)
                {
                    Value = statusStart.Add(utcOffset)
                };

                SqlParameter statusEndParam = new SqlParameter("@StatusEnd", SqlDbType.DateTime2)
                {
                    Value = statusEnd.Add(utcOffset)
                };

                SqlCommand command = new SqlCommand(sql, connection);

                command.Parameters.Add(userIdParam);
                command.Parameters.Add(statusStartParam);
                command.Parameters.Add(statusEndParam);

                if (_log.IsDebugEnabled)
                {
                    string logQuery = $"SQL query = {command.CommandText}";
                    foreach (SqlParameter com in command.Parameters)
                    {
                        logQuery = logQuery.Replace(com.ToString(), $"'{com.Value.ToString()}'");
                    }
                    _log.Debug($"GetStatuses.GetStatusesListAsync - Executing query = {Environment.NewLine} {logQuery}");
                }

                SqlDataReader reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    string statusLabel;

                    if (!int.TryParse(reader["StatusTime"].ToString(), out int statusTime))
                    {
                        throw new Exception("Unable to cast status time in seconds returned from database to int");
                    };

                    statusLabel = reader["StatusLabel"].ToString();

                    Status status = new Status()
                    {
                        DayName = statusStart.DayOfWeek.ToString(),
                        StatusLabel = statusLabel,
                        StatusTime = statusTime
                    };
                    statuses.Add(status);
                }
            }
            return statuses;
        }

        public async Task<string> GetUserIdFromNameAsync(string name)
        {
            string userId = "";

            using (SqlConnection connection = _database.GetConnection())
            {
                await connection.OpenAsync();

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

                if (_log.IsDebugEnabled)
                {
                    string logQuery = $"SQL query = {command.CommandText}";
                    foreach (SqlParameter com in command.Parameters)
                    {
                        logQuery = logQuery.Replace(com.ToString(), $"'{com.Value.ToString()}'");
                    }
                    _log.Debug($"GetStatuses.GetUserIdFromNameAsync - Executing query = {Environment.NewLine} {logQuery}");
                }


                SqlDataReader reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    userId = reader["UserID"].ToString();
                }

            }
            return userId;

        }

    }
}
