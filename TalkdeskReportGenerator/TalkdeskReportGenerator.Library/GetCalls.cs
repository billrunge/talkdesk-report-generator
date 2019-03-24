using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TalkdeskReportGenerator.Library
{
    class GetCalls
    {
        private IDatabase _database;
        private readonly ILog _log;

        public GetCalls(IDatabase database, ILog log)
        {
            _database = database;
            _log = log;
        }


        public async Task<List<Call>> GetCallListAsync(string userName, DateTime statusStart, DateTime statusEnd, int offset)
        {
            List<Call> calls = new List<Call>();

            TimeSpan utcOffset = await Task.Run(() => TimeSpan.FromHours(offset));

            using (SqlConnection connection = _database.GetConnection())
            {
                if (_log.IsDebugEnabled)
                {
                    _log.Debug("GetStatuses.GetStatusesListAsync - Opening SQL connection");
                }

                await connection.OpenAsync();

                string sql = @"
                        SELECT Count(*) AS [Count], 
                               [CallType] 
                        FROM   [Calls] 
                        WHERE  [UserName] = @UserName 
                               AND [CallEnd] > @StatusStart 
                               AND [CallStart] < @StatusEnd 
                        GROUP  BY [CallType] ";

                SqlParameter userIdParam = new SqlParameter("@UserName", SqlDbType.NVarChar)
                {
                    Value = userName
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
                    _log.Debug($"GetCalls.GetCallListAsync - Executing query = {Environment.NewLine} {logQuery}");
                }

                SqlDataReader reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    if (!int.TryParse(reader["Count"].ToString(), out int count))
                    {
                        throw new Exception("Unable to cast count returned from database to int");
                    };

                    Enum.TryParse(reader["CallType"].ToString(), out CallType type);

                    Call call = new Call()
                    {
                        Count = count,
                        Type = type                       
                    };
                    calls.Add(call);
                }
            }
            return calls;
        }
    }
}
