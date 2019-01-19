using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;

namespace WpfTalkdeskReportGenerator
{ 
    internal interface IGetStatuses
    {
        //List<Status> GetStatusesList(string userId, DateTime statusStart, DateTime statusEnd, int UtcOffset);
        //string GetUserIdFromName(string name);
        Task<string> GetUserIdFromNameAsync(string name);
        Task<List<Status>> GetStatusesListAsync(string userId, DateTime statusStart, DateTime statusEnd, int offset);
    }

    internal class GetStatuses : IGetStatuses
    {
        private IDatabase _database;

        public GetStatuses(IDatabase database)
        {
            _database = database;
        }

        //public List<Status> GetStatusesList(string userId, DateTime statusStart, DateTime statusEnd, int offset)
        //{
        //    List<Status> statuses = new List<Status>();
        //    TimeSpan utcOffset = TimeSpan.FromHours(offset);
            
        //    SqlConnection connection = _database.OpenConnection();

        //    string sql = @"
        //        SELECT Sum(CASE 
        //                     WHEN ( @StatusStart <= [StatusStart] AND @StatusEnd >= [StatusEnd] ) 
        //                            THEN [StatusTime] 
        //                     WHEN ( @StatusStart <= [StatusStart] AND @StatusEnd < [StatusEnd] ) 
        //                            THEN Datediff(SECOND, [StatusStart], @StatusEnd) 
        //                     WHEN ( @StatusStart > [StatusStart] AND @StatusEnd <= [StatusEnd] ) 
        //                            THEN Datediff(SECOND, @StatusStart, @StatusEnd) 
        //                     WHEN ( @StatusStart > [StatusStart] AND @StatusEnd > [StatusEnd] ) 
        //                            THEN Datediff(SECOND, @StatusStart, [StatusEnd])  
        //                   END) AS [StatusTime], 
        //               [StatusLabel] 
        //        FROM   [UserStatus] WITH(NOLOCK) 
        //        WHERE  [UserID] = @UserID 
        //               AND [StatusEnd] > @StatusStart 
        //               AND [StatusStart] < @StatusEnd 
        //        GROUP  BY [StatusLabel]";

        //    SqlParameter userIdParam = new SqlParameter("@UserID", SqlDbType.NVarChar)
        //    {
        //        Value = userId                
        //    };

        //    SqlParameter statusStartParam = new SqlParameter("@StatusStart", SqlDbType.DateTime2)
        //    {
        //        Value = statusStart.Add(utcOffset)
        //    };

        //    SqlParameter statusEndParam = new SqlParameter("@StatusEnd", SqlDbType.DateTime2)
        //    {
        //        Value = statusEnd.Add(utcOffset)
        //    };

        //    //Console.WriteLine($"UTC Offset: {utcOffset.ToString()}");
        //    //Console.WriteLine($"UserID: {userId}, Status Start: {statusStart}, With Offset: {statusStart.Add(utcOffset)} Status End: {statusEnd}, With Offset: {statusEnd.Add(utcOffset)}");

        //    SqlCommand command = new SqlCommand(sql, connection);

        //    command.Parameters.Add(userIdParam);
        //    command.Parameters.Add(statusStartParam);
        //    command.Parameters.Add(statusEndParam);



        //    SqlDataReader reader = command.ExecuteReader();
        //    while (reader.Read())
        //    {

        //        string statusLabel;

        //        if (!int.TryParse(reader["StatusTime"].ToString(), out int statusTime))
        //        {
        //            throw new Exception("Unable to cast status time in seconds returned from database to int");
        //        };

        //        statusLabel = reader["StatusLabel"].ToString();

        //        Status status = new Status()
        //        {
        //            DayName = statusStart.DayOfWeek.ToString(),
        //            StatusLabel = statusLabel,
        //            StatusTime = statusTime
        //        };

        //        statuses.Add(status);
        //    }
        //    connection.Close();
        //    return statuses;
        //}


        public async Task<List<Status>> GetStatusesListAsync(string userId, DateTime statusStart, DateTime statusEnd, int offset)
        {
            List<Status> statuses = new List<Status>();
            TimeSpan utcOffset = await Task.Run(() => TimeSpan.FromHours(offset));

            using (SqlConnection connection = _database.GetConnection())
            {
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




        //public string GetUserIdFromName(string name)
        //{


        //    SqlConnection connection = _database.OpenConnection();
        //    string userId = "";

        //    string sql = @"
        //        SELECT TOP 1 [UserID]
        //        FROM [UserStatus] WITH(NOLOCK)
        //        WHERE [UserName] = @UserName";

        //    SqlParameter userNameParam = new SqlParameter("@UserName", SqlDbType.NVarChar)
        //    {
        //        Value = name
        //    };

        //    SqlCommand command = new SqlCommand(sql, connection);
        //    command.Parameters.Add(userNameParam);
        //    SqlDataReader reader = command.ExecuteReader();
        //    while (reader.Read())
        //    {
        //        userId = reader["UserID"].ToString();
        //    }

        //    connection.Close();
        //    return userId;
        //}

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
