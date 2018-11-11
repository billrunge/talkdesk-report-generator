﻿using System;
using System.Data;
using System.Data.SqlClient;

namespace WpfTalkdeskReportGenerator
{
    public interface IDatabase
    {
        SqlConnection OpenConnection();
        void CloseConnection();
    }

    public class Database : IDatabase
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

        public SqlConnection OpenConnection()
        {
            if (_connection.State != ConnectionState.Open)
            {
                try
                {
                    _connection.Open();
                }
                catch (Exception e)
                {
                    /* need better error handling */
                    Console.WriteLine(e.ToString());
                }
            }

            return _connection;
        }

        public void CloseConnection()
        {
            try
            {
                _connection.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

        }
    }
}