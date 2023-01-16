using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace MergeStudentPDF2.Librarys
{
    public class DBHelper
    {
        protected DBHelper() 
        {
        } // end DBHelper

        private static string _ConnectionString = string.Empty;


        public static void SetConnectionInfo(string IP, string Port, string DBName, string DBUser, string Pwd)
        {
            _ConnectionString = $"server={IP};port={Port};database={DBName};user={DBUser};password={Pwd};charset=utf8;";
        } // end SetConnectionString


        public static bool TestConnection()
        {
            bool Result = false;

            try
            {
                using (MySqlConnection Conn = GetSqlConnection())
                {
                    Conn.Open();
                } // end uisng

                Result = true;
            }
            catch
            {
                Result = false;
            } // end try catch

            return Result;
        } // end TestConnection


        public static MySqlConnection GetSqlConnection()
        {
            if (_ConnectionString != string.Empty)
            {
                return new MySqlConnection(_ConnectionString);
            }
            else
            {
                throw new ArgumentException("ConnectionString is Null Or Empty", nameof(_ConnectionString));
            } // end if
        } // end GetSqlConnection
    } // end DBHelper
}
