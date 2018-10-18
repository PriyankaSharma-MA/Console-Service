using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace OTIS_Email_Service.App_Code
{
    public class LoggerHelper
    {
       // private static String exepurl;
        static SqlConnection con;

        /// <summary> 
        /// Log the exception
        /// </summary>
        public static void ExcpLogger(string _excpClass, string _excpMethod, Exception exdb)
        {
            string constr = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString.ToString();
            con = new SqlConnection(constr);
            con.Open();
            SqlCommand com = new SqlCommand("APP.PostExceptionLog", con);
            com.CommandType = CommandType.StoredProcedure;
            com.Parameters.AddWithValue("@LogExceptionClass", _excpClass);
            com.Parameters.AddWithValue("@LogExceptionMethod", _excpMethod);
            com.Parameters.AddWithValue("@LogException", exdb.Message.ToString());
            com.Parameters.AddWithValue("@LogExceptionCreatedDate", DateTime.Now);
            com.ExecuteNonQuery();
            con.Close();
        }
      
    }
}