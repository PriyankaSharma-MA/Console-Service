using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace OTIS_Email_Service.App_Code
{
    public class SQLHelper
    {
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString.ToString();

        /// <summary> 
        /// Set the connection, command, and then execute the command with non query. 
        /// </summary> 
        public static Int32 ExecuteNonQuery(String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);
                try
                {
                    using (SqlCommand cmd = new SqlCommand(commandText, conn))
                    {
                        cmd.CommandType = commandType;
                        cmd.Parameters.AddRange(parameters);

                        conn.Open();
                        return cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.ExcpLogger("SQLHelper", "ExecuteNonQuery", ex);
                    conn.Close();
                    return 0;
                }

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("SQLHelper", "ExecuteNonQuery", ex);
                return 0;
            }
            finally
            {
                conn.Close();
            }

        }

        /// <summary> 
        /// Set the connection, command, and then execute the command with non query. 
        /// </summary> 
        public static string ExecuteNonQueryWithOutPut(String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);
                try
                {
                    using (SqlCommand cmd = new SqlCommand(commandText, conn))
                    {
                        cmd.CommandType = commandType;
                        cmd.Parameters.AddRange(parameters);
                        cmd.Parameters["@Result"].Direction = ParameterDirection.Output;
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        return cmd.Parameters["@Result"].Value.ToString();
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.ExcpLogger("SQLHelper", "ExecuteNonQueryWithOutPut", ex);
                    conn.Close();
                    return "";
                }

            }
            catch (Exception ex)
            {
                App_Code.LoggerHelper.ExcpLogger("SQLHelper", "ExecuteNonQueryWithOutPut", ex);
                return "";
            }
            finally
            {
                conn.Close();
            }

        }

        /// <summary> 
        /// Set the connection, command, and then execute the command and only return one value. 
        /// </summary> 
        public static Object ExecuteScalar(String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);
                try
                {
                    using (SqlCommand cmd = new SqlCommand(commandText, conn))
                    {
                        cmd.CommandType = commandType;
                        cmd.Parameters.AddRange(parameters);

                        conn.Open();
                        return cmd.ExecuteScalar();
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.ExcpLogger("SQLHelper", "ExecuteScalar", ex);
                    conn.Close();
                    return null;
                }

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("SQLHelper", "ExecuteScalar", ex);
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        /// <summary> 
        /// Set the connection, command, and then execute the command with query and return the reader. 
        /// </summary> 
        public static SqlDataReader ExecuteReader(String commandText, CommandType commandType, params SqlParameter[] parameters)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);

                using (SqlCommand cmd = new SqlCommand(commandText, conn))
                {
                    cmd.CommandType = commandType;
                    cmd.Parameters.AddRange(parameters);

                    conn.Open();
                    SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                    return reader;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("SQLHelper", "ExecuteReader", ex);
                conn.Close();
                return null;
            }
        }

        /// <summary> 
        /// Set the connection, command, and then execute the command with query and return the reader. 
        /// </summary> 
        public static SqlDataReader ExecuteReaderWithoutParameter(String commandText, CommandType commandType)
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection(connectionString);

                using (SqlCommand cmd = new SqlCommand(commandText, conn))
                {
                    cmd.CommandType = commandType;
                    conn.Open();
                    SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                    return reader;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("SQLHelper", "ExecuteReader", ex);
                conn.Close();
                return null;
            }
        }

        /// <summary> 
        /// Set the connection, command, and then execute the command with query. 
        /// </summary> 
        public static DataTable ExecuteAdapterWithoutParameter(String commandText, CommandType commandType)
        {
            SqlConnection conn = null;
            DataTable ExecutableData = null;
            try
            {
                conn = new SqlConnection(connectionString);

                using (SqlCommand cmd = new SqlCommand(commandText, conn))
                {
                    cmd.CommandType = commandType;
                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        using (ExecutableData = new DataTable())
                        {
                            sda.Fill(ExecutableData);
                        }
                    }

                    return ExecutableData;
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("SQLHelper", "ExecuteReader", ex);
                conn.Close();
                return null;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}