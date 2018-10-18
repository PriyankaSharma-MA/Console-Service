using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace OTIS_Email_Service.App_Code
{
    public class NprintingReqResRepository
    {
        public static CookieContainer cookies = new CookieContainer();
        public static HttpWebRequest request = null;
        public static HttpWebResponse response = null;
        public static HttpWebResponse nPrintingResponse = null;
        public static string result = string.Empty;
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString.ToString();
        public static List<Action> listOfActions = new List<Action>();

        static NprintingReqResRepository()
        {
            NPrintingAthentication();
        }

        /// <summary> 
        /// NPrinting Get Request
        /// </summary> 
        public static HttpWebResponse GetRequest(string _uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_uri);
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            request.CookieContainer = cookies;
            request.Method = "GET";
            request.KeepAlive = true;
            request.UserAgent = "Windows";
            request.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
            request.UseDefaultCredentials = true;
            request.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["NpritingUserName"].ToString(),
                ConfigurationManager.AppSettings["NpritingPassword"].ToString());
            request.ContentType = "application/json";
            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("NprintingReqResRepository", "GetRequest", ex);
                response = null;
            }
            finally
            {
                request.Abort();
                response.Close();
            }
            return response;
        }

        /// <summary> 
        /// NPrinting Post Request
        /// </summary> 
        public static string PostRequest(string _uri, string _body)
        {
            request = (HttpWebRequest)WebRequest.Create(_uri);
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            request.Headers.Add("X-XSRF-TOKEN", cookies.GetCookies(request.RequestUri)["NPWEBCONSOLE_XSRF-TOKEN"].Value);
            request.CookieContainer = cookies;
            request.Method = "POST";
            request.KeepAlive = true;
            request.UserAgent = "Windows";
            request.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
            request.UseDefaultCredentials = true;
            request.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["NpritingUserName"].ToString(),
            ConfigurationManager.AppSettings["NpritingPassword"].ToString());
            request.ContentType = "application/json"; //"application/x-www-form-urlencoded";
            request.Accept = "application/json";
            //using (var streamWriter = new StreamWriter(request.GetRequestStream()))
            //{
            //    streamWriter.Write(_body);
            //    streamWriter.Flush();
            //    streamWriter.Close();
            //}
            string postResoponse = "";
            try
            {
                nPrintingResponse = (HttpWebResponse)request.GetResponse();
                if (nPrintingResponse.StatusCode == HttpStatusCode.OK || nPrintingResponse.StatusCode == HttpStatusCode.Accepted)
                {
                    StreamReader responseStream = new StreamReader(nPrintingResponse.GetResponseStream());
                    postResoponse = responseStream.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("NprintingReqResRepository", "PostRequest", ex);
                postResoponse = "";
            }
            finally
            {
                request.Abort();
                response.Close();
            }
            
            return postResoponse;
        }

        /// <summary> 
        /// Set the NPrinting Authetication
        /// </summary> 
        public static HttpWebResponse NPrintingAthentication()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            request = (HttpWebRequest)WebRequest.Create(ConfigurationManager.AppSettings["NPrintingAuthentication"].ToString());
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            request.CookieContainer = cookies;
            request.Method = "Get";
            request.KeepAlive = true;
            request.UserAgent = "Windows";
            request.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
            request.UseDefaultCredentials = true;
            request.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["NpritingUserName"].ToString(),
            ConfigurationManager.AppSettings["NpritingPassword"].ToString());
            request.ContentType = "application/json";
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    string nPrintResponse = "";
                    using (System.IO.StreamReader sr = new System.IO.StreamReader(response.GetResponseStream()))
                    {
                        nPrintResponse = sr.ReadToEnd();
                    }
                }
                return response;
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("NprintingReqResRepository", "NPrintingAthentication", ex);
                return null;
            }
            finally
            {
                request.Abort();
            }
        }

        /// <summary> 
        /// NPrinting Task Execution
        /// </summary>
        public static string TasksExecution(string passedExecutionTaskID, Int32 reportid)
        {
            // POST TASK EXECUTION
          

            string nPrintingResponse = string.Empty;
            if (response.StatusCode == HttpStatusCode.OK)
            {
                string taskExecutionUrl = ConfigurationManager.AppSettings["NPrintingPutGetTaskExecution"] + passedExecutionTaskID + "/executions";
                string reqBody = string.Empty;
                TaskExecution TaskExec = new TaskExecution();
                TaskExec.type = "Qlik.NPrinting.Repo.Model.PublishReportTask";
                TaskExec.task = passedExecutionTaskID;
                reqBody = JsonConvert.SerializeObject(TaskExec);
                listOfActions.Add(() => PostRequest(taskExecutionUrl, reqBody));
                //nPrintingResponse = PostRequest(taskExecutionUrl, reqBody);
            }
            Parallel.Invoke(listOfActions.ToArray());
            NprintingReqResRepository obj = new NprintingReqResRepository();
           // obj.WriteToFile(reportid.ToString());

            obj.InsertLog(reportid.ToString(), passedExecutionTaskID, DateTime.Now);
            result = result + "_+" + reportid.ToString();
            return nPrintingResponse;
        }

        public static string excuteReportBunch(DataTable taskExecutionDetails, Int32 ReportID)
        {
            string result = string.Empty;
            var tasks = new List<Task>();
            var dtUniqueTaskList = taskExecutionDetails.AsEnumerable()
              .Select(row => new
              {
                  TaskID = row.Field<string>("NprintingTaskID"),
                  ReportID = row.Field<Int32>("ReportID"),

              }).Where(item => item.ReportID == ReportID)
              .Distinct();


            foreach (var Tlist in dtUniqueTaskList.ToList())
            {

                string TaskID = Tlist.TaskID;
               // NprintingReqResRepository obj = new NprintingReqResRepository();
              //  obj.WriteToFile(ReportID.ToString(), taskExecutionDetails.Rows.Count);
                NprintingReqResRepository.TasksExecution(TaskID, ReportID);

            }
            return result;

        }
        public void InsertLog(string ReportId, string TaskId, DateTime Time)
        {
            SqlConnection conn = null;
            conn = new SqlConnection(connectionString);
            try
            {
               
                string sql = "INSERT INTO Log (ReportId,TaskId,Time) values (@ReportId,@TaskId,@Time)";
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add("@ReportId", SqlDbType.VarChar);
                cmd.Parameters.Add("@TaskId", SqlDbType.VarChar);
                cmd.Parameters.Add("@Time", SqlDbType.DateTime);
                cmd.Parameters["@ReportId"].Value = ReportId;
                cmd.Parameters["@TaskId"].Value = TaskId;
                cmd.Parameters["@Time"].Value = Time;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                conn.Close();

            }
            finally
            {
                conn.Close();
            }
        }
        public void WriteToFile(string text, int count)
        {
            string path = "C:\\OTIS_project\\Console Service\\OTIS_Email_Service\\log.txt";
            //StreamWriter writer = new StreamWriter(path, true);
          
            //try
            //{
            //    using (writer)
            //    {
                  
            //        //    writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
            //        writer.WriteLine(text + "_" + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
            //        writer.Close();
            //    }
            //}
            //catch(Exception ex)
            //{
            //    writer.Close();
            //   // WriteToFile(text);
            //}

            FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
           StreamWriter fp = new StreamWriter(fs);
           Object locker = new Object();
            Parallel.For(0, count, (k) =>
            {
                try
                {
                   // Int32 iReturn = ProcRegister((Int32)table.Rows[0][0]);
                    lock (locker)
                    {
                        fp.WriteLine(text + "_" + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                        fp.Flush();
                    }
                }
                catch { }
                finally { }
            });
        }
    }

}