using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using System.Net.Mail;
using System.IO;
using System.Globalization;
using OTIS_Email_Service.App_Code;
using System.Data.SqlClient;
using System.Data;
using System.Web;
using System.Net;
using System.Threading;
using DansCSharpLibrary.Threading;

namespace OTIS_Email_Service
{
    class Program
    {
        public static string _sender = ConfigurationManager.AppSettings["OutlookUser"].ToString();
        public static string _password = ConfigurationManager.AppSettings["OutlookPass"].ToString();
        public static SqlDataReader rdr = null;
        public static bool IsAdhoc = false;

        static void Main(string[] args)
        {
            //Parallel
            Program obj = new Program();
            //NprintingReqResRepository objn = new NprintingReqResRepository();
            //objn.WriteToFile("hello",1);
            ExecuteNprintingTask();
            
            //ExecuteNprintingTaskAdhoc();
            //run after 5 min
            // System.Threading.Thread.Sleep(2 * 60 * 1000);
            
            GetEmailAndSaveAttachment();

        }



        /// <summary> 
        /// Set autentication in Nprinting and Get the Task details from database.
        /// </summary>
       // public async Task ExecuteNprintingTask()
        static  void ExecuteNprintingTask()
        {
            IsAdhoc = false;
            DataTable taskExecutionDetails = null;
            DataTable distinctReprots = null;
            try
            {
                taskExecutionDetails = new DataTable();
                taskExecutionDetails = GetTaskExecutionDetails();


                if (taskExecutionDetails != null && taskExecutionDetails.Rows.Count > 0)
                {
                  //  var dtUniqueReportList = taskExecutionDetails.AsEnumerable()
                  //.Select(row => new
                  //{
                  //    ReportID = row.Field<Int32>("ReportID"),

                  //})
                  //.Distinct();
                  //  dtUniqueReportList.ToArray();

                    //var tasks = new List<Task>();
                    var listOfActions = new List<Action>();
                    distinctReprots = taskExecutionDetails.AsEnumerable()
                  .GroupBy(row => new
                  {
                      ReportID = row.Field<Int32>("ReportID"),

                  })
                     .Select(group => group.First())
                     .CopyToDataTable();

                    NprintingReqResRepository obj = new NprintingReqResRepository();
                    obj.InsertLog("start", "start", DateTime.Now);

                    #region Parallel

                    foreach (DataRow dr in distinctReprots.Rows)
                    {
                        string TaskID = dr["NprintingTaskID"].ToString();
                        Int32 ReportID = Convert.ToInt32(dr["ReportID"]);

                        listOfActions.Add(() => NprintingReqResRepository.excuteReportBunch(taskExecutionDetails, ReportID));
                    }

                    Parallel.Invoke(listOfActions.ToArray());
                    #endregion


                    #region sequential

                    //foreach (DataRow dr in taskExecutionDetails.Rows)
                    //{
                    //    string TaskID = dr["NprintingTaskID"].ToString();
                    //    Int32 ReportID = Convert.ToInt32(dr["ReportID"]);

                    //    NprintingReqResRepository.TasksExecution(TaskID, ReportID);
                    //}

                    #endregion

                }
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("Window Service Program", "ExecuteNprintingTask", ex);
            }
        }
        //Task t1 = Task.Factory.StartNew(() => NprintingReqResRepository.TasksExecution("1"));
        //Task t2 = Task.Factory.StartNew(() => NprintingReqResRepository.TasksExecution("1"));

        //  var result = t1.Result.Concat(t2.Result);





        public async Task RunTasks()
        {
            IsAdhoc = false;
            DataTable taskExecutionDetails = null;
            var tasks = new List<Task>();
            try
            {
                taskExecutionDetails = new DataTable();
                taskExecutionDetails = GetTaskExecutionDetails();
                if (taskExecutionDetails != null && taskExecutionDetails.Rows.Count > 0)
                {

                    foreach (DataRow dr in taskExecutionDetails.Rows)
                    {
                        string TaskID = dr["NprintingTaskID"].ToString();
                        // NprintingReqResRepository.TasksExecution(TaskID);
                        tasks.Add(Task.Run(() => NprintingReqResRepository.TasksExecution(TaskID, 1)));
                    }
                    //Parallel.For(0, taskExecutionDetails.Rows.Count,
                    // async i =>
                    // {
                    //     await NprintingReqResRepository.TasksExecution(taskExecutionDetails.Rows[i]["NprintingTaskID"].ToString());
                    // });

                    await Task.WhenAll(tasks);
                }



                //  var tasks = new List<Task>
                //{
                // new Task(async () => await DoWork()),
                // //and so on with the other 9 similar tasks
                //}


                Parallel.ForEach(tasks, task =>
                {
                    task.Start();
                });

                Task.WhenAll(tasks).ContinueWith(done =>
                {
                    //Run the other tasks
                    Console.WriteLine(tasks);
                });
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("Window Service Program", "ExecuteNprintingTask", ex);
            }
        }
        //public async Task DoWork()
        //{
        //    var results = await NprintingReqResRepository.TasksExecution("1");
        //    foreach (var result in results)
        //    {
        //        await ReadFromNetwork(result.Url);
        //    }
        //}

        //}


        /// <summary> 
        /// Get the Task details from database
        /// </summary>
        static void ExecuteNprintingTaskAdhoc()
        {
            IsAdhoc = true;
            DataTable taskExecutionDetails = null;
            try
            {
                taskExecutionDetails = new DataTable();
                taskExecutionDetails = GetAdhocTaskExecutionDetails();
                if (taskExecutionDetails != null && taskExecutionDetails.Rows.Count > 0)
                {
                    foreach (DataRow dr in taskExecutionDetails.Rows)
                    {
                        string TaskID = dr["NprintingTaskID"].ToString();
                        NprintingReqResRepository.TasksExecution(TaskID, 1);
                    }

                }

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("Window Service Program", "ExecuteNprintingTask", ex);
            }
        }


        /// <summary> 
        /// Get email details from outlook and save attachments in local(Base directory).
        /// </summary>
        static void GetEmailAndSaveAttachment()
        {
            try
            {

                //Outlook.Application outlookApp = new Outlook.Application();
                //Outlook.NameSpace outlookNamspace = outlookApp.GetNamespace("mapi");
                // Outlook.MAPIFolder outlookInbox = outlookNamspace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                // outlookNamspace.Logon(_sender, _password, true, true);
                string startMailDate = DateTime.Now.ToString();
                string endMailDate = DateTime.Today.AddHours(23).AddMinutes(59).ToString("g");
                bool countInbox = false;
                //Outlook.Items outlookItems = null;
                //for (int i = 0; i <= 2; i++)
                //{
                //    if (!countInbox)
                //    {
                //        outlookItems = outlookInbox.Items;
                //        outlookItems = outlookInbox.Items.Restrict("[ReceivedTime] >= '" + startMailDate + "' AND [ReceivedTime] <='" + endMailDate + "'");
                //        if (outlookItems != null && outlookItems.Count > 0)
                //        {
                //            countInbox = true;
                //        }
                //        else
                //        {
                //            continue;
                //        }
                //    }
                //    else
                //    {
                //        continue;
                //    }
                //}

                //if (outlookItems != null && outlookItems.Count > 0)
                //{
                //    //Outlook.Attachments outlookAttachment;
                //    string directoryPath = AppDomain.CurrentDomain.BaseDirectory + "EmailAttachment";
                //    if (!Directory.Exists(directoryPath))
                //        Directory.CreateDirectory(directoryPath);

                //    foreach (Outlook.MailItem mail in outlookItems)
                //    {
                //        string[] mailSubject = mail.Subject.Split('#');
                //        string mailSubjectTaskId = string.Empty;
                //        if (mailSubject != null && mailSubject.Count() > 1)
                //            mailSubjectTaskId = mailSubject[1];
                //        if (!string.IsNullOrWhiteSpace(mailSubjectTaskId) && IsAdhoc == true)
                //        {
                //            outlookAttachment = mail.Attachments;
                //            if (outlookAttachment != null && outlookAttachment.Count > 0)
                //            {
                //                string attachmentDate = DateTime.Today.ToString("MM/dd/yyyy");
                //                string attachmentDatePath = AppDomain.CurrentDomain.BaseDirectory + "\\EmailAttachment\\" + "\\Adhoc\\" + attachmentDate.Trim();
                //                string emailAttPath = attachmentDatePath + "\\" + mailSubjectTaskId.Trim();

                //                if (!Directory.Exists(attachmentDatePath))
                //                    Directory.CreateDirectory(attachmentDatePath);

                //                if (!Directory.Exists(emailAttPath))
                //                    Directory.CreateDirectory(emailAttPath);

                //                for (int i = 1; i <= outlookAttachment.Count; i++)
                //                {
                //                    Outlook.Attachment attachment = outlookAttachment[i];

                //                    if (attachment.Type == Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue)
                //                    {
                //                        string fileName = Path.Combine(emailAttPath, attachment.FileName);
                //                        attachment.SaveAsFile(fileName);
                //                        string Result = SaveTaskAndAttDetails(mailSubjectTaskId.Trim(), attachment.FileName.Trim());
                //                    }
                //                }
                //            }
                //            else if (!string.IsNullOrWhiteSpace(mailSubjectTaskId) && IsAdhoc == false)
                //            {
                //                outlookAttachment = mail.Attachments;
                //                if (outlookAttachment != null && outlookAttachment.Count > 0)
                //                {
                //                    string attachmentDate = DateTime.Today.ToString("MM/dd/yyyy");
                //                    string attachmentDatePath = AppDomain.CurrentDomain.BaseDirectory + "\\EmailAttachment\\" + attachmentDate.Trim();
                //                    string emailAttPath = attachmentDatePath + "\\" + mailSubjectTaskId.Trim();

                //                    if (!Directory.Exists(attachmentDatePath))
                //                        Directory.CreateDirectory(attachmentDatePath);

                //                    if (!Directory.Exists(emailAttPath))
                //                        Directory.CreateDirectory(emailAttPath);

                //                    for (int i = 1; i <= outlookAttachment.Count; i++)
                //                    {
                //                        Outlook.Attachment attachment = outlookAttachment[i];

                //                        if (attachment.Type == Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue)
                //                        {
                //                            string fileName = Path.Combine(emailAttPath, attachment.FileName);
                //                            attachment.SaveAsFile(fileName);
                //                            string Result = SaveTaskAndAttDetails(mailSubjectTaskId.Trim(), attachment.FileName.Trim());
                //                        }
                //                    }
                //                }
                //            }
                //        }

                //    }
                //    SendEmailToUser(GetTaskAndUserDetail());
                //}
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "Main", ex);
            }
        }


        /// <summary> 
        /// Save attachments and task detail from outlook in database.
        /// </summary>
        static string SaveTaskAndAttDetails(string TaskID, string AttachmentName)
        {
            try
            {
                string taskQuery = DbRepository.SaveTaskDetail();
                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@TaskID", TaskID));
                parameters.Add(new SqlParameter("@AttachmentName", AttachmentName));
                parameters.Add(new SqlParameter("@Result", SqlDbType.NVarChar, 2000));
                string _result = SQLHelper.ExecuteNonQueryWithOutPut(taskQuery, CommandType.StoredProcedure, parameters.ToArray());
                return _result;
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("Window Service Program", "SaveTaskAndAttDetails", ex);
                return "";
            }

        }


        /// <summary> 
        /// Get Task  And User details from database.
        /// </summary>
        static DataTable GetTaskAndUserDetail()
        {
            try
            {
                DataTable getTaskDetail = new DataTable();
                string TaskQuery = DbRepository.GetAttachmentDetails();
                getTaskDetail = SQLHelper.ExecuteAdapterWithoutParameter(TaskQuery, CommandType.StoredProcedure);
                return getTaskDetail;
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "GetTaskAndUserDetail", ex);
                return null;
            }
        }

        /// <summary> 
        /// Get Task execution details from database.
        /// </summary>
        static DataTable GetTaskExecutionDetails()
        {
            DataTable getTaskExecutionDetails = null;

            try
            {
                string taskQuery = DbRepository.GetTaskExecutionDetails();
                getTaskExecutionDetails = new DataTable();

                getTaskExecutionDetails = SQLHelper.ExecuteAdapterWithoutParameter(taskQuery, CommandType.StoredProcedure);

                return getTaskExecutionDetails;

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "GetTaskExecutionDetails", ex);
                return null;
            }
        }

        /// <summary>
        /// Get Adhoc Task Execution Details
        /// </summary>
        /// <returns></returns>
        static DataTable GetAdhocTaskExecutionDetails()
        {
            DataTable getAdhocTaskExecutionDetails = null;

            try
            {
                string taskQuery = DbRepository.GetAdhocTaskExecutionDetails();
                getAdhocTaskExecutionDetails = new DataTable();

                getAdhocTaskExecutionDetails = SQLHelper.ExecuteAdapterWithoutParameter(taskQuery, CommandType.StoredProcedure);

                return getAdhocTaskExecutionDetails;

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "GetAdhocTaskExecutionDetails", ex);
                return null;
            }
        }


        /// <summary> 
        /// Send emails to users
        /// </summary>
        static string SendEmailToUser(DataTable taskList)
        {
            try
            {
                string result = string.Empty;
                if (taskList != null && taskList.Rows.Count > 0)
                {
                    SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["OutlookServer"]);
                    client.Port = Convert.ToInt32(ConfigurationManager.AppSettings["OutlookServerPort"]);
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;

                    System.Net.NetworkCredential credentials =
                        new System.Net.NetworkCredential(ConfigurationManager.AppSettings["OutlookUser"].ToString(),
                            ConfigurationManager.AppSettings["OutlookPass"].ToString());
                    client.EnableSsl = true;
                    client.Credentials = credentials;

                    //Get distinct task id
                    var dtUniqueTaskList = taskList.AsEnumerable()
                        .Select(row => new
                        {
                            TaskID = row.Field<string>("TaskID"),
                            AttachmentID = row.Field<Int32>("AttachmentID"),
                            ReportName = row.Field<string>("ReportName"),
                            AttachmentName = row.Field<string>("AttachmentName")
                        })
                        .Distinct();

                    MailMessage mail = null;
                    foreach (var list in dtUniqueTaskList)
                    {
                        DataTable userList = taskList.Select("TaskID ='" + list.TaskID.ToString() + "'").CopyToDataTable();
                        mail = new MailMessage();
                        mail.From = new MailAddress(ConfigurationManager.AppSettings["OutlookUser"].ToString().Trim());
                        mail.Subject = list.ReportName + " - " + DateTime.Today.ToString("MM/dd/yyyy");
                        mail.Body = "Please find the attachment";
                        System.Net.Mail.Attachment attachment;
                        attachment = new System.Net.Mail.Attachment(GetAttachment(list.TaskID, list.AttachmentName));
                        mail.Attachments.Add(attachment);
                        foreach (DataRow dr in userList.Rows)
                        {
                            mail.To.Add(new MailAddress(dr["UserEmail"].ToString().Trim()));
                        }
                        client.Send(mail);
                        result = "Success";
                        if (IsAdhoc == false)
                        {
                            UpdateTaskEmailFlag(Convert.ToInt32(list.AttachmentID));

                        }
                        else
                        {
                            AdhocUpdateTaskEmailFlag(Convert.ToInt32(list.AttachmentID));

                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "SendEmailToUser", ex);
                return "";
            }
        }


        /// <summary> 
        /// Get attachment from local directory
        /// </summary>
        static string GetAttachment(string taskID, string attachmentName)
        {
            try
            {
                string attachmentDate = DateTime.Today.ToString("MM/dd/yyyy");
                string attachmentDatePath = AppDomain.CurrentDomain.BaseDirectory + "\\EmailAttachment\\" + attachmentDate.Trim();
                string emailAttPath = attachmentDatePath + "\\" + taskID.Trim();
                var taskAttPath = Directory.GetDirectories(attachmentDatePath, taskID, SearchOption.AllDirectories);
                string emailAttachment = string.Empty;
                if (taskAttPath != null && taskAttPath.Count() > 0)
                {
                    emailAttachment = taskAttPath[0] + "\\" + attachmentName;//.Split('.')[0];

                }
                return emailAttachment;
            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "GetAttachment", ex);
                return "";
            }
        }


        /// <summary> 
        /// Update task email flag in database
        /// </summary>
        static void UpdateTaskEmailFlag(int attachmentID)
        {
            try
            {
                string taskQuery = DbRepository.UpdateTaskEmailFlag();
                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@AttachmentID", attachmentID));

                var result = SQLHelper.ExecuteScalar(taskQuery, CommandType.StoredProcedure, parameters.ToArray());

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "UpdateTaskEmailFlag", ex);
            }
        }

        /// <summary> 
        /// Update Adhoc task email flag in database
        /// </summary>
        static void AdhocUpdateTaskEmailFlag(int attachmentID)
        {
            try
            {
                string taskQuery = DbRepository.AdhocUpdateTaskEmailFlag();
                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@AttachmentID", attachmentID));

                var result = SQLHelper.ExecuteScalar(taskQuery, CommandType.StoredProcedure, parameters.ToArray());

            }
            catch (Exception ex)
            {
                LoggerHelper.ExcpLogger("WindowService Program", "AdhocUpdateTaskEmailFlag", ex);
            }
        }

    }
}
