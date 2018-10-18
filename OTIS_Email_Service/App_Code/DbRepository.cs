using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace OTIS_Email_Service.App_Code
{
    public class DbRepository
    {
        /// <summary>
        /// Save task details in db
        /// </summary>
        /// <returns></returns>
        public static string SaveTaskDetail()
        {
            return  "App.SaveEmailTaskDetail";
          
        }

        /// <summary>
        /// Get User and Task Details
        /// </summary>
        /// <returns></returns>
        public static string GetAttachmentDetails()
        {
           return "App.GetAttachmentDetails";
        }

        /// <summary>
        /// Get TaskExecution Details
        /// </summary>
        /// <returns></returns>
        public static string GetTaskExecutionDetails()
        {
            return "App.ExecuteNprintingTask";
        }

        /// <summary>
        /// Update Task Email Flag
        /// </summary>
        /// <returns></returns>
        public static string UpdateTaskEmailFlag()
        {
            return "App.UpdateTaskEmailFlag";
        }


        /// <summary>
        /// Update AdhocTask Email Flag
        /// </summary>
        /// <returns></returns>
        public static string AdhocUpdateTaskEmailFlag()
        {
            return "App.AdhocUpdateTaskEmailFlag";
        }


        /// <summary>
        /// Get Adhoc TaskExecution Details
        /// </summary>
        /// <returns></returns>
        public static string GetAdhocTaskExecutionDetails()
        {
            return "APP.ExecutNprintingTaskAdhoc";

        }
    }
}
