using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OTIS_Email_Service.App_Code
{
    class TaskExecution
    {
        public string id { get; set; }
        public string type { get; set; }
        public string task { get; set; }
        public DateTime created { get; set; }
        public DateTime lastupdate { get; set; }
        public DateTime completed { get; set; }
        public float progress { get; set; }
        public string status { get; set; }
        public string result { get; set; }
        public float priority { get; set; }
    }
}
