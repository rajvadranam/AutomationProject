using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RelevantCodes.ExtentReports.Model;

namespace BaseProject.Report
{
    public class ExtentReport
    {
        /// <summary>
        /// Gets or sets Title
        /// </summary>
       public String Details { get; set; }
        public String TestCaseId { get; set; }
        public String PublishedReportPath { get; set; }
        public Exception Exceptions { get; set; }
        
    }
}
