using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseProject.Report
{
    public class ExceptionReport
    {
       
        public String TestCaseId { get; set; }
        public String PublishedReportPath { get; set; }
        public Exception ExceptionDetails { get; set; }
    }
}
