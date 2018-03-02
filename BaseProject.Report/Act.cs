using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseProject.Report
{
    public class Act
    {
        private Boolean isSuccess = true;

        /// <summary>
        /// Creates Action instance
        /// </summary>
        /// <param name="title"></param>
        public Act(String title)
        {
            this.Title = title;
            //this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"));
            this.TimeStamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"));
        }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Gets or sets TimeStamp
        /// </summary>
        public DateTime TimeStamp { get; set; }

        /// <summary>
        /// Gets or sets Extra
        /// </summary>
        public String Extra { get; set; }

        /// <summary>
        /// Gets or sets isSuccess
        /// </summary>
        public Boolean IsSuccess
        {
            get
            {
                return isSuccess;
            }
            set
            {
                isSuccess = value;
            }
        }
    }
}