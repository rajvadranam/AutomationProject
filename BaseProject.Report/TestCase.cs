using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseProject.Report
{
    public class TestCase
    {
        private List<Browser> browsers = new List<Browser>();

        /// <summary>
        /// Creates a new Chapter
        /// </summary>
        /// <param name="title">Title</param>
        public TestCase(String id, String name, String requirementFeature)
        {
            this.Title = id;
            this.Name = name;
            this.RequirementFeature = requirementFeature;
        }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Gets or sets Test Case Name
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Gets or sets Test Case Name
        /// </summary>
        public String RequirementFeature { get; set; }

        /// <summary>
        /// Gets Browsers
        /// </summary>
        public List<Browser> Browsers
        {
            get
            {
                return browsers;
            }
        }

        /// <summary>
        /// Gets current Browser
        /// </summary>
        public Browser Browser
        {
            get
            {
                return Browsers.Last();
            }
        }

        /// <summary>
        /// Gets Passed Count
        /// </summary>
        public int PassedCount
        {
            get
            {
                return Browsers.FindAll(i => i.IsSuccess == true).Count;
            }
        }

        /// <summary>
        /// Gets Failed Count
        /// </summary>
        public int FailedCount
        {
            get
            {
                return Browsers.FindAll(i => i.IsSuccess == false).Count;
            }
        }

        /// <summary>
        /// Gets IsSuccess
        /// </summary>
        public bool IsSuccess
        {
            get
            {
                return Browsers.FindAll(i => i.IsSuccess == false).Count == 0;
            }
        }

        /// <summary>
        /// Gets or sets BugInfo
        /// </summary>
        public String BugInfo { get; set; }

        public Summary Summary { get; set; }
    }
}