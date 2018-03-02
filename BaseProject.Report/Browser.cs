using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseProject.Report
{
    public class Browser
    {
        private List<Iteration> iterations = new List<Iteration>();

        /// <summary>
        /// Creates a new Browser
        /// </summary>
        /// <param name="title">Title</param>
        public Browser(String title)
        {
            this.Title = title;
        }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Gets or sets Browser Name
        /// </summary>
        public String BrowserName { get; set; }

        /// <summary>
        /// Gets or sets Browser Version
        /// </summary>
        public String BrowserVersion { get; set; }

        /// <summary>
        /// Gets or sets Platform Name
        /// </summary>
        public String PlatformName { get; set; }

        /// <summary>
        /// Gets or sets Platform Version
        /// </summary>
        public String PlatformVersion { get; set; }

        /// <summary>
        /// Gets Iterations
        /// </summary>
        public List<Iteration> Iterations
        {
            get
            {
                return iterations;
            }
        }

        /// <summary>
        /// Gets current Iteration
        /// </summary>
        public Iteration Iteration
        {
            get
            {
                return Iterations.Last();
            }
        }

        /// <summary>
        /// Gets Passed Count
        /// </summary>
        public int PassedCount
        {
            get
            {
                return Iterations.FindAll(i => i.IsSuccess == true && i.IsCompleted == true).Count;
            }
        }

        /// <summary>
        /// Gets Failed Count
        /// </summary>
        public int FailedCount
        {
            get
            {
                return Iterations.FindAll(i => i.IsSuccess == false && i.IsCompleted == true).Count;
            }
        }

        /// <summary>
        /// Gets IsSuccess
        /// </summary>
        public bool IsSuccess
        {
            get
            {
                return Iterations.FindAll(i => i.IsSuccess == false && i.IsCompleted == true).Count == 0;
            }
        }

        public TestCase TestCase { get; set; }
    }
}
