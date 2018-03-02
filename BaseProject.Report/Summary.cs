using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BaseProject.Report
{
    public class Summary
    {
        private List<TestCase> testCases = new List<TestCase>();

        /// <summary>
        /// Gets TestCases
        /// </summary>
        public List<TestCase> TestCases
        {
            get
            {
                return testCases;
            }
        }


        /// <summary>
        /// Gets Passed Count
        /// </summary>
        public int PassedCount
        {
            get
            {
                return TestCases.FindAll(i => i.IsSuccess == true).Count;
            }
        }

        /// <summary>
        /// Gets Failed Count
        /// </summary>
        public int FailedCount
        {
            get
            {
                return TestCases.FindAll(i => i.IsSuccess == false).Count;
            }
        }

        /// <summary>
        /// Gets IsSuccess
        /// </summary>
        public bool IsSuccess
        {
            get
            {
                return TestCases.FindAll(i => i.IsSuccess == false).Count == 0;
            }
        }

        /// <summary>
        /// Builds Status Statistics with Browser Name as Pivot
        /// </summary>
        /// <returns><see cref="Dictionary<String, Dictionary<String, long>>"/>
        /// <remarks>Browser name as Key, <see cref="Dictionary<String, long>"/> as Value.
        /// Inner Dictionary contains "PASSED", "FAILED" as keys and count as Value
        /// </remarks>
        /// </returns>
        public Dictionary<String, Dictionary<String, long>> GetStatusByBrowser()
        {
            Dictionary<String, Dictionary<String, long>> result = new Dictionary<string, Dictionary<string, long>>();

            foreach (TestCase testCase in this.TestCases)
            {
                foreach (Browser browser in testCase.Browsers)
                {
                    String browserQaulifier = String.Format("{0} {1}", browser.BrowserName, browser.BrowserVersion);
                    if (!result.Keys.Contains(browserQaulifier))
                    {
                        result.Add(browserQaulifier, new Dictionary<string, long>());
                        result[browserQaulifier].Add("PASSED", 0);
                        result[browserQaulifier].Add("FAILED", 0);
                    }

                    result[browserQaulifier]["PASSED"] += browser.PassedCount;
                    result[browserQaulifier]["FAILED"] += browser.FailedCount;
                }
            }

            return result;
        }
    }
}
