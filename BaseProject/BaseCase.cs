using System.Runtime.Remoting.Messaging;
using BaseProject.Report;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Management;
using System.Linq;
using System.Reflection;
using System.IO;
using RelevantCodes.ExtentReports;
using RelevantCodes.ExtentReports.Model;

namespace BaseProject
{
    /// <summary>
    /// Abstracts Common Test Case functionality. 
    /// Derived classes should implement specifics
    /// </summary>
    public abstract class BaseCase
    {
    	
        /// <summary>
        /// Gets or Sets Driver
        /// </summary>
        public RemoteWebDriver Driver { get; set; }

        /// <summary>
        /// Gets or Sets Reporter
        /// </summary>
        public Report.Iteration Reporter { get; set; }

        /// <summary>
        /// Gets or Sets Step
        /// </summary>
        protected string Step
        {
            get
            {
                //TODO: Get should go away
                return Reporter.Chapter.Step.Title;
            }
            set
            {
                Reporter.Add(new Report.Step(value));
            }
        }

        /// <summary>
        /// Gets or Sets Identity of Test Case
        /// </summary>
        public string TestCaseId { get; set; }

        
       
        /// <summary>
        /// Gets or Sets Identity of Test Data
        /// </summary>
        public string TestDataId { get; set; }

        /// <summary>
        /// Gets or Sets Test Data as Dictionary<string, string>
        /// </summary>
        public Dictionary<string, string> TestData { get; set; }

        /// <summary>
        /// Executes Test Cases
        /// </summary>
        public void Execute(Dictionary<String, String> browserConfig,
            String testCaseId,
            String iterationId,
            Report.Iteration iteration,
            Dictionary<String, String> testData,
            Report.Engine reportEngine)
        {
            try
            {
                this.Driver = Util.GetDriver(browserConfig);
                this.Reporter = iteration;
                this.TestCaseId = testCaseId;
                this.TestDataId = iterationId;
                this.TestData = testData;

                if (browserConfig["target"] == "local")
                {
                    var wmi = new ManagementObjectSearcher("select * from Win32_OperatingSystem").Get().Cast<ManagementObject>().First();

                    this.Reporter.Browser.PlatformName = String.Format("{0} {1}", ((string)wmi["Caption"]).Trim(), (string)wmi["OSArchitecture"]);
                    this.Reporter.Browser.PlatformVersion = ((string)wmi["Version"]);
                    this.Reporter.Browser.BrowserName = Driver.Capabilities.BrowserName;
                    this.Reporter.Browser.BrowserVersion = Driver.Capabilities.Version;
                }
                else
                {
                    this.Reporter.Browser.PlatformName = browserConfig.ContainsKey("os") ? browserConfig["os"] : browserConfig["device"];
                    this.Reporter.Browser.PlatformVersion = browserConfig.ContainsKey("os_version") ? browserConfig["os_version"] : browserConfig.ContainsKey("realMobile") ? "Real" : "Emulator";
                    this.Reporter.Browser.BrowserName = browserConfig.ContainsKey("browser") ? browserConfig["browser"] : "Safari";
                    this.Reporter.Browser.BrowserVersion = browserConfig.ContainsKey("browser_version") ? browserConfig["browser_version"] : "";
                }

                // By now, I get all needed stuff
                // Let Page get reference to Driver, Reporter
                //PrepareSeed();
//Extent Experimental
                // Does Seed having anything?
                if (this.Reporter.Chapter.Steps.Count == 0)
                    this.Reporter.Chapters.RemoveAt(0);
                this.Reporter.StartTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"));
                ExecuteTestCase();
               }
            catch (Exception ex)
            {
                
                this.Reporter.Chapter.Step.Action.IsSuccess = false;
                //reportEngine.Extent.Add(new Report.ExtentReport {TestCaseId = testCaseId , PublishedReportPath = String.Format("{0} {1} {2}.html", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title)});
                reportEngine.MetTelExceptions.Add(new Report.ExceptionReport { ExceptionDetails = ex, PublishedReportPath = String.Format("{0} {1} {2}.html", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title), TestCaseId = testCaseId });
                this.Reporter.Chapter.Step.Action.Extra = "Exception Message :<font color=red>" + ex.Message + "</font><br/>" + ex.InnerException + ex.StackTrace;           
                
            }
            finally
            {
                try
                {

                    this.Reporter.IsCompleted = true;
                   
                    // If current iteration is a failure, get screenshot
                    if (!Reporter.IsSuccess)
                    {
                        ITakesScreenshot iTakeScreenshot = Driver;
                        this.Reporter.Screenshot = iTakeScreenshot.GetScreenshot().AsBase64EncodedString;
                    }
                    else
                    {
                        ITakesScreenshot iTakeScreenshot = Driver;
                        this.Reporter.Screenshot = iTakeScreenshot.GetScreenshot().AsBase64EncodedString;
                    }
                    this.Reporter.EndTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"));

                    lock (reportEngine)
                    {
                    	reportEngine.PublishIteration(this.Reporter);
                    	reportEngine.PublishExtent(this.Reporter);
                        reportEngine.Summarize(false);
                    }

                  
                    Driver.Quit();
                }
                catch (Exception e)
                {
                    //throw new Exception("Exception is : " + e);
                    //ITakesScreenshot iTakeScreenshot = Driver;
                    //this.Reporter.Screenshot = iTakeScreenshot.GetScreenshot().AsBase64EncodedString;
                    this.Reporter.EndTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"));
                    reportEngine.Extent.Add(new Report.ExtentReport {TestCaseId = testCaseId , PublishedReportPath = String.Format("{0} {1} {2}.html", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title)});
                    lock (reportEngine)
                    {
                    	reportEngine.PublishIteration(this.Reporter);
                    	reportEngine.PublishExtent(this.Reporter);
                        reportEngine.Summarize(false);
                    }
                    Driver.Quit();
                }
            }
        }

        /// <summary>
        /// Executes Test Case, should be overriden by derived
        /// </summary>
        protected virtual void ExecuteTestCase()
        {
            Reporter.Add(new Report.Chapter("Execute Test Case"));
        }

        /// <summary>
        /// Prepares Seed Data, should be overriden by derived
        /// </summary>
        protected virtual void PrepareSeed()
        {
            Reporter.Add(new Report.Chapter("Prepare Seed Data"));
        }

        /// <summary>
        /// Gets or Sets Chapter
        /// </summary>
        protected string Chapter
        {
            get
            {
                //TODO: Get should go away
                return Reporter.Chapter.Title;
            }
            set
            {
                Reporter.Add(new Report.Chapter(value));
            }
        }

        // Get Attachments Directory
        public static string AttachmentsDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path) + "\\Data\\Attachments\\";
            }
        }
        
//        
//         public void Validate(String result) {
//		  	if (result.Contains("Passed")) {
//			ExtentTestCaseId.Log(LogStatus.Pass, "Test passed");
//		}
//		else if (result.Contains("Failed")) {
//			ExtentTestCaseId.Log(LogStatus.Fail, "Test failed");
//		}
////		else if (result.getStatus() == ITestResult.SKIP) {
////			ExtentTestManager.getTest().log(LogStatus.SKIP, "Test skipped");
////		}
//		
//		//ExtentTestManager.endTest();
//		ExtentManager.getInstance().Flush();
//	}
    }
}
    