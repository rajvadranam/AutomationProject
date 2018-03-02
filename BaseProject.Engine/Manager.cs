using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Configuration;
using BaseProject.Report;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net.Mail;
using System.Text.RegularExpressions;
namespace BaseProject.Engine
{

    class Manager
	{
		public static List<Object[]> works = new List<object[]>();
		public static bool reportSummary = true;   
		public static int TestsTemp = 0;
        public static bool detailedreportflag = true;
        public static string browserTypeValue = null;
        public static Dictionary<String, String> qualifiedNames = new Dictionary<string, string>();
        //[DllImport("kernel32.dll")]
        //static extern IntPtr GetConsoleWindow();

        //[DllImport("user32.dll")]
        //static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //const int SW_HIDE = 0;
        //const int SW_SHOW = 5;



        public static void Main(string[] args)
		{
            // var handle = GetConsoleWindow();

            // // Hide
            //ShowWindow(handle, SW_HIDE);


            // string workingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";
            string workingDirectory ="/home/nvadran/DataExtractor/";
            Report.Engine reportEngine = new Report.Engine(Util.EnvironmentSettings["ReportsPath"], Util.EnvironmentSettings["Server"]);
			DeleteTempFiles();
            try
            {
                KillPreviousBrowser("phanthom");
            }
            catch(Exception e)
            { 
            }
			Type typeBaseCase = typeof(BaseCase);

			foreach (Type type in Assembly.LoadFrom("BaseProject.Tests.dll").GetTypes().Where(x => x.IsSubclassOf(typeBaseCase))) {
				qualifiedNames.Add(type.Name, type.AssemblyQualifiedName);
			}

			// test cases
			foreach (DataRow eachRow in Util.ReadCSVContent(workingDirectory, Util.EnvironmentSettings["TestSuite"]).Rows) {
				try {
					if (eachRow["Run"].ToString().ToUpper() != "YES")
						continue;

					String testCaseId = eachRow["TestCaseID"].ToString().Trim();
					String testCaseName = eachRow["TestCaseTitle"].ToString().Trim();
					String testCaseRequirementFeature = eachRow["RequirementFeature"].ToString().Trim();

					Report.TestCase testCaseReporter = new Report.TestCase(testCaseId, testCaseName, testCaseRequirementFeature);
					testCaseReporter.Summary = reportEngine.Reporter;
					reportEngine.Reporter.TestCases.Add(testCaseReporter);

					// browsers
					foreach (String browserId in eachRow["Browser"].ToString().Split(new char[] { ';' })) {
						String overRideBrowserId = String.Empty;

						if (!string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString()))
							overRideBrowserId = ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString();

						Report.Browser browserReporter = new Report.Browser(overRideBrowserId != String.Empty ? overRideBrowserId : browserId);
						browserReporter.TestCase = testCaseReporter;
						testCaseReporter.Browsers.Add(browserReporter);

						// iterations
						foreach (DataRow iterationTestData in Util.GetIterationsTestData(Path.Combine(workingDirectory, "Data"), testCaseId).Rows) {
							if (iterationTestData["Run"].ToString().ToUpper() != "YES")
								continue;

							Dictionary<String, String> testData = iterationTestData.Table.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col => iterationTestData.Field<string>(col.ColumnName));

							Dictionary<String, String> browserConfig = Util.GetBrowserConfig(overRideBrowserId != String.Empty ? overRideBrowserId : browserId);
							String iterationId = iterationTestData["TDID"].ToString();
							String defectID = iterationTestData["DefectID"].ToString();

							Report.Iteration iterationReporter = new Report.Iteration(iterationId, defectID);
							iterationReporter.Browser = browserReporter;
							browserReporter.Iterations.Add(iterationReporter);


							works.Add(new Object[] {
							          	browserConfig,
							          	testCaseId,
							          	iterationId,
							          	iterationReporter,
							          	testData,
							          	reportEngine
							          });
						}
					}
				} catch (Exception e) {
					continue;
				}
			}
			//adding hub control
			if (ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString().ToUpper().Contains("HUB")) {
				string[] br = ConfigurationManager.AppSettings.GetValues("DefaultBrowser");
				string browser = br[0];
				Manager.HubbeforeActions(true, browser);
			}
			// Parllel Processing
			Processor(Int32.Parse(ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism")));
			
           reportEngine.PublishException();
            reportEngine.Summarize();

            // generate re-run suite
            StringBuilder suiteContent = new StringBuilder();
            suiteContent.AppendLine("TestCaseID,TestCaseTitle,RequirementFeature,Run,Browser");
            foreach (TestCase eachCase in reportEngine.Reporter.TestCases)
            {
                if (!eachCase.IsSuccess)
                {
                    String browsers = String.Empty;

                    foreach (Browser eachBrowser in eachCase.Browsers)
                    {
                        if (!eachBrowser.IsSuccess)
                        {
                            browsers += String.Format("{0};", eachBrowser.Title);
                        }
                    }

                    browsers = browsers.TrimEnd(new char[] { ';' });
                    suiteContent.AppendLine(String.Format("{0},{1},{2},Yes,{3}", eachCase.Title, eachCase.Name, eachCase.RequirementFeature, browsers));
                }
            }

            String fileName = Path.Combine(reportEngine.ReportPath, "FailedSuite.csv");

            using (StreamWriter output = new StreamWriter(fileName))
            {
                output.Write(suiteContent.ToString());
            }

            //String[] mailConfig = ConfigurationManager.AppSettings.Get("EmailReports").Split(new char[] { ';' });
            //if (mailConfig.Length > 0 && mailConfig[0].ToUpper() == "TRUE")
            //{
               
            //    if(!(mailConfig[1].ToLower().Contains("gmail")||mailConfig[1].ToLower().Contains("outlook")||mailConfig[1].ToLower().Contains("yahoo")))
            //    {
                    
            //    // Compose Email
            //        RestMailGun(reportEngine, mailConfig);
            //    }
            //    else{
            //       bool status= SendThroughService(mailConfig, reportEngine);
            //       if (!status)
            //       {
            //           RestMailGun(reportEngine, mailConfig);
            //       }
            //    }
            //}
            System.Environment.Exit(0);
        }

        public static void RestMailGun(Report.Engine reportEngine, String[] mailConfig)
        {
            // Prepare to send reports via Email
            String zipFilePath = Path.Combine(
               Directory.GetParent(reportEngine.ReportPath).ToString(),
               reportEngine.Timestamp + ".zip");
            try
            {
                ZipFile.CreateFromDirectory(reportEngine.ReportPath, zipFilePath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Zip already exists");
            }

        //    RestClient client = new RestClient();
        //    client.BaseUrl = new Uri("https://api.mailgun.net/v3");
        //    //client.Authenticator = new HttpBasicAuthenticator("api", "key-a790236dd7b2bff508eaf13509a2d72a");
        //    client.Authenticator = new HttpBasicAuthenticator("api", "key-791763ac0781b3909b21c800869e57d9");
        //    RestRequest request = new RestRequest();
        //    // request.AddParameter("domain", "sandbox1190ad81def541bc9dfac73ed08c16b8.mailgun.org", ParameterType.UrlSegment);
        //    request.AddParameter("domain", "sandbox25147b835c954cbe9eb486187117518e.mailgun.org", ParameterType.UrlSegment);
        //    request.Resource = "{domain}/messages";
        //    request.AddParameter("from", "automation@mailgun.org");
        //    request.AddParameter("to", mailConfig[2]);
        //    string cc = mailConfig[3];
        //    if (cc.Contains('@'))
        //    {
        //        request.AddParameter("cc", mailConfig[3]);
        //    }
        //    string bcc = mailConfig[4];
        //    if (bcc.Contains('@'))
        //    {
        //        request.AddParameter("cc", mailConfig[4]);
        //    }
        //    request.AddParameter("subject", "MetTel Automation Report");
        //    request.AddParameter("html", File.ReadAllText(Path.Combine(reportEngine.ReportPath, "summary.html")));
        //    request.AddFile("attachment", Path.Combine(reportEngine.ReportPath, "summary.html"));
        //    request.AddFile("attachment", zipFilePath);
        //    request.Method = Method.POST;
        //    IRestResponse response = client.Execute(request);
        }

        public static bool SendThroughService(string[] mailConfig,Report.Engine reporter)
        {
            string[] serviceAuthDetails = new string[]{};
            bool mailstatus = false;
            serviceAuthDetails = ConfigurationManager.AppSettings.Get("EmailCreds").Split(new char[] { ';' });
           // string encodedpass =null;
            //string encodeduser = null;
            string decodeduser =null;
            string decodedpass = null;
            if (serviceAuthDetails.Length > 0)
            {
                //encodeduser = Base64Encode(serviceAuthDetails[0]);
                //encodedpass = Base64Encode(serviceAuthDetails[1]);
                if (!IsBase64String(serviceAuthDetails[0]))
                {
                    decodeduser = serviceAuthDetails[0];
                    decodedpass = serviceAuthDetails[1];
                }
                else
                {
                    decodeduser = Base64Decode(serviceAuthDetails[0]);
                    decodedpass = Base64Decode(serviceAuthDetails[1]);
                }

            }
            
            if (mailConfig.Length > 0 && mailConfig[0].ToUpper() == "TRUE")
            {
                
                 String zipFilePath = Path.Combine(
                        Directory.GetParent(reporter.ReportPath).ToString(),
                        reporter.Timestamp + ".zip");
                ZipFile.CreateFromDirectory(reporter.ReportPath, zipFilePath);
                string summary = Path.Combine(reporter.ReportPath, "summary.html");
                List<string> attachements = new List<string>() { zipFilePath, summary };
                switch (mailConfig[1].ToLower())
                {
                    case "gmail":mailstatus= sendEMailThroughHotMail("gmail", attachements, mailConfig, decodeduser, decodedpass);
                        break;
                    case "outlook": mailstatus=sendEMailThroughHotMail("outlook", attachements, mailConfig, decodeduser, decodedpass);
                        break;
                    case "yahoo": mailstatus=sendEMailThroughHotMail("yahoo", attachements, mailConfig, decodeduser, decodedpass);
                        break;
                }
            }
            return mailstatus;
        }
        public static bool IsBase64String(string s)
        {
            
            if (string.IsNullOrWhiteSpace(s))
                return false;

            s = s.Trim();
            return (s.Length % 4 == 0) && Regex.IsMatch(s, @"^[a-zA-Z0-9\+/]*={0,3}$", RegexOptions.None);

        }

        public static bool sendEMailThroughHotMail(string mailingservice,List<string> attachement,string[] mailconfig,string decodeuser,string decodepass)
            {
                bool mailstatus = false;
             try
         { 
    //Mail Message
        MailMessage mM = new MailMessage();
        //Mail Address
      
            mM.From = new MailAddress(decodeuser);
   
        //receiver email id
            mM.To.Add(mailconfig[2]);
        if (mailconfig[3].Contains("@"))
        {
            mM.CC.Add(mailconfig[3]);
        }
        if (mailconfig[4].Contains("@"))
        {
            mM.Bcc.Add(mailconfig[3]);
        }
        //subject of the email
        mM.Subject = "MetTel Automation Report :: " + DateTime.Now.ToString("dd-MMM-yyyy");
        //deciding for the attachment
        mM.Attachments.Add(new Attachment(@attachement[0]));
                 mM.Attachments.Add(new Attachment(@attachement[1]));
        //add the body of the email
                 mM.IsBodyHtml = true;
                 mM.Body = File.ReadAllText(Path.Combine(attachement[1]));
       
        //SMTP client
                 if(mailingservice.Contains("outlook"))
                 {
        SmtpClient sC = new SmtpClient("smtp.outlook.com");
        //port number for Hot mail
        sC.Port = 25;
        sC.Credentials = new System.Net.NetworkCredential(decodeuser, decodepass);
        //enabled SSL
        sC.EnableSsl = true;
        //Send an email
      
            sC.Send(mM);
        
        mailstatus = true;
                 }
                 else if (mailingservice.Contains("yahoo"))
                 {
                     SmtpClient sC = new SmtpClient();
                     //your credential will go here
                     sC.Credentials = new System.Net.NetworkCredential(decodeuser, decodepass);
                     //port number to login yahoo server
                     sC.Port = 587;
                     //yahoo host name
                     sC.Host = "smtp.mail.yahoo.com";
                     sC.Credentials = new System.Net.NetworkCredential(decodeuser, decodepass);
                     //enabled SSL
                     sC.EnableSsl = true;
                     //Send an email
                     sC.Send(mM);
                 }
                 else if (mailingservice.Contains("gmail"))
                 {
                     SmtpClient sC = new SmtpClient("smtp.gmail.com");
                     //port number for Gmail mail
                     sC.Port = 587;
                     //credentials to login in to Gmail account
                     sC.Credentials = new System.Net.NetworkCredential(decodeuser, decodepass);
                     //enabled SSL
                     sC.EnableSsl = true;
                     //Send an email
                     sC.Send(mM);
                     mailstatus = true;
                 }


                 
        
    }//end of try block
    catch (Exception ex)
    {
        mailstatus = false;
       
         Console.WriteLine("Error occured with mail client ,it might be due to the credentials provided or mail smtp server");
        
    }//end of catch
             return mailstatus;
}
        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

		
		static void Processor(int maxDegree)
		{
			Parallel.ForEach(works,
			                 new ParallelOptions { MaxDegreeOfParallelism = maxDegree },
			                 work => {
			                 	try {
			                 		ProcessEachWork(work);
			                 	} catch (Exception e) {
			                 		
			                 	}
			                 });
		}

		static void ProcessEachWork(Object[] data)
		{
			Type typeTestCase = Type.GetType(qualifiedNames[data[1].ToString()]); //data[1]: TestCaseId
			// disable once SuggestUseVarKeywordEvident
			BaseCase baseCase = Activator.CreateInstance(typeTestCase) as BaseCase;
			Console.WriteLine("Executing TestScript : {0} on thread : {1}", data[1].ToString(), Thread.CurrentThread.ManagedThreadId);
			Console.WriteLine("------------------------");
			typeTestCase.GetMethod("Execute").Invoke(
				baseCase, data);
		}


		public static void HubbeforeActions(bool status, string browser)
		{
			String[] ipandport = ConfigurationManager.AppSettings.Get("IPandPort").Split(new Char[] { ':' });
			string ip = ipandport[0].Trim();
			string port = ipandport[1].Trim();
			string projectpath = Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\";
			string[] files = Directory.GetFiles(projectpath, "*.jar", SearchOption.AllDirectories);
			string pathtoserver = files[0].ToString();
			string commandexecregister = "java -jar " + " " + pathtoserver + " -role node -hub http://" + ip + ":" + port + "/grid/register -port 8475";
			string startserver = "java -jar " + " " + pathtoserver + " -role hub";
			Manager.KillPreviousCMD();
			//java -jar selenium-server-standalone-2.45.0.jar -role node -hub localhost:4444/grid/register -browser "browserName=chrome,version=ANY,platform=WINDOWS,maxInstances=20" -Dwebdriver.chrome.driver=" " -maxSession 20
			if (status && browser.ToUpper().Contains("FF")) {
				if (browser.ToUpper().Contains("LOCAL")) {

					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					System.Diagnostics.Process.Start("CMD.exe", "/K" + commandexecregister);
					System.Threading.Thread.Sleep(3000);
				} else {
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					
				}

			}
			//realm is local
			if (status && browser.ToUpper().Contains("CHROME")) {

				if (browser.ToUpper().Contains("LOCAL")) {
					//string regesterchrome = "java -jar  " + " " + pathtoserver + " -role node -hub http://"+ ip + ":" + port + "/grid/register/ -browser -Dwebdriver.chrome.driver=" + Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString()+"\\chromedriver.exe -maxSession 20";
					string regesterchrome = "java -jar  " + " " + pathtoserver + " -Dwebdriver.chrome.driver=" + Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\chromedriver.exe -role node -hub http://" + ip + ":4444/grid/register -port 30000";

					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					// System.Diagnostics.Process.Start("CMD.exe", "/K" + commandexecregister);
					System.Diagnostics.Process.Start("CMD.exe", "/K" + regesterchrome);
					System.Threading.Thread.Sleep(3000);

				} else {
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					
				}

			}
			if (status && browser.ToUpper().Contains("SAFARI")) {

				if (browser.ToUpper().Contains("LOCAL")) {
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					System.Diagnostics.Process.Start("CMD.exe", "/K" + commandexecregister);
					System.Threading.Thread.Sleep(3000);

				} else {
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					
				}

			}
			if (status && browser.ToUpper().Contains("IE")) {

				if (browser.ToUpper().Contains("LOCAL")) {
					// java -jar selenium-server.jar -role wd -hub http://localhost:4444/grid/register -browser browserName="internet explorer",platform=WINDOWS ensureCleanSession=true
					string RegisterIE = "java -jar  " + " " + pathtoserver + " -Dwebdriver.internetexplorer.driver=" + Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\IEDriverServer.exe -role node -hub http://" + ip + ":" + port + "/grid/register -port 30000 -browser browserName=internet explorer,version=9.0,platform=WINDOWS ensureCleanSession=true";
					//java -jar selenium-server-standalone-2.20.0.jar -role webdriver -hub http://192.168.1.248:4444/grid/register -browser browserName="internet explorer",version=8.0,platform=WINDOWS -Dwebdriver.internetexplorer.driver=c:\Selenium\InternetExplorerDriver.exe
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					// System.Diagnostics.Process.Start("CMD.exe", "/K" + commandexecregister);
					//System.Diagnostics.Process.Start("CMD.exe", "/K" + RegisterIE);
					System.Threading.Thread.Sleep(3000);
				} else {
					System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
					
				}

			}
			if (status && browser.ToUpper().Contains("UNIT")) {

				string RegiterUnit = "java -jar  " + " " + pathtoserver + " -role node -hub http://" + ip + ":" + port + "/grid/register -browser browserName=htmlunit -port 30000";

				System.Diagnostics.Process.Start("CMD.exe", "/K" + startserver);
				// System.Diagnostics.Process.Start("CMD.exe", "/K" + commandexecregister);
				System.Diagnostics.Process.Start("CMD.exe", "/K" + RegiterUnit);
				
				// selenium-server-standalone-2.38.0.jar -role node -hub http://"localhost":4444/grid/register -browser browserName=htmlunit -port 5558

			}
		}

		public static void KillPreviousCMD()
		{
			Process[] myProcesses = Process.GetProcesses();

			string wintitle = null;
			foreach (Process P in myProcesses) {
				if (P.MainWindowTitle.Contains("C:\\Windows\\System32\\cmd.exe - java  -jar")) {
					Process pe = Process.GetProcessById(P.Id);
					pe.Kill();
					//System.Diagnostics.Process.Start("CMD.exe", "/K" + "taskkill /pid " + P.Id + " /f");
					//pe.Kill();
					pe.WaitForExit();
				}
			}
		}

        public static void KillPreviousBrowser(string browsername)
        {
            Process[] myProcesses = Process.GetProcesses();

            string wintitle = null;
            foreach (Process P in myProcesses)
            {
                if (P.MainWindowTitle.ToLower().Contains(browsername.ToLower()))
                {
                    Process pe = Process.GetProcessById(P.Id);
                    pe.Kill();
                    //System.Diagnostics.Process.Start("CMD.exe", "/K" + "taskkill /pid " + P.Id + " /f");
                    //pe.Kill();
                    pe.WaitForExit();
                }
            }
        }

        public static void DeleteTempFiles()
        {
            try
            {
                Directory.Delete(System.IO.Path.GetTempPath(), true);
            }
            catch (Exception e)
            {
                if (e.Message.Contains("eing used"))
                {

                }
            }
		
	}
   
		
		
		
	}
}

