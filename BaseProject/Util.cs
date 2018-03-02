using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Safari;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Threading.Tasks;
using System.Diagnostics;


namespace BaseProject
{
    public class Util
    {
        private static Dictionary<string, Locator> locators = new Dictionary<string, Locator>();
        private static Dictionary<string, string> commonTestData = new Dictionary<string, string>();
        private static Dictionary<string, string> environmentSettings = new Dictionary<string, string>();

        /// <summary>
        /// Gets settings for current environment
        /// </summary>
        public static Dictionary<string, string> EnvironmentSettings
        {
            get
            {
                String environment = ConfigurationManager.AppSettings.Get("Environment");
                if (environmentSettings.Count > 0) return environmentSettings;
                String[] KeyValue = null;

                lock (environmentSettings)
                {
                    foreach (String setting in ConfigurationManager.AppSettings.Get(environment).Split(new Char[] { ';' }))
                    {
                        KeyValue = setting.Split(new Char[] { '=' }, 2);
                        if (KeyValue.Length > 1)
                        {
                            environmentSettings.Add(KeyValue[0].Trim(), KeyValue[1].Trim());
                        }
                    }
                }
                return environmentSettings;
            }
        }

        /// <summary>
        /// Gets all iteration related test data of specified test case
        /// </summary>
        /// <param name="testCaseId"></param>
        /// <returns></returns>
        public static DataTable GetIterationsTestData(String location, String testCaseId)
        {
            lock (commonTestData)
            {
                if (commonTestData.Count == 0) LoadCommonTestData(location);
            }

            String[] foundFiles = Directory.GetFiles(location, String.Format("{0}.csv", testCaseId), SearchOption.AllDirectories);

            if (foundFiles.Length == 0)
                throw new FileNotFoundException(String.Format("Test Data file not found at {0}", location), String.Format("{0}.csv", testCaseId));

            DataTable tableTestData = ReadCSVContent("", foundFiles[0]);

            foreach (DataRow eachRow in tableTestData.Rows)
            {
                foreach (DataColumn eachColumn in tableTestData.Columns)
                {
                    if (eachRow[eachColumn].ToString().StartsWith("#"))
                    {
                        eachRow[eachColumn] = commonTestData[eachRow[eachColumn].ToString().Replace("#", "")];
                    }
                }
            }
            return tableTestData;
        }

        /// <summary>
        /// Loads Common Test Data from Common.csv
        /// </summary>
        /// <returns></returns>
        public static void LoadCommonTestData(String location)
        {
            String ColumnValue = String.Empty;
            DataTable tableCommonData = ReadCSVContent(location, EnvironmentSettings["CommonData"]);

            foreach (DataRow eachRow in tableCommonData.Rows)
            {
                commonTestData.Add(eachRow["Key"].ToString(), eachRow["Value"].ToString());
            }
        }

        /// <summary>
        /// Loads specified CSV content to DataTable
        /// </summary>
        /// <param name="filename">Filename of CSV</param>
        /// <returns>DataTable</returns>
        public static DataTable ReadCSVContent(String fileDirectory, String filename)
        {
            DataTable table = new DataTable();
            int temp = 0;

            foreach (String fName in filename.Split(','))
            {
                string[] lines = File.ReadAllLines(Path.Combine(fileDirectory, fName));

                if (temp == 0)
                {
                    temp = 1;
                    // identify columns
                    foreach (String columnName in lines[0].Split(new char[] { ',' }))
                    {
                        table.Columns.Add(columnName, typeof(String));
                    }
                }
                foreach (String data in lines.Where((val, index) => index != 0))
                {
                    table.Rows.Add(data.Split(new Char[] { ',' }));
                }
            }
            return table;
        }

        /// <summary>
        /// Utility functions that wraps object repository.
        /// Loads and maintains object locators
        /// </summary>
        /// <param name="name">Name of locator</param>
        /// <returns><see cref="By"/></returns>
        public static Locator GetLocator(String name)
        {
            if (locators.Count == 0)
            {
                lock (locators)
                {

                    // load all for one time
                    XmlDocument objectRepository = new XmlDocument();
                    //TODO: Assume ObjectRepository is always @ exe location. Set project build to deploy it to bin
                    objectRepository.Load(Path.Combine(Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString(), "Objects.xml"));

                    foreach (XmlNode page in objectRepository.SelectNodes("/PageFactory/page"))
                    {
                        foreach (XmlNode eachObject in page.ChildNodes)
                        {
                            Locator locator = null;

                            switch (eachObject.SelectSingleNode("identifyBy").InnerText.ToLower())
                            {
                                case "linktext":
                                    locator = Locator.Get(LocatorType.LinkText, eachObject.SelectSingleNode("value").InnerText);
                                    break;

                                case "id":
                                    locator = Locator.Get(LocatorType.ID, eachObject.SelectSingleNode("value").InnerText);
                                    break;

                                case "xpath":
                                    locator = Locator.Get(LocatorType.XPath, eachObject.SelectSingleNode("value").InnerText);
                                    break;

                                case "classname":
                                    locator = Locator.Get(LocatorType.ClassName, eachObject.SelectSingleNode("value").InnerText);
                                    break;
                                case "cssselector":
                                    locator = Locator.Get(LocatorType.CssSelector, eachObject.SelectSingleNode("value").InnerText);
                                    break;
                            }

                            locators.Add(eachObject.SelectSingleNode("name").InnerText, locator);
                        }
                    }
                }
            }
            return locators[name];
        }

        /// <summary>
        /// Gets Browser related configuration data from App.Config
        /// </summary>
        /// <param name="browserId">Identity of Browser</param>
        /// <returns><see cref="Dictionary<String, String>"/></returns>
        public static Dictionary<String, String> GetBrowserConfig(String browserId)
        {
            Dictionary<String, String> config = new Dictionary<string, string>();
            String[] KeyValue = null;

            foreach (String attribute in ConfigurationManager.AppSettings.Get(browserId).Split(new Char[] { ';' }))
            {
                KeyValue = attribute.Split(new Char[] { ':' });
                config.Add(KeyValue[0].Trim(), KeyValue[1].Trim());
            }
            return config;
        }

        /// <summary>
        /// Prepares RemoteWebDriver basing on configuration supplied
        /// </summary>
        /// <param name="browserConfig"></param>
        /// <returns></returns>
        public static RemoteWebDriver GetDriver(Dictionary<String, String> browserConfig)
        {
            RemoteWebDriver driver = null;

            System.IO.Directory.CreateDirectory(Path.Combine(ConfigurationManager.AppSettings.Get("ReportsDownloadpath").ToString(), "ReportDownloads"));
            string dirdown = Path.Combine(ConfigurationManager.AppSettings.Get("ReportsDownloadpath").ToString(), "ReportDownloads");

            if (browserConfig["target"] == "local")
            {
                if (browserConfig["browser"] == "Firefox")
                {
                    FirefoxProfile profile = new FirefoxProfile();
                    profile.SetPreference("browser.download.folderList", 2);
                    profile.SetPreference("browser.download.dir", dirdown);
                    profile.SetPreference("browser.download.manager.alertOnEXEOpen", false);
                    profile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/msword,application/csv,text/csv,image/png ,image/jpeg, application/pdf, text/html,text/plain,application/octet-stream");
                    profile.SetPreference("browser.download.manager.focusWhenStarting", false);
                    profile.SetPreference("browser.download.useDownloadDir", true);
                    profile.SetPreference("browser.helperApps.alwaysAsk.force", false);
                    profile.SetPreference("browser.download.manager.alertOnEXEOpen", false);
                    profile.SetPreference("browser.download.manager.closeWhenDone", false);
                    profile.SetPreference("browser.download.manager.showAlertOnComplete", false);
                    profile.SetPreference("browser.download.manager.useWindow", false);
                    profile.SetPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
                    profile.SetPreference("pdfjs.disabled", true);
                    driver = new FirefoxDriver(profile);
                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Manage().Window.Maximize();

                }
                else if (browserConfig["browser"] == "IE")
                {
                    //TODO: Get rid of Framework Path
                    driver = new InternetExplorerDriver(Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString());
                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));

                    driver.Manage().Window.Maximize();
                }
                else if (browserConfig["browser"] == "Chrome")
                {

                    DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                    ChromeOptions chrOpts = new ChromeOptions();
                    chrOpts.AddArguments("test-type");
                    chrOpts.AddUserProfilePreference("profile.content_settings.pattern_pairs.*.multiple-automatic-downloads", 1);
                    chrOpts.AddUserProfilePreference("download.default_directory", dirdown);
                    chrOpts.AddArgument("disable-popup-blocking");
                     chrOpts.AddArgument("--disable-extensions");
                     chrOpts.AddArguments("ignore-certificate-errors", "--disable-features");
                    driver = new ChromeDriver(Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString(), chrOpts);
                    //driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds())));
                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Manage().Window.Maximize();

                }
                 if (browserConfig["browser"].ToUpper() == "PHANTHOMJS")
                    {

                        DesiredCapabilities capabilities = new DesiredCapabilities();
                            capabilities  =DesiredCapabilities.PhantomJS();
                        capabilities.SetCapability(CapabilityType.BrowserName, "Opera");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        capabilities.SetCapability(CapabilityType.Platform,new Platform(PlatformType.Any));
                        capabilities.SetCapability(CapabilityType.IsJavaScriptEnabled, true);
                        //capabilities.SetCapability(CapabilityType.Version, "9");
                        driver = new PhantomJSDriver(Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString());
                        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();
                        

                    }


                else if (browserConfig["browser"] == "Safari")
                {


                    SafariOptions opts = new SafariOptions();
                    opts.AddAdditionalCapability("browser.download.dir", dirdown);
                    opts.AddAdditionalCapability("browser.helperApps.neverAsk.saveToDisk", "application/msword,application/csv,text/csv,image/png ,image/jpeg, application/pdf, text/html,text/plain,application/octet-stream");
                    driver = new SafariDriver(opts);
                    driver.Manage().Window.Maximize();

                }



            }
            else if (browserConfig["target"] == "browserstack")
            {
                DesiredCapabilities desiredCapabilities = new DesiredCapabilities();

                String[] bsCredentials = ConfigurationManager.AppSettings.Get("BrowserStackCredentials").Split(new Char[] { ':' });
                desiredCapabilities.SetCapability("browserstack.user", bsCredentials[0].Trim());
                desiredCapabilities.SetCapability("browserstack.key", bsCredentials[1].Trim());

                foreach (KeyValuePair<String, String> capability in browserConfig)
                {
                    if (capability.Key != "target")
                        desiredCapabilities.SetCapability(capability.Key, capability.Value);
                }

                driver = new RemoteWebDriver(new Uri("http://hub.browserstack.com/wd/hub/"), desiredCapabilities);
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                driver.Manage().Cookies.DeleteAllCookies();
            }
            else if (browserConfig["target"] == "HUB")
            {
                if (ConfigurationManager.AppSettings.Get("DefaultBrowser").ToString().ToUpper().Contains("LOCAL"))
                {
                    String[] ipandport = ConfigurationManager.AppSettings.Get("IPandPort").Split(new Char[] { ':' });
                    string ip = ipandport[0].Trim();
                    string port = ipandport[1].Trim();
                    string NodeURl = "http://" + ip + ":" + port + "/wd/hub";

                    if (browserConfig["browser"].ToUpper() == "FIREFOX")
                    {

                        DesiredCapabilities capabilities = new DesiredCapabilities();
                        capabilities = DesiredCapabilities.Firefox();
                        capabilities.SetCapability(CapabilityType.BrowserName, "firefox");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));

                        driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);

                    }
                    if (browserConfig["browser"].ToUpper() == "CHROME")
                    {
                      
                        DesiredCapabilities capabilities = new DesiredCapabilities();
                        capabilities = DesiredCapabilities.Chrome();
                        capabilities.SetCapability(CapabilityType.BrowserName, "chrome");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));

                        driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();


                    }
                    if (browserConfig["browser"].ToUpper() == "SAFARI")
                    {

                        DesiredCapabilities capabilities = DesiredCapabilities.Safari();
                        capabilities.SetCapability(CapabilityType.BrowserName, "Safari");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);
                        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();
                        

                    }
                    if (browserConfig["browser"].ToUpper() == "IE")
                    {

                        DesiredCapabilities capabilities = new DesiredCapabilities();
                            capabilities  =DesiredCapabilities.InternetExplorer();
                        capabilities.SetCapability(CapabilityType.BrowserName, "ie");
                       // capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        capabilities.SetCapability(CapabilityType.Platform,new Platform(PlatformType.Any));
                        //capabilities.SetCapability(CapabilityType.IsJavaScriptEnabled, true);
                        //capabilities.SetCapability(CapabilityType.Version, "9");
                        driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);
                        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();
                        

                    }
                    
                    

                    //unit drivers

                    if (browserConfig["browser"].ToUpper() == "UNIT")
                    {
                        DesiredCapabilities capabilities = DesiredCapabilities.HtmlUnitWithJavaScript();
                        capabilities.SetCapability(CapabilityType.BrowserName, "htmlunit");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        //driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        //driver.Manage().Cookies.DeleteAllCookies();
                        driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);


                    }
                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Manage().Window.Maximize();

                }
                else
                {
                    String[] TargetIP = ConfigurationManager.AppSettings.Get("TargetNode").Split(new Char[] { ':' });
                    string Tip = TargetIP[0].Trim();
                    string Tport = TargetIP[1].Trim();
                    string TNodeURl = "http://" + Tip + ":" + Tport + "/wd/hub";

                    if (browserConfig["browser"].ToUpper() == "FIREFOX")
                    {

                        DesiredCapabilities capabilities = new DesiredCapabilities();
                        capabilities = DesiredCapabilities.Firefox();
                        capabilities.SetCapability(CapabilityType.BrowserName, "firefox");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));

                        driver = new RemoteWebDriver(new Uri(TNodeURl), capabilities);

                    }
                    if (browserConfig["browser"].ToUpper() == "CHROME")
                    {
                        // ChromeOptions opt = new ChromeOptions();
                        // opt.AddExtension(Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\chromedriver.exe");

                        // DesiredCapabilities capabilities = DesiredCapabilities.Chrome();
                        // //capabilities.SetCapability("webdriver.chrome.driver", Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\chromedriver.exe");
                        // capabilities.SetCapability(CapabilityType.BrowserName, "chrome");
                        // ////capabilities.SetCapability(CapabilityType.Platform, "ANY");
                        // capabilities.SetCapability(CapabilityType.Version, "2.4");
                        //capabilities.SetCapability("webdriver.chrome.driver", Directory.GetParent(Assembly.GetEntryAssembly().Location).ToString() + "\\chromedriver.exe");
                        // capabilities.SetCapability(ChromeOptions.Capability, opt);

                        // //driver = new RemoteWebDriver(new Uri(NodeURl), capabilities);
                        DesiredCapabilities capabilities = new DesiredCapabilities();
                        capabilities = DesiredCapabilities.Chrome();
                        capabilities.SetCapability(CapabilityType.BrowserName, "chrome");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));

                        driver = new RemoteWebDriver(new Uri(TNodeURl), capabilities);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();


                    }
                    if (browserConfig["browser"].ToUpper() == "SAFARI")
                    {

                        DesiredCapabilities capabilities = DesiredCapabilities.Safari();
                        capabilities.SetCapability(CapabilityType.BrowserName, "Safari");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();
                        driver = new RemoteWebDriver(new Uri(TNodeURl), capabilities);

                    }
                    if (browserConfig["browser"].ToUpper() == "IE")
                    {

                        DesiredCapabilities capabilities = DesiredCapabilities.InternetExplorer();
                        capabilities.SetCapability(CapabilityType.BrowserName, "IE");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromMinutes(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        driver.Manage().Cookies.DeleteAllCookies();
                        driver = new RemoteWebDriver(new Uri(TNodeURl), capabilities);

                    }

                    ////unit drivers

                    if (browserConfig["browser"].ToUpper() == "UNIT")
                    {
                        DesiredCapabilities capabilities = DesiredCapabilities.HtmlUnitWithJavaScript();
                        capabilities.SetCapability(CapabilityType.BrowserName, "htmlunit");
                        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                        //driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementPageLoad"))));
                        //driver.Manage().Cookies.DeleteAllCookies();
                        driver = new RemoteWebDriver(new Uri(TNodeURl), capabilities);


                    }
                }
                    
                    driver.Manage().Window.Maximize();
                }

               
            
            return driver;
        }

        /// <summary>
        /// Replaces first occurence
        /// </summary>
        /// <param name="s"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public static string ReplaceFirstOccurrence(string s, string oldValue, string newValue)
        {
            int i = s.IndexOf(oldValue);
            return s.Remove(i, oldValue.Length).Insert(i, newValue);
        }



       
    }
}