using System.Data;
using BaseProject;
using BaseProject.Report;
using OpenQA.Selenium.Remote;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Http;
using System.IO;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Diagnostics;

namespace BaseProject.Pages
{
	public class Common : BasePage
	{

		/// <summary>
		/// Verifies whether the page is Login or not
		/// </summary>
		/// <param name="Driver">Initialized RemoteWebDriver instance</param>
		public static void VerifyPage(RemoteWebDriver driver, Iteration reporter)
		{
			reporter.Add(new Act("Verify Login Page"));
			//Selenide.WaitForElementVisible(driver);
		}

		/// <summary>
		/// Click Project Account Information Link
		/// </summary>
		public static void ClickAccountTopMenuLink(RemoteWebDriver driver, Iteration reporter, string data)
		{
			Selenide.Click(driver, Locator.Get(LocatorType.XPath, "//a[contains(text(),'" + data + "')]"));
		}

		/// <summary>
		/// Navigates page to specific location
		/// </summary>
		/// <param name="Driver">Initialized RemoteWebDriver instance</param>
		/// <param name="location">Location to navigate</param>
		public static void NavigateTo(RemoteWebDriver driver, Iteration reporter, String location)
		{
			Selenide.NavigateTo(driver, location);
		}

		/// <summary>
		/// Refreshs The Browser
		/// </summary>
		/// <param name="Driver">Initialized RemoteWebDriver instance</param>
		/// <param name="location">Location to navigate</param>
		public static void RefreshBrowser(RemoteWebDriver driver, Iteration reporter)
		{
			Selenide.BrowserRefresh(driver);
		}

		/// <summary>
		/// Browser back
		/// </summary>
		/// <param name="Driver">Initialized RemoteWebDriver instance</param>
		/// <param name="location">Location to navigate</param>
		public static void ClickBackBrowser(RemoteWebDriver driver, Iteration reporter)
		{
			Selenide.BrowserBack(driver);
			System.Threading.Thread.Sleep(5000);
		}

		/// <summary>
		/// Performs login
		/// </summary>
		/// <param name="Driver">Initialized RemoteWebDriver instance</param>
		/// <param name="username">Login Username</param>
		/// <param name="password">Login Password</param>
		public static void Login(RemoteWebDriver driver, Iteration reporter, string username, string password)
		{
			reporter.Add(new Act(String.Format("Set Username {0}, Password {1} and Click Login", username, password)));
			Selenide.WaitForElementVisible(driver, Util.GetLocator("UserNamTxtBox"));
			Selenide.SetText(driver, Util.GetLocator("UserNamTxtBox"), Selenide.ControlType.Textbox, username);
			Selenide.SetText(driver, Util.GetLocator("PasswordTxtBox"), Selenide.ControlType.Textbox, password);
			Selenide.Click(driver, Util.GetLocator("LoginBtn"));
			System.Threading.Thread.Sleep(2000);
			Actions actions = new Actions(driver);
               actions.SendKeys(OpenQA.Selenium.Keys.Escape);
			//var ajaxIsComplete = (bool)(driver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
			((IJavaScriptExecutor)driver).ExecuteScript("return window.stop();");
			var ajaxIsComplete = (bool)(driver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
			//Selenide.WaitForElementVisible(driver, Util.GetLocator("DashBoardOpenTickets"));
		}
		/// <summary>
		/// Clicks [Project Search] on top bar
		/// </summary>
		

		public static void QuickNavLineItem(RemoteWebDriver driver, Iteration reporter, String lineItemId)
		{
			reporter.Add(new Act("Set Quick Navigation Line Item " + lineItemId));
			Selenide.SetText(driver, new Locator(LocatorType.ID, "quickNavInput"), Selenide.ControlType.Textbox, lineItemId);

			if (Selenide.Browser.isIPadSafari(driver))
				Selenide.JS.Enter(driver, new Locator(LocatorType.ID, "quickNavInput"));
			else
				Selenide.Enter(driver, new Locator(LocatorType.ID, "quickNavInput"));
		}

		

		
		///<summary>
		///method to take a screenshot
		///</summary>
		public static void GetScreener(RemoteWebDriver driver, Iteration reporter, string Message)
		{
			ITakesScreenshot iTakeScreenshot = driver;
			reporter.Screenshot = iTakeScreenshot.GetScreenshot().AsBase64EncodedString;

			if (reporter.Screenshot.Length == 0)
			{
				reporter.Add(new Act("unable to add screenshot"));
			}
			else
			{
				reporter.Add(new Act(Message + "<span class='pull-right'><a href='" + Path.Combine("PassedScreenshots", String.Format("{0} {1} {2} Passed.png", reporter.Browser.TestCase.Title, reporter.Browser.Title, reporter.Title)) + "'><span class='glyphicon glyphicon-paperclip normal'></span></a>&nbsp; </span>"));

			}

		}



		///<summary>
		///verify that images in the web page are displayed are not
		///</summary>
		public static void IsImagesDisplayed(RemoteWebDriver driver, Locator locator, Iteration reporter)
		{
			IWebElement element = driver.FindElementByXPath(locator.Location);
			IList<IWebElement> lst = driver.FindElements(By.XPath(locator.Location));
			for (int j = 0; j < lst.Count; j++)
			{
				string sorc = lst[j].GetAttribute("src");
				string name = lst[j].GetAttribute("alt");
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sorc);
				request.Method = "GET";
				try
				{
					using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())

						if (response.StatusCode == HttpStatusCode.OK)
					{
						reporter.Add(new Act("Image displayed  " + name));
					}
					else
					{
						reporter.Add(new Act("Image not displayed" + name));
					}

				}
				catch (Exception e)
				{
					j++;
					reporter.Add(new Act("Image not displayed    " + name));
					if (e.Message.Contains("404") && j == lst.Count)
					{
						throw new Exception("image not loaded");
					}
					continue;

				}


			}

		}

		///<summary>
		///Find a specified file type and fine path in all drives
		///</summary>
		public static string FindfiletypeNpath(String type, String Fname)
		{
			string filepath = "";
			var files = new List<string>();
			DriveInfo[] d = DriveInfo.GetDrives();
			int drivtotal = d.Length;
			if(Fname.Contains("@")|| Fname.Contains("/")|| Fname.Contains("//")|| Fname.Contains("\\"))
			{
				filepath = Path.GetFullPath(Fname+"."+type);
				drivtotal =0;
			}
			for (int i = 0; i <= drivtotal - 1; i++)
			{
				string current = d[i].ToString();
				System.IO.DirectoryInfo dirInfo = d[i].RootDirectory;
				System.IO.DirectoryInfo[] dirInfos = dirInfo.GetDirectories("*.*");
				int subies = dirInfos.Length - 1;

				for (int j = 0; j <= subies; )
				{
					if (dirInfos[j].ToString().Contains("$") || dirInfos[j].ToString().Contains("."))
					{
						j++;
						continue;
					}
					else
					{
						string currentsubby = dirInfos[j].ToString();
						try
						{
							files.AddRange(Directory.GetFiles(current + dirInfos[j].ToString() + "\\", "*." + type, SearchOption.AllDirectories));
							if (files.Count == 0)
							{
								j++;
								continue;
							}
							else
							{

								for (int k = 0; k <= files.Count; k++)
								{
									if (files[k].ToString().Contains(Fname))
									{
										filepath = files[k].ToString();
										break;
									}
								}
							}
						}


						catch (Exception e)
						{
							if (e.Message.Contains("denied"))
							{
								j++;
								continue;
							}

						}
					}
				}
			}
			return filepath;

		}

		///<summary>
		///Find a specified link in page and click it
		///</summary>
		public static void ClickLink(RemoteWebDriver driver, String Name)
		{
			string target2 = null;
			string target = Name.ToLowerInvariant();
			target2 = Regex.Replace(Name.ToLower(), @"\b[a-zA-Z]", m => m.Value.ToUpper());
			String identifierXpath = String.Format("//*[contains(text(),'" + target2 + "')]");
			By identifier = By.XPath(identifierXpath);
			IList<IWebElement> elementsWithText = driver.FindElements(identifier);

			for (int i = 0; i < elementsWithText.Count; i++)
			{
				if (elementsWithText[i].Text.Equals(target2))
				{
					string og = elementsWithText[i].TagName.ToString();
					
					string[] sa = identifierXpath.Split('*');
					string firsthalf = "";
					string secondhalf = "";
					for (int k = 0; k <= sa.Length; k++)
					{
						firsthalf = sa[k].ToString();
						secondhalf = sa[k + 1].ToString();

						break;
					}
					string xpath_link = firsthalf + og + secondhalf;
					driver.FindElement(By.XPath(xpath_link)).Click();
					break;
				}
			}

		}

		///<summary>
		///Find a specified link in page and focus it and clcik submenu
		///</summary>
		public static void ClickLink(RemoteWebDriver driver, String Name, string submenu)
		{
			string og = null;
			string target2 = null;
			string submenus = null;
			string target = Name.ToLowerInvariant();
			target2 = Regex.Replace(Name.ToLower(), @"\b[a-zA-Z]", m => m.Value.ToUpper());
			if(submenu.Length==0)
			{
				submenu = "NOTdefined";
			}
			string submenuss = submenu.ToLowerInvariant();
			submenus = Regex.Replace(submenuss.ToLower(), @"\b[a-zA-Z]", m => m.Value.ToUpper());
			String identifierXpath = String.Format("//*[contains(text(),'" + target2 + "')]");
			String submenuxpath = String.Format("//*[contains(text(),'" + submenus + "')]");
			By identifier = By.XPath(identifierXpath);
			IList<IWebElement> elementsWithText = driver.FindElements(identifier);

			
			for (int i = 0; i < elementsWithText.Count;)
			{
				if (elementsWithText[i].Text.Equals(target2)&&elementsWithText[i].TagName.Length==1)
				{
					og = elementsWithText[i].TagName.ToString();
					char tag = Convert.ToChar(og);
					string xpath_link = identifierXpath.Replace('*', tag);
					IWebElement elems1 = driver.FindElement(By.XPath(xpath_link));
					Actions builder = new Actions(driver);
					Actions hoverOverRegistrar = builder.MoveToElement(elems1);
					hoverOverRegistrar.Perform();
					if (submenu.Equals("NOTdefined"))
					{
						builder.MoveToElement(elems1).Click().Build().Perform();
						i = elementsWithText.Count;
						break;
					}
					i = elementsWithText.Count;
					By identifier2 = By.XPath(submenuxpath);
					IList<IWebElement> elementsWithText2 = driver.FindElements(identifier2);
					if (elementsWithText2.Count==0)
					{
						throw new Exception("Sub menu not displayed check spelling");
					}
					for (int j = 0; j < elementsWithText2.Count; j++)
					{
						if (elementsWithText2[j].Text.Equals(submenus))
						{
							string og2 = elementsWithText2[j].TagName.ToString();
							string[] sa = submenuxpath.Split('*');
							string firsthalf = "";
							string secondhalf="";
							for (int k = 0; k <= sa.Length; k++)
							{
								firsthalf = sa[k].ToString();
								secondhalf = sa[k+1].ToString();

								break;
							}

							//*[contains(text(),'Charge History')]/parent::a/parent::li/parent::ul[contains(@class,'submenu')]
							//ul[contains(@class,'submenu')]/li/a/span[contains(text(),'Charge History')]
							string xpath_link2 = firsthalf + "ul[contains(@class,'submenu')]/li/" + og + "/" + og2 + secondhalf;
							driver.FindElement(By.XPath(xpath_link2)).Click();
						}
					}
				}
				else
				{
					if(!elementsWithText[i].Text.Equals(target2))
					{
						i++;
						continue;
					}
					By identifier2 = By.XPath(submenuxpath);
					IList<IWebElement> elementsWithText2 = driver.FindElements(identifier2);

					for (int j = 0; j < elementsWithText2.Count; j++)
					{
						if (elementsWithText2[j].Text.Equals(submenus))
						{
							string og2 = elementsWithText2[j].TagName.ToString();
							string[] sa = submenuxpath.Split('*');
							string firsthalf = "";
							string secondhalf = "";
							for (int k = 0; k <= sa.Length; k++)
							{
								firsthalf = sa[k].ToString();
								secondhalf = sa[k + 1].ToString();

								break;
							}

							//*[contains(text(),'Charge History')]/parent::a/parent::li/parent::ul[contains(@class,'submenu')]
							//ul[contains(@class,'submenu')]/li/a/span[contains(text(),'Charge History')]
							string xpath_link2  = firsthalf + og2 + secondhalf;
							
							driver.FindElement(By.XPath(xpath_link2)).Click();
							i = elementsWithText.Count;
							break;
						}
					}
				}
			}
		}


        private static void Clean(string folder)
        {
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }

            Directory.CreateDirectory(folder);
        }

        private static byte[] GetFileContent(int i)
        {
            Random r = new Random(i);
            byte[] buffer = new byte[1024];
            r.NextBytes(buffer);
            return buffer;
        }

        private static void CreateWithFileStream(string folder)
        {
            var sw = new Stopwatch();
            sw.Start();

            for (int i = 0; i < 1000; i++)
            {
                string path = Path.Combine(folder, string.Format("file{0}.dat", i));

                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.None))
                {
                    byte[] bytes = GetFileContent(i);
                    fs.Write(bytes, 0, bytes.Length);
                }
            }

            Console.WriteLine("Time for CreateWithFileStream: {0}ms", sw.ElapsedMilliseconds);
        }

        private static void CreateWithFileWriteBytes(string folder)
        {
            var sw = new Stopwatch();
            sw.Start();

            for (int i = 0; i < 1000; i++)
            {
                string path = Path.Combine(folder, string.Format("file{0}.dat", i));
                File.WriteAllBytes(path, GetFileContent(i));
            }

            Console.WriteLine("Time for CreateWithFileWriteBytes: {0}ms", sw.ElapsedMilliseconds);
        }

        public static void CreateWithParallelFileWriteBytes(string folder, string Filename,string data)
        {
            var sw = new Stopwatch();
            sw.Start();

         
                string path = Path.Combine(folder, string.Format("{0}.txt", Filename));
                File.WriteAllBytes(path, Encoding.ASCII.GetBytes(data));
         
            Console.WriteLine("Time for Creating file {0}: {1}ms", Filename,sw.ElapsedMilliseconds);
        }
    }
	public class Excelhandler :BasePage
	{
		
		private struct ExcelDataTypes
		{
			public const string NUMBER = "NUMBER";
			public const string DATETIME = "DATETIME";
			public const string TEXT = "TEXT"; // also works with "STRING".
		}

		private struct NETDataTypes
		{
			public const string SHORT = "int16";
			public const string INT = "int32";
			public const string LONG = "int64";
			public const string STRING = "string";
			public const string DATE = "DateTime";
			public const string BOOL = "Boolean";
			public const string DECIMAL = "decimal";
			public const string DOUBLE = "double";
			public const string FLOAT = "float";
		}

		/// <summary>
		/// Takes file name into consideration ad read the file and returns a data table
		/// </summary>
		/// <param name="FullFilename">EX:Somefile.csv</param>
		/// <param name="column"></param>
		/// <returns>DataTable</returns>
		public static DataTable ReadDocumentContentAndReturnDT(string FullFilename,string column="1")
		{
			string[] temp = FullFilename.Split('.');
			DataTable dt =default(DataTable);
			
			string filepath = Common.FindfiletypeNpath(temp[1].ToString(),@temp[0].ToString());
			switch(temp[1].ToString().ToLower())
			{
					case "csv": dt =Util.ReadCSVContent("",filepath);
					break;
				case "xls":
					dt = ReadContentXls(filepath,column);
					break;
				case "xlsx":
					dt = ReadContentXls(filepath,column);
					break;
				case "xlsm":
					dt = ReadContentXls(filepath,column);
					break;
					
					
			}
			
			return dt;
		}
		
		/// <summary>
		/// Takes file name into consideration ad read the file and returns a data table
		/// </summary>
		/// <param name="FullFilename">EX:Somefile.csv</param>
		/// <param name="column"></param>
		/// <returns>DataSet</returns>
		public static DataSet ReadDocumentContentndReturnDS(string FullFilename,string column="1")
		{
			string[] temp = FullFilename.Split('.');
			DataTable dt =default(DataTable);
			DataSet ds = default(DataSet);
			string filepath = Common.FindfiletypeNpath(temp[1].ToString(),@temp[0].ToString());
			switch(temp[1].ToString().ToLower())
			{
					case "csv": dt =Util.ReadCSVContent("",filepath);
					break;
				case "xls":
					dt = ReadContentXls(column,filepath);
					break;
				case "xlsx":
					dt = ReadContentXls(column,filepath);
					break;
				case "xlsm":
					dt = ReadContentXls(column,filepath);
					break;
					
					
			}
			  
			ds = dt.DataSet;
       
			return ds;
		}
		/// <summary>
		/// Return required dictionary of key values
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="key"></param>
		/// <param name="value"></param>
		/// <returns> Dictionary<string,string></returns>
		public static Dictionary<string,string> RequiredKeyValues(DataTable dt,string key,string value)
		{
			Dictionary<string,string> y = new Dictionary<string, string>();
			try{
				
				foreach(DataRow e in dt.Rows)
				{
					for(int i=0;i<=dt.Rows.Count;i++)
					{
						
						string s = e.Table.Rows[i].ItemArray[1].ToString();
						if(!y.ContainsKey(s))
						{
							y.Add(e.Table.Rows[i].ItemArray[1].ToString(),e.Table.Rows[i].ItemArray[2].ToString());
						}
					}
					
				}
			}
			catch(Exception e)
			{
				
			}
			
			return y;
		}
		
		public static DataTable ReadContentXls(string e,string SHEETNAME=null)
		{
			string sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+e+";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";
			DataSet ds = new DataSet();
			DataTable dt = new DataTable();
			string[] names= BuildExcelSheetNames(sConnection);
			foreach(var sheetName in BuildExcelSheetNames(sConnection))
			{
				using (OleDbConnection con = new OleDbConnection(sConnection))
				{
					
					string query = string.Format("SELECT * FROM [{0}]", sheetName);
					con.Open();
					OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
					adapter.Fill(dt);
					ds.Tables.Add(dt);
					
				}
			}

			
			return dt;
		}
		
	
		
		public static void WriteToExcel(string file,string SHEETNAME, DataTable datatable)
		{
			try{
			string sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+file+";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=3;READONLY=FALSE\"";
			OleDbConnection oleDBCon = new OleDbConnection(sConnection);
			string[] sheets= Excelhandler.BuildExcelSheetNames(sConnection);
					
			Excelhandler.ExportToExcelOleDb(datatable.DataSet,sConnection,file,false);
			}
			catch(Exception e1)
			{
			
			}
		}
		
		
		public static void WriteToExcelMultiplesheets(string file,string SHEETNAME, DataTable datatable,string previousfile)
		{
			try{
			string sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+file+";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=3;READONLY=FALSE\"";
string readconnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+previousfile+";Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";  
			OleDbConnection oleDBCon = new OleDbConnection(sConnection);
			string[] sheets= Excelhandler.BuildExcelSheetNames(readconnection);
					
			Excelhandler.ExportToExcelOleDb(datatable.DataSet,sConnection,file,false);
			}
			catch(Exception e1)
			{
			
			}
		}
		
		
		
		

    /// <summary>
    /// Routine to export a given DataSet to Excel. For each DataTable contained 
    /// in the DataSet the overloaded routine will create a new Excel sheet based 
    /// upon the currently selected DataTable. The proceedure loops through all 
    /// DataRows in the selected DataTable and pushes each one to the specified 
    /// Excel file using ADO.NET and the Access Database Engine (Excel is not a 
    /// prerequisit).
    /// </summary>
    /// <param name="dataSet">The DataSet to be written to Excel.</param>
    /// <param name="connectionString">The connection string.</param>
    /// <param name="fileName">The Excel file name to export to.</param>
    /// <param name="deleteExistFile">Delete existing file?</param>
    public static void ExportToExcelOleDb(DataSet dataSet, string connectionString, 
                                                      string fileName, bool deleteExistFile)
    {
        // Support for existing file overwrite.
        if (deleteExistFile && File.Exists(fileName))
            File.Delete(fileName);
        ExportToExcelOleDb(dataSet, connectionString, fileName);
    }

    /// <summary>
    /// Overloaded version of the above.
    /// </summary>
    /// <param name="dataSet">The DataSet to be written to Excel.</param>
    /// <param name="connectionString">The SqlConnection string.</param>
    /// <param name="fileName">The Excel file name to export to.</param>
    public static bool ExportToExcelOleDb(DataSet dataSet, string connectionString, string fileName)
    {
        try
        {
            // Check for null set.
            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                using (OleDbConnection connection = new OleDbConnection(String.Format(connectionString, fileName)))
                {
                    // Initialise SqlCommand and open.
                    OleDbCommand command = null;
                    connection.Open();

                    // Loop through DataTables.
                    foreach (DataTable dt in dataSet.Tables)
                    {
                        // Build the Excel create table command.
                        string strCreateTableStruct = BuildCreateTableCommand(dt);
                        if (String.IsNullOrEmpty(strCreateTableStruct))
                            return false;
                        command = new OleDbCommand(strCreateTableStruct, connection);
                        command.ExecuteNonQuery();

                        // Puch each row into Excel.
                        for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
                        {
                            command = new OleDbCommand(BuildInsertCommand(dt, rowIndex), connection);
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            return true;
        }
        catch (Exception eX)
        {
           
            return false;
        }
    }

    /// <summary>
    /// Build the various sheet names to be inserted based upon the 
    /// number of DataTable provided in the DataSet. This is not required
    /// for XCost purposes. Coded for completion.
    /// </summary>
    /// <param name="connectionString">The connection string.</param>
    /// <returns>String array of sheet names.</returns>
    private static string[] BuildExcelSheetNames(string connectionString)
    {
        // Variables.
        DataTable dt = null;
        string[] excelSheets = null;

        using (OleDbConnection schemaConn = new OleDbConnection(connectionString))
        {
            schemaConn.Open();
            dt = schemaConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            // No schema found.
            if (dt == null)
                return null;

            // Insert 'TABLE_NAME' to sheet name array.
            int i = 0;
            excelSheets = new string[dt.Rows.Count];
            foreach (DataRow row in dt.Rows)
                excelSheets[i++] = row["TABLE_NAME"].ToString();
        }
        return excelSheets;     
    }

    /// <summary>
    /// Routine to build the CREATE TABLE command. The conversion of 
    /// .NET to Excel data types is also handled here (supposedly!). 
    /// Help: http://support.microsoft.com/kb/316934/en-us.
    /// </summary>
    /// <param name="dataTable"></param>
    /// <returns>The CREATE TABLE command string.</returns>
    private static string BuildCreateTableCommand(DataTable dataTable)
    {
        // Get the type look-up tables.
        StringBuilder sb = new StringBuilder();
        Dictionary<string, string> dataTypeList = BuildExcelDataTypes();

        // Check for null data set.
        if (dataTable.Columns.Count <= 0)
            return null;

        // Start the command build.
        sb.AppendFormat("CREATE TABLE [{0}] (", BuildExcelSheetName(dataTable));

        // Build column names and types.
        foreach (DataColumn col in dataTable.Columns)
        {
            string type = ExcelDataTypes.TEXT;
            if (dataTypeList.ContainsKey(col.DataType.Name.ToString().ToLower()))
            {
                type = dataTypeList[col.DataType.Name.ToString().ToLower()];
            }
            sb.AppendFormat("[{0}] {1},", col.Caption.Replace(' ', '_'), type);
        }
        sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);
        return sb.ToString();   
    }

    /// <summary>
    /// Routine to construct the INSERT INTO command. This does not currently 
    /// work with the data type miss matches.
    /// </summary>
    /// <param name="dataTable"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    private static string BuildInsertCommand(DataTable dataTable, int rowIndex)
    {
        StringBuilder sb = new StringBuilder();

        // Remove whitespace.
        sb.AppendFormat("INSERT INTO [{0}$](", BuildExcelSheetName(dataTable));
        foreach (DataColumn col in dataTable.Columns)
            sb.AppendFormat("[{0}],", col.Caption.Replace(' ', '_'));
        sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);

        // Write values.
        sb.Append("VALUES (");
        foreach (DataColumn col in dataTable.Columns)
        {
            string type = col.DataType.ToString();
            string strToInsert = String.Empty;
            strToInsert = dataTable.Rows[rowIndex][col].ToString().Replace("'", "''");
            sb.AppendFormat("'{0}',", strToInsert);
            //strToInsert = String.IsNullOrEmpty(strToInsert) ? "NULL" : strToInsert;
            //String.IsNullOrEmpty(strToInsert) ? "NULL" : strToInsert);
        }
        sb = sb.Replace(',', ')', sb.ToString().LastIndexOf(','), 1);
        return sb.ToString();
    }

    /// <summary>
    /// Build the Excel sheet name.
    /// </summary>
    /// <param name="dataTable"></param>
    /// <returns></returns>
    private static string BuildExcelSheetName(DataTable dataTable)
    {
        string retVal = dataTable.TableName;
        if (dataTable.ExtendedProperties.ContainsKey("TABLE_NAME_PROPERTY"))
            retVal = dataTable.ExtendedProperties["TABLE_NAME_PROPERTY"].ToString();
        return retVal.Replace(' ', '_');
    }

            /// <summary>
    /// Dictionary for conversion between .NET data types and Excel 
    /// data types. The conversion does not currently work, so I am 
    /// puching all data upto excel as Excel "TEXT" type.
    /// </summary>
    /// <returns></returns>
    private static Dictionary<string, string> BuildExcelDataTypes()
    {
        Dictionary<string, string> dataTypeLookUp = new Dictionary<string, string>();

        // I cannot get the Excel formatting correct here!?
        dataTypeLookUp.Add(NETDataTypes.SHORT, ExcelDataTypes.NUMBER);
        dataTypeLookUp.Add(NETDataTypes.INT, ExcelDataTypes.NUMBER);
        dataTypeLookUp.Add(NETDataTypes.LONG, ExcelDataTypes.NUMBER);
        dataTypeLookUp.Add(NETDataTypes.STRING, ExcelDataTypes.TEXT);
        dataTypeLookUp.Add(NETDataTypes.DATE, ExcelDataTypes.DATETIME);
        dataTypeLookUp.Add(NETDataTypes.BOOL, ExcelDataTypes.TEXT);
        dataTypeLookUp.Add(NETDataTypes.DECIMAL, ExcelDataTypes.NUMBER);
        dataTypeLookUp.Add(NETDataTypes.DOUBLE, ExcelDataTypes.NUMBER);
        dataTypeLookUp.Add(NETDataTypes.FLOAT, ExcelDataTypes.NUMBER);
        return dataTypeLookUp;
    }
	}
}
