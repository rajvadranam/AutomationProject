using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using HtmlAgilityPack;
using RelevantCodes.ExtentReports;
using RelevantCodes.ExtentReports.Model;

namespace BaseProject.Report
{
    public class Engine
    {
        private String reportsPath = String.Empty;
        private String serverName = String.Empty;
        private String timestamp = String.Empty;
       private String extentpath = String.Empty;
        private Object _provisionalSummaryLocker = new Object();
        public List<ExceptionReport> MetTelExceptions = new List<ExceptionReport>();
         public List<ExtentReport> Extent = new List<ExtentReport>();

        Summary summary = new Summary();

        /// <summary>
        /// Gets Report Path
        /// </summary>
        public String ReportPath
        {
            get
            {
                return reportsPath;
            }
        }
        
          public String Extentpath
        {
            get
            {
                return extentpath;
            }
        }

        /// <summary>
        /// Gets Reports TimeStamp
        /// </summary>
        public String Timestamp
        {
            get
            {
                return timestamp;
            }
        }

        /// <summary>
        /// Gets Server name
        /// </summary>
        public String ServerName
        {
            get
            {
                return serverName;
            }
        }

        /// <summary>
        /// Gets or sets Reporter
        /// </summary>
        public Summary Reporter
        {
            get
            {
                return summary;
            }
        }

        /// <summary>
        /// Creates Engine instance
        /// </summary>
        /// <param name="resultPath">Path to Report Results</param>
        public Engine(String resultPath, String serverName)
        {
            this.serverName = serverName;


            try
            {
                this.timestamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")).ToString("MM-dd-yyyy HH-mm-ss");
            }
            catch (TimeZoneNotFoundException)
            {
                this.timestamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("America/New_York")).ToString("MM-dd-yyyy HH-mm-ss");
            }
            //this.timestamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")).ToString("MM-dd-yyyy HH-mm-ss");
           // this.timestamp = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")).ToString("MM-dd-yyyy HH-mm-ss");
            this.reportsPath = Path.Combine(resultPath, this.timestamp);
            this.extentpath = Path.Combine(resultPath,"Extent_"+this.timestamp);
            System.IO.Directory.CreateDirectory(this.reportsPath);
            System.IO.Directory.CreateDirectory(this.Extentpath);
            System.IO.Directory.CreateDirectory(Path.Combine(this.reportsPath, "Screenshots"));

        }

        /// <summary>
        /// Publishes Summary Report of an iteration
        /// </summary>
        public void PublishIteration(Iteration iteration)
        {
            // If current iteration is a failure, get screenshot
            if (!iteration.IsSuccess)
            {
                try
                {
                    File.WriteAllBytes(Path.Combine(this.reportsPath, "Screenshots", String.Format("{0} {1} {2} Error.png", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title)),
                        Convert.FromBase64String(iteration.Screenshot));
                }
                catch (Exception)
                {
                }
            }

            #region Write HTML Content

            String template = @"
            <html>
            <head>
	        <link href='http://netdna.bootstrapcdn.com/bootstrap/3.1.1/css/bootstrap.min.css' rel='stylesheet'>
	        <script src='http://code.jquery.com/jquery-1.11.0.min.js' type='text/javascript'></script>
	        <script src='http://netdna.bootstrapcdn.com/bootstrap/3.1.1/js/bootstrap.min.js'></script>	
	        <style>
            html {
	        overflow: -moz-scrollbars-vertical; /* Vertical Scroll bar always visible, to avoid flicker while collapse/expand */
	        overflow-y: scroll;
	        }
            .bigger-icon {
		        transform:scale(3.0,3.0);
		        -ms-transform:scale(3.0,3.0); /* IE 9 */
		        -moz-transform:scale(3.0,3.0); /* Firefox */
		        -webkit-transform:scale(3.0,3.0); /* Safari and Chrome */
		        -o-transform:scale(3.0,3.0); /* Opera */}
             .default {
		
		            font-family: Courier New;
					font-size: 15px;
	            }
	        .Report-Chapter { 
                padding:12px; margin-bottom: 5px;		
		        background-color: #26466D; color: #fff;
		        font-size: 90%; font-weight:bold;
		        border: 1px solid #03242C; border-radius: 4px;
		        font-family: Menlo,Monaco,Consolas,'Courier New',monospace; cursor: pointer; }	
	        .Report-Step {
		        padding:12px; margin-bottom: 5px;		
		        background-color: #ddd; color: #000;
		        font-size: 90%; font-weight:bold;
		        border: 1px solid #bebebe; border-radius: 4px;
		        font-family: Menlo,Monaco,Consolas,'Courier New',monospace; cursor: pointer;}	
	        .Report-Action {
		        padding:12px; margin-bottom: 5px;		
		        background-color: #f7f7f9; color: #000; font-size: 90%;
		        border: 1px solid #e1e1e8; border-radius: 4px;
		        font-family: Menlo,Monaco,Consolas,'Courier New',monospace;}	
	        .green { color:green; }
	        .red {color: red; }
	        .normal {color: black; }
            .brightgreen {color:lime;}
	        .brightred {color: orangered;}
            .darkbg {background-image:url('https://passportplus2btraining.freemanco.com/Content/img/freeman_header_bkg.jpg');}
            .timestamp {color:#555;}
	        </style>
            <script language='javascript'>
            	$(function() {
		            $('.Report-Chapter').click(function(){
			            $(this).parent().children('.wrapper').slideToggle();
		            });
		
		            $('.Report-Step').click(function(){
			            $(this).parent().children('.Report-Action').slideToggle();
		            });
	            });
            </script>
            </head>
            <body>
	            <div class='container'>
		            <div style='padding-top: 5px; padding-bottom:5px;'>
			            <img src='https://upload.wikimedia.org/wikipedia/en/2/20/Big_Data_Logo.png' style='padding-top:20px; width:200px; height:auto;' />
			            <div class='pull-right'><img src='https://omni.okstate.edu/_resources/images/logo-banner.png'/></div>
			
		            </div>
	            </div>
	            <div class='container default'>
                    <div class='darkbg' style='background-color:#26466D; color:#fff; min-height:100px; padding:20px; margin-bottom:5px; margin-top:5px; top:-40px;'>
		                <div class='row'>		                  
		                  <div class='col-md-6' > <b> Server: </b>{{SERVER}}<br/> <b> Browser: </b>{{BROWSER}}<br/> <b> Environment: </b>{{ENVIRONMENT}}</div>
		                  <div class='col-md-6' > <b> Start: </b>{{EXECUTION_BEGIN}}<br/> <b> End: </b>{{EXECUTION_END}}<br/><b> Duration: </b>{{EXECUTION_DURATION}}</div>		  
		                </div>		
	                </div>
                </div>
                <div class='container default'>
                    <div class='darkbg' style='background-color:#26466D; color:#fff; min-height:60px; padding:20px; margin-bottom:5px; margin-top:5px; top:-20px;'>
		                <div class='row'>
                          <div class='col-md-3' > <b> {{TCID}} </b> </div>
                          <div class='col-md-8' > <b> {{TC_NAME}} </b> </div>
                          <div class='col-md-1' > <span class='glyphicon glyphicon-{{STATUS_ICON}} bigger-icon' style='padding-left:10px;'></span>  </div>                          
                        </div>		
	                </div>
                </div>
                <div class='container'>
                    {{CONTENT}}
                </div>
            </body>
            </html>";
            #endregion

            StringBuilder builder = new StringBuilder();

            foreach (Chapter chapter in iteration.Chapters)
            {
                builder.AppendFormat("<div><p class='Report-Chapter'>Chapter: {0}<span class='pull-right'><span class='glyphicon glyphicon-{1}'></span></span></p>", chapter.Title, chapter.IsSuccess ? "ok brightgreen" : "remove brightred");

                foreach (Step step in chapter.Steps)
                {
                    builder.AppendFormat("<div class='wrapper'><p class='Report-Step'>Step: {0}<span class='pull-right'><span class='glyphicon glyphicon-{1}'></span></span></p>", step.Title, step.IsSuccess ? "ok green" : "remove red");

                    foreach (Act action in step.Actions)
                    {
                        builder.AppendFormat("<p class='Report-Action' style='display:none;'>{0}<span class='pull-right'><span class='timestamp'>{1}</span>&nbsp;&nbsp; ", action.Title, action.TimeStamp.ToString("H:mm:ss"));
                        if (action.IsSuccess)
                        {
                            builder.Append("<span class='glyphicon glyphicon-ok green'></span>");
                        }
                        else
                        {
                            builder.Append("<a href='" + Path.Combine("Screenshots", String.Format("{0} {1} {2} Error.png", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title)) + "'><span class='glyphicon glyphicon-paperclip normal'></span></a>&nbsp;");
                            builder.Append("<span class='glyphicon glyphicon-remove red'></span>");
                        }

                        builder.Append("</span></p>");
                    }

                    builder.Append("</div>");
                }

                builder.Append("</div>");
            }

            if (!iteration.IsSuccess)
            {
                //builder.AppendFormat("<div class='default'><p>URL: {0}</p></div>", driver.Url);
                builder.AppendFormat("<div class='default'><p>{0}</p></div>", iteration.Chapter.Step.Action.Extra);
            }

            template = template.Replace("{{STATUS_ICON}}", iteration.IsSuccess ? "ok brightgreen" : "remove brightred");
            template = template.Replace("{{TCID}}", iteration.Browser.TestCase.Title);
            template = template.Replace("{{TC_NAME}}", iteration.Browser.TestCase.Name);
            template = template.Replace("{{SERVER}}", this.ServerName);
            template = template.Replace("{{BROWSER}}", String.Format("{0} {1}", iteration.Browser.BrowserName, iteration.Browser.BrowserVersion));
            template = template.Replace("{{ENVIRONMENT}}", String.Format("{0} {1}", iteration.Browser.PlatformName, iteration.Browser.PlatformVersion));
            //template = template.Replace("{{EXECUTION_BEGIN}}", String.Format("{0} {1}", iteration.StartTime.ToString("MM-dd-yyyy HH:mm:ss"), "Central Time (US & Canada)"));
            //template = template.Replace("{{EXECUTION_END}}", String.Format("{0} {1}", iteration.EndTime.ToString("MM-dd-yyyy HH:mm:ss"), "Central Time (US & Canada)"));
            template = template.Replace("{{EXECUTION_BEGIN}}", iteration.StartTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_END}}",  iteration.EndTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_DURATION}}", iteration.EndTime.Subtract(iteration.StartTime).ToString());

            String fileName = Path.Combine(this.reportsPath, String.Format("{0} {1} {2}.html", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title));

            using (StreamWriter output = new StreamWriter(fileName))
            {
                output.Write(template.Replace("{{CONTENT}}", builder.ToString()));
            }
        }
        
        
         /// <summary>
        /// Publishes Summary Report of an iteration
        /// </summary>
      


        public void PublishExtent(Iteration iteration)
        {
        	var ExtentTest = new ExtentReports(this.Extentpath+"//Extent.html",false);
        	
        	 DateTime FirstCaseBeginTime = DateTime.Now;
            DateTime LastCaseEndTime = DateTime.Now;
            TimeSpan ExecutionTimeCumulative = TimeSpan.Zero;
        	ExtentTest.AddSystemInfo("reportName","DataExtractor Report");
        	Dictionary<string,string> sysInfo = new Dictionary<string,string>();
        	sysInfo.Add("Selenium Version", "2.53");
        	sysInfo.Add("Environment",this.ServerName);
        	sysInfo.Add("Browser", iteration.Browser.BrowserName+ iteration.Browser.BrowserVersion);
        	sysInfo.Add("ThreadsCount", ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism"));
        	
        	try{
    
        		 foreach (Iteration iteration2 in iteration.Browser.Iterations.FindAll(itr => itr.IsCompleted == true))
                    {
                        
                        if (iteration2.StartTime < FirstCaseBeginTime) FirstCaseBeginTime = iteration2.StartTime;
                        if (iteration2.EndTime > LastCaseEndTime) LastCaseEndTime = iteration2.EndTime;
                        ExecutionTimeCumulative = ExecutionTimeCumulative.Add(iteration2.EndTime.Subtract(iteration2.StartTime));
                    }
                
        		 sysInfo.Add("Total Test Duration", "<b>"+ExecutionTimeCumulative.ToString()+"</b>");
        		 sysInfo.Add("FirstTestBeginTime", "<b>"+FirstCaseBeginTime.ToString("MM-dd-yyyy HH:mm:ss")+"</b>");
        		 sysInfo.Add("Last Test END Duration", "<b>"+LastCaseEndTime.ToString("MM-dd-yyyy HH:mm:ss")+"</b>");
        	    ExtentTest.AddSystemInfo(sysInfo);


        		foreach (Chapter chapter in iteration.Chapters)
        		{
        			var test = ExtentTest.StartTest(chapter.Title);
        		
        			test.StartTime = iteration.StartTime;
        			test.EndTime = iteration.EndTime;
        			test.Description = iteration.Browser.Title;
        			

        			test.Log(LogStatus.Info, "Test Case Name", chapter.Title);

        			foreach (Step step in chapter.Steps)
        			{
        				test.Log(LogStatus.Info,step.Title," ");

        				foreach (Act action in step.Actions)
        				{
        					if(action.IsSuccess)
        					{
        						test.Log(LogStatus.Pass,action.Title , "Action is passed");
        					}
        					else if(!action.IsSuccess)
        					{
        						test.Log(LogStatus.Fail,action.Title , "Action is failed ");
        						test.Log(LogStatus.Error,action.Extra,"Exception");
        						test.Log(LogStatus.Warning,"Failed Image is"+test.AddScreenCapture(Path.Combine(this.reportsPath, "Screenshots", String.Format("{0} {1} {2} Error.png", iteration.Browser.TestCase.Title, iteration.Browser.Title, iteration.Title))));
        					}


        				}

        			}
        		}

        	}
        	catch(Exception e)
        	{

        	}
        	try
        	{
        		ExtentTest.Flush();
        		EditTheHtml(this.Extentpath+"//Extent.html",FirstCaseBeginTime,LastCaseEndTime,ExecutionTimeCumulative);
        		
        		
        	}
        	catch(Exception e)
        	{
        		
        	}
        }
        
           		
         
        public void EditTheHtml(string path,DateTime start,DateTime End,TimeSpan cumm )
        { 
        	
        	string html = File.ReadAllText(path);
        	StringBuilder sb = new StringBuilder();
        	
        	 var doc = new HtmlDocument();
              doc.LoadHtml(html);

 
HtmlNodeCollection nodesMatchingXPath = doc.DocumentNode.SelectNodes("//div[contains(@class,'logo-contain')]/a/span");

        		
foreach (HtmlNode node in nodesMatchingXPath)
{
	node.Remove();
}
HtmlNodeCollection nodesMatchingXPath2 = doc.DocumentNode.SelectNodes("//span[contains(@class,'panel-lead suite-started-time')]");
foreach (HtmlNode node in nodesMatchingXPath2)
{
	node.InnerHtml = start.ToString();
}

HtmlNodeCollection nodesMatchingXPath3 = doc.DocumentNode.SelectNodes("//span[contains(@class,'panel-lead suite-ended-time')]");
foreach (HtmlNode node in nodesMatchingXPath3)
{
	node.InnerHtml =End.ToString();
	//node.ParentNode.ReplaceChild(HtmlTextNode.CreateNode(node.InnerText + End.ToString()), node);
}

HtmlNodeCollection nodesMatchingXPath4 = doc.DocumentNode.SelectNodes("//span[contains(@class,'report-name')]");
foreach (HtmlNode node in nodesMatchingXPath4)
{
	//node.ParentNode.ReplaceChild(HtmlTextNode.CreateNode(node.InnerText + "MetTel Automation Report"), node);
	node.InnerHtml = "Automation Report";
}


HtmlNodeCollection nodesMatchingXPath5 = doc.DocumentNode.SelectNodes("//span[contains(@class,'suite-total-time-taken panel-lead')]");
foreach (HtmlNode node in nodesMatchingXPath5)
{
	//node.ParentNode.ReplaceChild(HtmlTextNode.CreateNode(node.InnerText + "MetTel Automation Report"), node);
	node.InnerHtml = cumm.ToString();
}

File.WriteAllText(path,doc.DocumentNode.OuterHtml );

   }



public void PublishException()
        {
            #region HTML Template

            String template = @"
            <!DOCTYPE html>
            <html>
            <head>
	            <link href='http://netdna.bootstrapcdn.com/bootstrap/3.1.1/css/bootstrap.min.css' rel='stylesheet'>
	            <script src='http://code.jquery.com/jquery-1.11.0.min.js' type='text/javascript'></script>	
	            <script src='http://netdna.bootstrapcdn.com/bootstrap/3.1.1/js/bootstrap.min.js'></script>	
	        <style>
	            html {
	            overflow: -moz-scrollbars-vertical; 
	            overflow-y: scroll;
	            }	
	            .bigger-icon {
		            transform:scale(2.0,2.0);
		            -ms-transform:scale(2.0,2.0); /* IE 9 */
		            -moz-transform:scale(2.0,2.0); /* Firefox */
		            -webkit-transform:scale(2.0,2.0); /* Safari and Chrome */
		            -o-transform:scale(2.0,2.0); /* Opera */
	            }
	            .default {
		
		            font-family: Courier New;
					font-size: 15px;
	            }
	            .Report-Chapter {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #26466D;
		            color: #fff;
		            font-size: 90%; font-weight:bold;
		            border: 1px solid #03242C;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
		            cursor: pointer;
	            }	
	            .Report-Step {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #ddd;
		            color: #000;
		            font-size: 90%; font-weight:bold;
		            border: 1px solid #bebebe;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
		            cursor: pointer;
	            }	
	            .Report-Action {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #f7f7f9;
		            color: #000;
		            font-size: 90%;
		            border: 1px solid #e1e1e8;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
	            }	
	            .green { color:green; }
	            .red {color: red; }
	            .normal {color: black; }
	            .darkbg {background-image:url('#')};
	        </style>	


            <style>			   
			   #example thead th {
				  background-color: #1B3F73;
				  color: white;
                  text-align:center;
				}
				
			  #example tbody td:hover {					
				    cursor: pointer;
			    }
			
			  #example tbody td a{
			        text-decoration: none;
                    color: black;
			    }
	        </style>

            <script type='text/javascript'>
			    $(document).ready(function() {

				    $('#example tr').click(function() {
					    var href = $(this).find('a').attr('href');
					    if(href) {
						    window.open(href);
					    }
				    });
			    });
			</script>

            </head>

            <body>
	        <div class='container'>
		        <div style='padding-top: 5px; padding-bottom:5px;'>
			        <img src='https://upload.wikimedia.org/wikipedia/en/2/20/Big_Data_Logo.png' style='padding-top:20px; width:200px; height:auto;' />
			        <div class='pull-right'><img src='https://omni.okstate.edu/_resources/images/logo-banner.png'/></div>
			
		        </div>
	        </div>

            <div class='container default'>
                <div class='darkbg' style='background-color:#26466D; color:#fff; min-height:100px; padding:20px; margin-bottom:20px; margin-top:10px; top:-20px;'>
		            <div class='row'>		  		              
		              <div class='col-md-6' >
                            <br/>
                            <b> Server: </b> {{SERVER}}<br/>
                            <b> Total Exceptions: {{TOT_EXPS}}
                      </div>		             
		            </div>		
	            </div>
            </div>

            <div class='container'>
		         <div class='col-md-6' style='padding-left:0px;'> {{BARCHART_TABLE}} </div>		         		  		            
            </div>
            </br>
            <div class='container'>
	        <table id='example' class='table table-striped table-bordered table-condensed default table-hover'>
                 <thead>
		            <tr>
			            <th>SNO.</th>
			            <th>Test Case ID</th>
			            <th>Exception</th>
			            <th>Iteration</th>
			            <th>Duration</th>
                        <th>Issue</th>
			            <th>Result</th>
		            </tr>
                </thead>
			    <tbody>
                    {{CONTENT}}
                </tbody>
             </table>
            </div>           
            </body>
            </html>";
            #endregion

            /*           
            StringBuilder er = new StringBuilder();            
            if (MetTelExceptions.Count > 0)
            {                //for (int i = 0; i < MetTelExceptions.Count; i++)
                //{
                //   // er.AppendFormat("{0},{1} \n", MetTelExceptions[0].Message, MetTelExceptions[0].StackTrace);
                //}
                int TimedOutExceptions = MetTelExceptions.Where(x => x.Message.Contains("Timed out after")).Count();
                er.AppendFormat("{0} - {1} \n", "Total Exceptions", MetTelExceptions.Count);
                er.AppendFormat("{0} - {1} \n", "Time out Exceptions", TimedOutExceptions);
                er.AppendFormat("{0} - {1} \n", "Other Exceptions", MetTelExceptions.Count-TimedOutExceptions);
           
            }
            */
            StringBuilder builder = new StringBuilder();
            String strReturn = String.Empty;
            int total = 0;
            int timeOut = 0;
            int otherExceptions = 0;
            total = MetTelExceptions.Count;
            timeOut = MetTelExceptions.Where(x => x.ExceptionDetails.Message.Contains("Timed out after")).Count();
            //InvalidElementState = MetTelExceptions.Where(x => x.ExceptionDetails.Message.Contains("invalid element state")).Count();
            otherExceptions = MetTelExceptions.Count - timeOut;

            strReturn = strReturn + "<table class='table table-striped table-bordered table-condensed default'> <tr> <th colspan='4' style='background-color: #1B3F73; color: white'> <center> Exception Results </center> </th> </tr>";
            strReturn = strReturn + "<tr><th colspan='3' style='background-color: #1B3F73; color: Red'> Exception </th> <th style='background-color: #1B3F73; color: white; font-weight: bold'> <center> Total </center> </th> </tr>";
            strReturn = strReturn + "<tr><td colspan='3'>Time Out Exception </td><td ><center>" + timeOut + "</center></td> </tr>";
            strReturn = strReturn + "<tr><td colspan='3'>Others</td><td><center>" + otherExceptions + "</center></td> </tr>";
            strReturn = strReturn + "<tr><td colspan='3'>Total</td><td><center>" + total + "</center></td>  </tr>  </table>";

            template = template.Replace("{{SERVER}}", this.ServerName);
            template = template.Replace("{{TOT_EXPS}}", total.ToString());
            template = template.Replace("{{BARCHART_TABLE}}", strReturn);

            Int16 caseCounter = 1;

            foreach (ExceptionReport er in MetTelExceptions)
            {
                builder.Append("<tr>");
                builder.AppendFormat("<td>{0}</td>", caseCounter.ToString());
                builder.AppendFormat("<td><a href='{0}' target='_blank'>{1}</a></td>", er.PublishedReportPath, er.TestCaseId);
                builder.AppendFormat("<td>{0}</td>", er.ExceptionDetails.Message + " : Object is not identified");
                builder.AppendFormat("<td>{0}</td>", string.Empty);
                builder.AppendFormat("<td>{0}</td>", string.Empty);
                builder.AppendFormat("<td>{0}</td>", string.Empty);
                builder.AppendFormat("<td><span class='glyphicon glyphicon-{0}'></span></td>", false ? "ok green" : "remove red");
                builder.Append("</tr>");
                caseCounter++;
            }

            String fileName = Path.Combine(this.reportsPath, "ExceptionReport.html");

            using (StreamWriter output = new StreamWriter(fileName))
            {
                output.Write(template.Replace("{{CONTENT}}", builder.ToString()));
            }
        }







        /// <summary>
        /// Publishes Summary Report
        /// </summary>
        /// 
        
        public void Summarize(bool isFinal = true)
        {
            #region HTML Template

            String template = @"
            <!DOCTYPE html>
            <html>
            <head>
	            <link href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css' rel='stylesheet'>
	            <script src='http://code.jquery.com/jquery-3.1.0.min.js' type='text/javascript'></script>	
	            <script src='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js'></script>	
	        <style>
	            html {
	            overflow: -moz-scrollbars-vertical; 
	            overflow-y: scroll;
	            }	
	            .bigger-icon {
		            transform:scale(2.0,2.0);
		            -ms-transform:scale(2.0,2.0); /* IE 9 */
		            -moz-transform:scale(2.0,2.0); /* Firefox */
		            -webkit-transform:scale(2.0,2.0); /* Safari and Chrome */
		            -o-transform:scale(2.0,2.0); /* Opera */
	            }
	            .default {
		
		            font-family: Courier New;
					font-size: 15px;
	            }
	            .Report-Chapter {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #26466D;
		            color: #fff;
		            font-size: 90%; font-weight:bold;
		            border: 1px solid #03242C;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
		            cursor: pointer;
	            }	
	            .Report-Step {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #ddd;
		            color: #000;
		            font-size: 90%; font-weight:bold;
		            border: 1px solid #bebebe;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
		            cursor: pointer;
	            }	
	            .Report-Action {
		            padding:12px;
		            margin-bottom: 5px;		
		            background-color: #f7f7f9;
		            color: #000;
		            font-size: 90%;
		            border: 1px solid #e1e1e8;
		            border-radius: 4px;
		            font-family: Menlo,Monaco,Consolas,'Courier New',monospace;
	            }	
	            .green { color:green; }
	            .red {color: red; }
	            .normal {color: black; }
	            .darkbg {background-image:url('#')};
	        </style>	


            <style>			   
			   #example thead th {
				  background-color: #1B3F73;
				  color: white;
                  text-align:center;
				}
				
			  #example tbody td:hover {					
				    cursor: pointer;
			    }
			
			  #example tbody td a{
			        text-decoration: none;
                    color: black;
			    }
	        </style>

            <script type='text/javascript'>
			    $(document).ready(function() {

				    $('#example tr').click(function() {
					    var href = $(this).find('a').attr('href');
					    if(href) {
						    window.open(href);
					    }
				    });
			    });
			</script>

            </head>

            <body>
	        <div class='container'>
		        <div style='padding-top: 5px; padding-bottom:5px;'>
			        <img align='left' src='https://upload.wikimedia.org/wikipedia/en/2/20/Big_Data_Logo.png' style='padding-top:10px; width:200px; height:75px;' />
          <img align='right' src='https://omni.okstate.edu/_resources/images/logo-banner.png' style='width:200px;height:80px;'/>
		        </div>
	        </div>
                        <br></br>
                        <br></br>
                           <br></br>
            <div class='container default'>
                <div class='darkbg' style='background-color:#26466D; color:#fff; min-height:100px; padding:20px; margin-bottom:20px; margin-top:10px; top:-20px;'>
		            <div class='row'>		  		              
		              <div class='col-md-6' >
                            <br/>
                            <b> Server: </b> {{SERVER}}<br/>
                            <b> Parallel Cases (Max): {{MAX_PARALLEL}}
                      </div>
		              <div class='col-md-6' > <b> Start: </b> {{EXECUTION_BEGIN}}<br/> <b> End: </b> {{EXECUTION_END}}<br/> <b> Duration: </b> {{EXECUTION_DURATION}}<br/>  </div>		  
<!-- <b> Duration (Cumulative): {{EXECUTION_DURATION_CUM}}</b> -->
		            </div>		
	            </div>
            </div>

            <div class='container'>
		         <div class='col-md-6' style='padding-left:0px;'> {{BARCHART_TABLE}} </div>
		         <div class='col-md-6' > <div id='barChart' style='height:200px; width:550px;'></div> </div>		  		            
            </div>
            </br>
            <div class='container'>
	        <table id='example' class='table table-striped table-bordered table-condensed default table-hover' border='1' bordercolor='black' >
                 <thead>
		            <tr>
			            <th>SNO.</th>
			            <th>Test Case ID</th>
			            <th>Browser</th>
			            <th>Iteration</th>
			            <th>Duration</th>
                        <th>Issue</th>
			            <th>Result</th>
		            </tr>
                </thead>
			    <tbody>
                    {{CONTENT}}
                </tbody>
             </table>
            </div>
            <script type='text/javascript' src='https://www.google.com/jsapi'></script>
            <script  type='text/javascript'>

                google.load('visualization', '1', {packages:['corechart']});
                var BarChartData = {{BARCHARTDATA}};
                google.setOnLoadCallback(drawVisualization);
                function drawVisualization() {                    
                    var data = google.visualization.arrayToDataTable(BarChartData);

		            var options = {
                      title: 'Browser Wise Status',
                      legend: {position: 'top', alignment:'center'},
                      vAxis: {title: 'Count'},
		              hAxis: {title: 'Browser'},
                      seriesType: 'bars',
                      colors: ['green', 'red']
                    };
       
                    var chart = new google.visualization.ComboChart(document.getElementById('barChart'));
                    chart.draw(data, options);
                  }
                
             </script>
            </body>
            </html>";

            #endregion

            Int16 caseCounter = 1;
            StringBuilder builder = new StringBuilder();
            DateTime FirstCaseBeginTime = DateTime.Now;
            DateTime LastCaseEndTime = DateTime.Now;
            TimeSpan ExecutionTimeCumulative = TimeSpan.Zero;

            foreach (TestCase testCase in Reporter.TestCases)
            {
                foreach (Browser browser in testCase.Browsers)
                {
                    foreach (Iteration iteration in browser.Iterations.FindAll(itr => itr.IsCompleted == true))
                    {
                        builder.Append("<tr>");
                        builder.AppendFormat("<td>{0}</td>", caseCounter.ToString());
                        builder.AppendFormat("<td><a href='{0}' target='_blank'>{1}</a></td>", String.Format("{0} {1} {2}.html", testCase.Title, browser.Title, iteration.Title), testCase.Title);
                        builder.AppendFormat("<td>{0} {1}</td>", browser.BrowserName, browser.BrowserVersion);
                        builder.AppendFormat("<td align='center'>{0}</td>", iteration.Title);
                        builder.AppendFormat("<td align='center'>{0}</td>", iteration.EndTime.Subtract(iteration.StartTime).ToString());
                        builder.AppendFormat("<td>{0}</td>", iteration.BugInfo);
                        string result= iteration.IsSuccess == true ? "Pass" : "Fail";
                        builder.AppendFormat("<td align='center'><font color='{0}'>"+result.Trim()+"</font></td>", iteration.IsSuccess == true ? "green" : "red");
                       // builder.AppendFormat("<td><span class='glyphicon glyphicon-{0}'></span></td>", iteration.IsSuccess == true ? "ok green" : "remove red");
                        builder.Append("</tr>");
                        caseCounter++;

                        if (iteration.StartTime < FirstCaseBeginTime) FirstCaseBeginTime = iteration.StartTime;
                        if (iteration.EndTime > LastCaseEndTime) LastCaseEndTime = iteration.EndTime;
                        ExecutionTimeCumulative = ExecutionTimeCumulative.Add(iteration.EndTime.Subtract(iteration.StartTime));
                    }
                }
            }
            Dictionary<String, Dictionary<String, long>> getStatusByBrowser = summary.GetStatusByBrowser();

            template = template.Replace("{{SERVER}}", this.ServerName);
            template = template.Replace("{{MAX_PARALLEL}}", ConfigurationManager.AppSettings.Get("MaxDegreeOfParallelism"));
           // template = template.Replace("{{EXECUTION_BEGIN}}", String.Format("{0} {1}", FirstCaseBeginTime.ToString("MM-dd-yyyy HH:mm:ss"), "Central Time (US & Canada)"));
           // template = template.Replace("{{EXECUTION_END}}", String.Format("{0} {1}", LastCaseEndTime.ToString("MM-dd-yyyy HH:mm:ss"), "Central Time (US & Canada)"));
            template = template.Replace("{{EXECUTION_BEGIN}}",  FirstCaseBeginTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_END}}", LastCaseEndTime.ToString("MM-dd-yyyy HH:mm:ss"));
            template = template.Replace("{{EXECUTION_DURATION}}", LastCaseEndTime.Subtract(FirstCaseBeginTime).ToString());
            template = template.Replace("{{EXECUTION_DURATION_CUM}}", ExecutionTimeCumulative.ToString());
            template = template.Replace("{{BARCHARTDATA}}", BuildBarChartData(getStatusByBrowser));
            template = template.Replace("{{BARCHART_TABLE}}", BuildBarChartTable(getStatusByBrowser));

            String fileName = Path.Combine(this.reportsPath, isFinal ? "Summary.html" : "Summary_Provisional.html");
            lock (_provisionalSummaryLocker)
            {
                using (StreamWriter output = new StreamWriter(fileName))
                {
                    output.Write(template.Replace("{{CONTENT}}", builder.ToString()));
                }
            }
        }

        
       
        /// <summary>
        /// Build Bar Chart Data
        /// </summary>
        public string BuildBarChartData(Dictionary<String, Dictionary<String, long>> browserStatus)
        {
            String strReturn = String.Empty;
            int temp;

            strReturn = strReturn + "[ ['Browser', 'Passed',  { role: 'style' }, 'Failed',  { role: 'style' } ],";

            foreach (String browserName in browserStatus.Keys)
            {
                strReturn = strReturn + "['" + browserName + "',";

                temp = 1;
                Dictionary<String, long> status = browserStatus[browserName];
                foreach (long statusCount in status.Values)
                {
                    if (temp == 1)
                        strReturn = strReturn + statusCount + ", 'green',";
                    else
                        strReturn = strReturn + statusCount + ", 'red',";

                    temp++;
                }

                strReturn = strReturn.TrimEnd(',') + " ],";
            }

            strReturn = strReturn.TrimEnd(',');
            strReturn = strReturn + " ]";
            return strReturn;
        }

        /// <summary>
        /// Build Bar Chart Table
        /// </summary>
        public string BuildBarChartTable(Dictionary<String, Dictionary<String, long>> browserStatus)
        {
            String strReturn = String.Empty;
            int total = 0;
            int passedTotal = 0;
            int failedTotal = 0;
            int temp;

            strReturn = strReturn + "<table class='table table-striped table-bordered table-condensed' border='1' bordercolor='black'> <tr> <th colspan='4' style='background-color: #1B3F73; color: white'> <center> Test Result Status </center> </th> </tr>";
            strReturn = strReturn + "<tr> <th style='background-color: #1B3F73; color: white'> Test Results </th> <th style='background-color: #1B3F73; color: #1BDE38'> <center> Passed </center> </th> <th style='background-color: #1B3F73; color: red'> <center> Failed </center> </th> <th style='background-color: #1B3F73; color: white; font-weight: bold'> <center> Total </center> </th> </tr>";

            foreach (String browserName in browserStatus.Keys)
            {
                strReturn = strReturn + "<tr> <td> " + browserName + "</td>";

                Dictionary<String, long> status = browserStatus[browserName];

                total = 0;
                temp = 1;

                foreach (long statusCount in status.Values)
                {
                    strReturn = strReturn + "<td> <center> " + statusCount + " </center> </td>";
                    total = total + Convert.ToInt32(statusCount);

                    if (temp == 1)
                    {
                        passedTotal = passedTotal + Convert.ToInt32(statusCount);
                    }
                    else
                    {
                        failedTotal = failedTotal + Convert.ToInt32(statusCount);
                    }

                    temp++;
                }

                strReturn = strReturn + "<td style='font-weight: bold'> <center> " + total + " </center> </td> </tr>";
            }

            strReturn = strReturn + "<tr style='font-weight: bold'> <td> Total </td> <td> <center> " + passedTotal + " </center> </td> <td> <center> " + failedTotal + " </center> </td> <td> <center> " + (passedTotal + failedTotal) + " </center> </td> </tr> </table>";

            return strReturn;
        }

    }
}