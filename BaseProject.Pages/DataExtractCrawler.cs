using BaseProject;
using BaseProject.Report;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BaseProject.Pages
{
    public class DataExtractCrawler : BasePage
    {
   


        /// <summary>
        /// logic goes here where we extract data
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="reporter"></param>
        public static void LookForStories(RemoteWebDriver driver, Iteration reporter,string sitename)
        {
            reporter.Add(new Act("looking into the site for stories"));
            Console.WriteLine("Going to run for given site name");
            List<Dictionary<string, string>> data = ChooseDataRetrevalScheme(driver, reporter, sitename);
            WriteToFIles(reporter, sitename, data);

        }

        private static void WriteToFIles(Iteration reporter, string sitename, List<Dictionary<string, string>> data)
        {
            reporter.Add(new Act("Writing Files all <b></b> of them to D:MinedData"));
            int j = 1;
            string site = Regex.Replace(sitename, @"[^0-9a-zA-Z]+", "_");
            foreach (Dictionary<string, string> dict in data)
            {
                String currentsnap = DateTime.Now.ToString("yyyyMMdd_hhmm");
                var StoryHeaders = dict.Select(kvp => kvp.Key).ToList();
                string seperator = "~~~~~Seperate~~~~~";
                var Storys = dict.Select(kvp => kvp.Value).ToList();


                string totalTxt = StoryHeaders[0] + seperator + Storys[0];

                Console.WriteLine("Going to create a TXT");

                Common.CreateWithParallelFileWriteBytes(@"/home/nvadran/DataTemp/Text/", site + "_News_" + currentsnap + "_" + j, totalTxt);
                j++;

            }
            reporter.Add(new Act("After filtering suitable for text analytics is :: <b> " + j + " </b> of them in D:MinedData"));
        }

        private static List<Dictionary<string,string>> ChooseDataRetrevalScheme(RemoteWebDriver driver, Iteration reporter, string sitename)
        {
            List<string> urls = new List<string>();
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            try
            {

                if (sitename.ToLower().Contains("abcnews"))
                {
                    IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount").Location);

                    if (Stories.Count > 2)
                    {
                        //tring to get all urls in one go
                        foreach (IWebElement story in Stories)
                        {
                            urls.Add(story.GetAttribute("href"));
                        }
                        var DistinctList = urls.Distinct();
                        reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                        foreach (string url in DistinctList)
                        {
                            if (url.StartsWith("http://abcnews.go.com"))
                            {
                                Selenide.NavigateTo(driver, url);

                                IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader").Location);
                                IList<IWebElement> Story = driver.FindElementsByXPath(Util.GetLocator("Story").Location);

                                Dictionary<string, string> tempdict = new Dictionary<string, string>();
                                for (int i = 0; i < StoryHeader.Count; i++)
                                {

                                    tempdict.Add(StoryHeader[i].Text, Story[i].Text);
                                }

                                data.Add(tempdict);
                            }
                        }


                    }
                }
                else if (sitename.ToLower().Contains("cbsnews"))
                {
                    IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_CBS").Location);

                    if (Stories.Count > 2)
                    {
                        //tring to get all urls in one go
                        foreach (IWebElement story in Stories)
                        {
                            urls.Add(story.GetAttribute("href"));
                        }
                        var DistinctList = urls.Distinct();
                        reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                        foreach (string url in DistinctList)
                        {
                            if ((url != null && !url.StartsWith("https://www.cbsnews.com/video/")) && (url.StartsWith("https://www.cbsnews.com/news/")))
                            {
                                Selenide.NavigateTo(driver, url);
                                
                                Dictionary<string, string> tempdict = new Dictionary<string, string>();
                                try
                                {
                                    IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_CBS").Location);
                                    IList<IWebElement> Story = driver.FindElementsByXPath(Util.GetLocator("Story_CBS").Location);
                                    string metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_CBS").Location).Text;

                                    for (int i = 0; i < StoryHeader.Count; i++)
                                    {

                                        tempdict.Add(StoryHeader[i].Text + " " + metadata, Story[i].Text);
                                    }
                                }
                                catch (Exception e)
                                {
                                    reporter.Add(new Act("Url doesnot fit current scheme " + url));
                                    continue;
                                }

                                data.Add(tempdict);

                            }

                        }
                    }
                }
                else if (sitename.ToLower().Contains("cnn"))
                {
                    IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_CNN").Location);

                    if (Stories.Count > 2)
                    {
                        //tring to get all urls in one go
                        foreach (IWebElement story in Stories)
                        {
                            urls.Add(story.GetAttribute("href"));
                        }
                        var DistinctList = urls.Distinct();
                        string extra = "";
                        string metadata = "";
                        reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                        foreach (string url in DistinctList)
                        {
                            IList<IWebElement> StoryHeader = null;
                            IList<IWebElement> Story = null;
                            Selenide.NavigateTo(driver, url);
                            
                            string totalstory = "";
                            Dictionary<string, string> tempdict = new Dictionary<string, string>();
                            try
                            {
                                if (url.Contains("money.cnn"))
                                { 
                                
                                if (driver.FindElementsByXPath("//h1[@class='article-title speakable']").Count > 0)
                                    {
                                        StoryHeader = driver.FindElementsByXPath("//h1[@class='article-title speakable']");
                                        IWebElement SecndaryHeader = driver.FindElementByXPath("//h2[@class='speakable']");
                                        extra = SecndaryHeader.Text;
                                        metadata = driver.FindElementByXPath("//span[@class='cnnbyline ']").Text;
                                        Story = driver.FindElementsByXPath("//div[@id='storytext']//p");
                                    }
                                    else
                                    {
                                        reporter.Add(new Act("Identified schemes not avialable for this type of url " + url));
                                    }
                                }
                                else
                                {
                                    if (driver.FindElementsByXPath(Util.GetLocator("storyHeader_CNN").Location).Count > 0)
                                    {
                                        StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_CNN").Location);
                                        Story = driver.FindElementsByXPath(Util.GetLocator("Story_CNN").Location);

                                        if (driver.FindElementsByXPath(Util.GetLocator("storyHeaderMeta_CNN").Location).Count > 0)
                                        {
                                            metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_CNN").Location).Text;
                                        }

                                    }
                                    else
                                    {
                                        reporter.Add(new Act("Identified schemes not avialable for this type of url " + url));
                                    }
                                }

                                for (int i = 0; i < Story.Count; i++)
                                {
                                    totalstory += Story[i].Text;


                                }

                            }
                            catch (Exception e)
                            {
                                reporter.Add(new Act("Url doesnot fit current scheme " + url));
                                continue;
                            }
                            tempdict.Add(StoryHeader[0].Text + metadata + extra, totalstory);

                            data.Add(tempdict);

                        }


                    }
                }
                else if (sitename.ToLower().Contains("infowars"))
                {
                    IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_INF").Location);

                    if (Stories.Count > 2)
                    {
                        //tring to get all urls in one go
                        foreach (IWebElement story in Stories)
                        {
                            urls.Add(story.GetAttribute("href"));
                        }
                        var DistinctList = urls.Distinct();
                        reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                        foreach (string url in DistinctList)
                        {
                            Selenide.NavigateTo(driver, url);
                            
                            IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_INF").Location);
                            IList<IWebElement> Story = driver.FindElementsByXPath(Util.GetLocator("Story_INF").Location);

                            Dictionary<string, string> tempdict = new Dictionary<string, string>();
                            for (int i = 0; i < StoryHeader.Count; i++)
                            {

                                tempdict.Add(StoryHeader[i].Text, Story[i].Text);
                            }

                            data.Add(tempdict);

                        }


                    }
                }
                else if (sitename.ToLower().Contains("nytimes"))
                {
                    int k = 0;
                    try
                    {
                        IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_NYT").Location);

                        if (Stories.Count > 2)
                        {
                            //tring to get all urls in one go
                            foreach (IWebElement story in Stories)
                            {
                                urls.Add(story.GetAttribute("href"));
                            }
                            var DistinctList = urls.Distinct();
                            reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                            foreach (string url in DistinctList)
                            {
                                if (url.StartsWith("https://www.nytimes.com/"))
                                {
                                    Selenide.NavigateTo(driver, url);
                                   
                                    k++;
                                    Dictionary<string, string> tempdict = new Dictionary<string, string>();
                                    try
                                    {
                                        if (driver.FindElementsByXPath(Util.GetLocator("storyHeader_NYT").Location).Count > 0)
                                        {
                                            IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_NYT").Location);
                                            IList<IWebElement> Story = driver.FindElementsByXPath(Util.GetLocator("Story_NYT").Location);
                                            string metadata = "";
                                            if (driver.FindElementsByXPath(Util.GetLocator("storyHeaderMeta_NYT").Location).Count > 0)
                                            {
                                                metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_NYT").Location).Text;
                                            }
                                            string totalstory = "";

                                            string storyheader = StoryHeader[0].Text;
                                            for (int i = 0; i < Story.Count; i++)
                                            {
                                                new Actions(driver).MoveToElement(Story[i]).Perform();
                                                Selenide.Wait(driver, 1, true);
                                                totalstory += Story[i].Text;

                                            }
                                            tempdict.Add(storyheader + metadata, totalstory);
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    data.Add(tempdict);
                                }
                            }


                        }

                    }
                    catch (Exception e)
                    {

                    }

                }
                else if (sitename.ToLower().Contains("breitbart"))
                {
                    int k = 0;
                    try
                    {
                        IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_BB").Location);

                        if (Stories.Count > 2)
                        {
                            //tring to get all urls in one go
                            foreach (IWebElement story in Stories)
                            {
                                urls.Add(story.GetAttribute("href"));
                            }
                            var DistinctList = urls.Distinct();
                            reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                            foreach (string url in DistinctList)
                            {
                                if (url.StartsWith("http://www.breitbart.com/"))
                                {
                                    Selenide.NavigateTo(driver, url);
                                 
                                    k++;
                                    Dictionary<string, string> tempdict = new Dictionary<string, string>();
                                    try
                                    {
                                        if (driver.FindElementsByXPath(Util.GetLocator("storyHeader_BB").Location).Count > 0)
                                        {
                                            IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_BB").Location);
                                            IList<IWebElement> Story = driver.FindElementsByClassName(Util.GetLocator("Story_BB").Location);
                                            string metadata = "";
                                            if (driver.FindElementsByXPath(Util.GetLocator("storyHeaderMeta_BB").Location).Count > 0)
                                            {
                                                metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_BB").Location).Text;
                                            }
                                            string totalstory = "";

                                            string storyheader = StoryHeader[0].Text;
                                           
                                                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                                            try
                                            {
                                                totalstory = (string)js.ExecuteScript("return document.getElementsByClassName('" + Util.GetLocator("Story_BB").Location + "')[0].innerText;");
                                            }
                                            catch
                                            {
                                                for (int i = 0; i < Story.Count; i++)
                                                {
                                                    new Actions(driver).MoveToElement(Story[i]).Perform();
                                                    Selenide.Wait(driver, 1, true);
                                                    totalstory += Story[i].Text;

                                                }
                                            }
                                              
                                            
                                            
                                            tempdict.Add(storyheader + metadata, totalstory);
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        continue;
                                    }

                                    data.Add(tempdict);
                                }
                            }


                        }

                    }
                    catch (Exception e)
                    {

                    }
                }
                else if (sitename.ToLower().Contains("huffingtonpost"))
                {
                    int k = 1;
                    try
                    {
                        
                        try
                        {
                            IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_HFP").Location);

                            if (Stories.Count > 2)
                            {
                                //tring to get all urls in one go
                                foreach (IWebElement story in Stories)
                                {
                                    urls.Add(story.GetAttribute("href"));
                                }
                                var DistinctList = urls.Distinct();

                                reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                                foreach (string url in DistinctList)
                                {
                                    if (url.StartsWith("https://www.huffingtonpost.com/entry/"))
                                    {
                                        Selenide.NavigateTo(driver, url);

                                        Dictionary<string, string> tempdict = new Dictionary<string, string>();
                                        try
                                        {
                                            if (Selenide.Query.isElementVisibleboolValue(driver, Util.GetLocator("Story_HFP"), false))
                                            {
                                                IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_HFP").Location);
                                                IList<IWebElement> Story = driver.FindElementsById(Util.GetLocator("Story_HFP").Location);
                                                string metadata = "";
                                                try
                                                {
                                                    if (Selenide.Query.isElementVisibleboolValue(driver, Util.GetLocator("storyHeaderMeta_HFP"), false))
                                                    {
                                                        metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_HFP").Location).Text;
                                                    }
                                                }
                                                catch (Exception e)
                                                {

                                                }
                                                string totalstory = "";

                                                string storyheader = StoryHeader[0].Text;

                                                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                                                try
                                                {
                                                    totalstory = (string)js.ExecuteScript("return document.getElementsByClassName('" + Util.GetLocator("Story_HFP_cn").Location + "')[0].innerText;");
                                                    tempdict.Add(storyheader + metadata, totalstory);
                                                }
                                                catch
                                                {
                                                    for (int i = 0; i < Story.Count; i++)
                                                    {
                                                        new Actions(driver).MoveToElement(Story[i]).Perform();
                                                        Selenide.Wait(driver, 1, true);
                                                        totalstory += Story[i].Text;

                                                    }
                                                    tempdict.Add(storyheader + metadata, totalstory);
                                                }

                                                data.Add(tempdict);


                                            }
                                            else
                                            {
                                                reporter.Add(new Act("Current url is not working for the known scheme ::" + url));
                                            }

                                        }
                                        catch (Exception e)
                                        {
                                            continue;
                                        }


                                    }

                                    else
                                    {
                                        List<string> ParentUrls = new List<string>();
                                        Selenide.NavigateTo(driver, url);
                                        IList<IWebElement> Stories2 = driver.FindElementsByXPath(Util.GetLocator("StoryCount_HFP").Location);

                                        if (Stories.Count > 2)
                                        {
                                            //tring to get all urls in one go
                                            foreach (IWebElement story1 in Stories2)
                                            {
                                                ParentUrls.Add(story1.GetAttribute("href"));
                                            }
                                            var DistinctList2 = ParentUrls.Distinct();

                                            reporter.Add(new Act("Looking into <b>" + DistinctList.Count() + "</b> Stories"));

                                            foreach (string url2 in DistinctList2)
                                            {
                                                if (url2.StartsWith("https://www.huffingtonpost.com/entry/"))
                                                {
                                                    Selenide.NavigateTo(driver, url2);

                                                    Dictionary<string, string> tempdict2 = new Dictionary<string, string>();
                                                    try
                                                    {
                                                        if (Selenide.Query.isElementVisibleboolValue(driver, Util.GetLocator("Story_HFP"), false))
                                                        {
                                                            IList<IWebElement> StoryHeader2 = driver.FindElementsByXPath(Util.GetLocator("storyHeader_HFP").Location);
                                                            IList<IWebElement> Story2 = driver.FindElementsById(Util.GetLocator("Story_HFP").Location);
                                                            string metadata = "";
                                                            try
                                                            {
                                                                if (Selenide.Query.isElementVisibleboolValue(driver, Util.GetLocator("storyHeaderMeta_HFP"), false))
                                                                {
                                                                    metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_HFP").Location).Text;
                                                                }
                                                            }
                                                            catch (Exception e)
                                                            {

                                                            }
                                                            string totalstory = "";

                                                            string storyheader = StoryHeader2[0].Text;

                                                            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                                                            try
                                                            {
                                                                totalstory = (string)js.ExecuteScript("return document.getElementsByClassName('" + Util.GetLocator("Story_HFP_cn").Location + "')[0].innerText;");
                                                                tempdict2.Add(storyheader + metadata, totalstory);
                                                            }
                                                            catch
                                                            {
                                                                for (int i = 0; i < Story2.Count; i++)
                                                                {
                                                                    new Actions(driver).MoveToElement(Story2[i]).Perform();
                                                                    Selenide.Wait(driver, 1, true);
                                                                    totalstory += Story2[i].Text;

                                                                }
                                                                tempdict2.Add(storyheader + metadata, totalstory);
                                                            }

                                                            data.Add(tempdict2);


                                                        }
                                                        else
                                                        {
                                                            reporter.Add(new Act("Current url is not working for the known scheme ::" + url2));
                                                        }

                                                    }
                                                    catch (Exception e)
                                                    {
                                                        continue;
                                                    }


                                                }

                                            }
                                        }
                                    }

                                }


                            }
                            k++;
                            Selenide.NavigateTo(driver, sitename);
                            if (Selenide.Query.isElementExist(driver, Locator.Get(LocatorType.XPath, "//*[@id='pagination']/div[2]//a[contains(@href,'?page=2')]"), false))
                            {
                                try
                                {
                                    driver.FindElementByXPath("//*[@id='pagination']/div[2]//a[contains(@href,'?page=2')]").Click();
                                }
                                catch (Exception e)
                                {
                                    reporter.Add(new Act("Click never happened for page ::" + k));
                                }
                            }
                        }
                        catch (Exception e)
                        {

                        }

                    }
                    catch (Exception e)
                    {

                    }
                }
            }


            catch (Exception e)
            {

            }
            return data;
        }



        public static List<Dictionary<string, string>> Huffstories(RemoteWebDriver driver, Iteration reporter, string sitename)
        {
            List<string> urls = new List<string>();
            List<string> ParentUrls = new List<string>();
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            try
            {
                for (int k = 2; k < 12; k++)
                {
                    List<Dictionary<string, string>> temp = new List<Dictionary<string, string>>();
                    temp = DataExtractFromHuff(driver, reporter, sitename, urls, k);
                  
                    data.AddRange(temp);
                  
                    if(temp.Count==0)
                    {
                        break;
                    }
                }
            }

            catch (Exception e)
            {

            }
            WriteToFIles(reporter, sitename, data);
            return data;
        }

        private static List<Dictionary<string,string>> DataExtractFromHuff(RemoteWebDriver driver, Iteration reporter, string sitename, List<string> urls, int k)
        {
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            try
            {
                IList<IWebElement> Stories = driver.FindElementsByXPath(Util.GetLocator("StoryCount_HFP").Location);

                if (Stories.Count > 2)
                {
                    //tring to get all urls in one go
                    foreach (IWebElement story in Stories)
                    {
                        urls.Add(story.GetAttribute("href"));
                    }
                    var DistinctList = urls.Distinct();

                    reporter.Add(new Act("Got <b>" + DistinctList.Count() + "</b> Stories"));
                    foreach (string url in DistinctList)
                    {
                        if (url.StartsWith("https://www.huffingtonpost.com/entry/"))
                        {
                            Selenide.NavigateTo(driver, url);

                            Dictionary<string, string> tempdict = new Dictionary<string, string>();
                            try
                            {
                                if (driver.FindElementsById(Util.GetLocator("Story_HFP").Location).Count > 0)
                                {
                                    IList<IWebElement> StoryHeader = driver.FindElementsByXPath(Util.GetLocator("storyHeader_HFP").Location);
                                    IList<IWebElement> Story = driver.FindElementsById(Util.GetLocator("Story_HFP").Location);
                                    string metadata = "";
                                    try
                                    {
                                        if (Selenide.Query.isElementVisibleboolValue(driver, Util.GetLocator("storyHeaderMeta_HFP"), false))
                                        {
                                            metadata = driver.FindElementByXPath(Util.GetLocator("storyHeaderMeta_HFP").Location).Text;
                                        }
                                    }
                                    catch (Exception e)
                                    {

                                    }
                                    string totalstory = "";

                                    string storyheader = StoryHeader[0].Text;

                                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                                    try
                                    {
                                        totalstory = (string)js.ExecuteScript("return document.getElementsByClassName('" + Util.GetLocator("Story_HFP_cn").Location + "')[0].innerText;");
                                        tempdict.Add(storyheader + metadata, totalstory);
                                    }
                                    catch
                                    {
                                        for (int i = 0; i < Story.Count; i++)
                                        {
                                            new Actions(driver).MoveToElement(Story[i]).Perform();
                                            Selenide.Wait(driver, 1, true);
                                            totalstory += Story[i].Text;

                                        }
                                        tempdict.Add(storyheader + metadata, totalstory);
                                    }

                                    data.Add(tempdict);


                                }
                                else
                                {
                                    reporter.Add(new Act("Current url is not working for the known scheme ::" + url));
                                }

                            }
                            catch (Exception e)
                            {
                                continue;
                            }


                        }



                    }


                }
                Selenide.NavigateTo(driver, sitename);
                if (Selenide.Query.isElementExist(driver, Locator.Get(LocatorType.XPath, "//*[@id='pagination']/div[2]//a[contains(@href,'?page=" + k + "')]"), false))
                {
                    try
                    {
                        Selenide.Click(driver, Locator.Get(LocatorType.XPath, "//*[@id='pagination']/div[2]//a[contains(@href,'?page=" + k + "')]"));
                        //driver.FindElementByXPath("//*[@id='pagination']/div[2]//a[contains(@href,'?page=2')]").Click();
                    }
                    catch (Exception e)
                    {
                        reporter.Add(new Act("Click never happened for page ::" + k));
                    }
                }
            }
            catch (Exception e)
            {
                
            }

            return data;
        }
    }
}
