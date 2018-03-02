using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using System.Net;
using System.Net.Http;

namespace BaseProject
{
    /// <summary>
    /// Mercury is wrapper around Selenium that abstracts many functions
    /// 
    /// Selenium is produced from Selenides (Salts of Selenium)
    /// Mercury Selenide [HgSe] is one of Selenide
    /// </summary>
    public static class Selenide
    {

        public static class Browser
        {
            /// <summary>
            /// Gets value indicating whether the current associated browser is Internet Explorer or not
            /// </summary>
            public static bool isInternetExplorer(RemoteWebDriver driver)
            {

                return (driver.Capabilities.BrowserName.ToUpper().Equals("IE") ||
                        driver.Capabilities.BrowserName.ToUpper().Equals("INTERNET EXPLORER"))
                    ? true
                    : false;

            }

            /// <summary>
            /// Gets value indicating whether the current associated browser is iPad Safari or not
            /// </summary>
            public static bool isIPadSafari(RemoteWebDriver driver)
            {

                return (driver.Capabilities.BrowserName.ToUpper().Equals("IPAD") ||
                        driver.Capabilities.BrowserName.ToUpper().Equals("IOS_SAFARI_IPADAIR_REAL"))
                    ? true
                    : false;

            }

        }

        /// <summary>
        /// Enumeration of Control Type
        /// </summary>
        public enum ControlType
        {
            Textbox,
            Select,
            Select2,
            Label,
            Checkbox,
            Button,
            IFrame,
            Listbox,
            Index
        }

        private static Nullable<Boolean> shouldHighlight;

        /// <summary>
        /// Gets ShouldHighlight
        /// </summary>
        private static Boolean ShouldHighlight
        {
            get
            {
                if (!shouldHighlight.HasValue)
                {
                    shouldHighlight = ConfigurationManager.AppSettings.Get("ElementHighlight").ToString() == "Yes" ? true : false;
                }
                return shouldHighlight.Value;
            }
        }

       


        /// <summary>
        /// Navigates to specified URL
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="location"></param>
        public static void NavigateTo(RemoteWebDriver driver, String location)
        {
            driver.Navigate().GoToUrl(location);
            //Wait(driver,1, false);
            try
            {
                //Selenide.WaitForPageLoad(driver);
                Thread.Sleep(1000);
            }
            catch (Exception e)
            {

            }
        }

        /// <summary>
        /// Refreshs The Browser
        /// </summary>
        /// <param name="Driver"></param>
        /// <param name="location"></param>
        public static void BrowserRefresh(RemoteWebDriver driver)
        {
            driver.Navigate().Refresh();
        }

        /// <summary>
        /// Back Browser
        /// </summary>
        /// <param name="Driver"></param>
        /// <param name="location"></param>
        public static void BrowserBack(RemoteWebDriver driver)
        {
            driver.Navigate().Back();
            System.Threading.Thread.Sleep(3000);
        }

        //#############################################################################
        // Function Name : SwithToLastWindow
        // Description : View the left column,View the center column,View the right column
        // Author : Priya Batchu
        // #############################################################################
        public static void SwithToLastWindow(RemoteWebDriver driver)
        {
            IList<string> windows = driver.WindowHandles;

            for (int i = 0; i < windows.Count; i++)
            {
                if (i == (windows.Count - 1))
                {
                    driver.SwitchTo().Window(windows[i]);
                }
            }
        }


        /// <summary>
        /// Removing Escape Sequences
        /// </summary>
        /// <param name="dataValue"></param>
        /// <returns></returns>
        public static string RemoveEscapeChars(string dataValue)
        {
            if (!String.IsNullOrEmpty(dataValue))
            {
                dataValue = dataValue.Replace("\r\n", "");
            }
            return dataValue;
        }
        /// <summary>
        /// Initializes and reuses WebDriverWait instance
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="seconds">Seconds to wait, or 0 to use default from config</param>
        /// <returns>WebDriverWait instance</returns>
        private static WebDriverWait GetWaiter(RemoteWebDriver driver, int seconds = 0)
        {
            return new WebDriverWait(driver, TimeSpan.FromSeconds(seconds > 0 ? seconds : Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementSyncTimeOut"))));
        }

        /// <summary>
        /// Executes JavaScript on an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="element"><see cref="IWebElement"/></param>
        /// <param name="js">Javascript code to execute</param>
        private static void ExecuteJS(RemoteWebDriver driver, IWebElement element, String js)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript(js, new object[] { element });
        }

        public static void ExecuteJS(RemoteWebDriver driver, Locator locator, String js)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript(js, new object[] { GetElement(driver, locator) });
        }

        /// <summary>
        /// Highlights an element by drawing border
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="element"><see cref="IWebElement"/></param>
        /// <param name="isOrange">Orange default, Red otherwise</param>
        /// <returns><see cref="IWebElement"/></returns>
        private static IWebElement Highlight(RemoteWebDriver driver, IWebElement element, Boolean isOrange = true)
        {
            if (ShouldHighlight)
            {
                //String script = String.Format(@"arguments[0].style.cssText = ""border-width: 4px; border-style: solid; border-color: {0}"";", isOrange ? "orange" : "red");
                //IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
               // jsExecutor.ExecuteScript(script, new object[] { element });
            }
            return element;
        }

        /// <summary>
        /// Waits for specified amount time
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="seconds">Number of seconds to wait</param>
        /// <param name="forceThread">Specify whether to wait at Driver level or Thread</param>
        public static void Wait(RemoteWebDriver driver, int seconds, bool threadLevel = false)
        {
            if (threadLevel)
            {
                Thread.Sleep(seconds * 1000);
            }
            else
            {
                driver.Manage().Timeouts().ImplicitlyWait(new TimeSpan(0, 0, seconds));
            }
        }



        /// <summary>
        /// Waits for an element to be visible on page
        /// </summary>
  
        /// <param name="seconds">Number of Seconds to wait. 0 (Default) means, globale wait time (from Settings)</param>
        public static void WaitForElementVisible(RemoteWebDriver driver, Locator locator, int seconds = 0)
        {
            Highlight(driver, GetWaiter(driver, seconds).Until(ExpectedConditions.ElementIsVisible(locator.GetBy())));
        }

        public static void WaitForPageLoad(RemoteWebDriver driver,int seconds = 0)
        {
            string state = string.Empty;
            try
            {
                try
                {

                    //int time = driver, GetWaiter(driver, seconds);
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds > 0 ? seconds : Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementSyncTimeOut"))));

                    //Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns ture
                    wait.Until(d =>
                    {

                        try
                        {
                            state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                        }
                        catch (InvalidOperationException)
                        {
                        //Ignore
                    }
                        catch (NoSuchWindowException)
                        {
                        //when popup is closed, switch to last windows
                        //driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    //In IE7 there are chances we may get state as loaded instead of complete
                    return (state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase));

                    });
                }
                catch (TimeoutException)
                {
                    //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                    if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                        throw;
                }
                catch (NullReferenceException)
                {
                    //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                    if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                        throw;
                }
                catch (WebDriverException)
                {
                    if (driver.WindowHandles.Count == 1)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[0]);
                    }
                    state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                    if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
                        throw;
                }
            }catch(Exception e)
            {

            }
        }


        /// <summary>
        /// Waits for an element to be exist on page
        /// </summary>
   
        /// <param name="seconds">Number of Seconds to wait. 0 (Default) means, globale wait time (from Settings)</param>
        public static void WaitForElementExist(RemoteWebDriver driver, Locator locator, int seconds = 0)
        {
            Highlight(driver, GetWaiter(driver).Until(ExpectedConditions.ElementExists(locator.GetBy())));
        }

        /// <summary>
        /// Verifies specified element is enabled or disabled
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="shouldEnabled">True if expectation is 'Enabled'</param>
        public static void VerifyEnabledOrDisabled(RemoteWebDriver driver, Locator locator, Boolean shouldEnabled = true, ControlType controlType = ControlType.Textbox)
        {
            IWebElement element = null;

            if (controlType == ControlType.Select2)
            {
                element = GetElement(driver, new Locator(LocatorType.XPath, locator.Location + "/../input[@type='text' and contains(@class,'select2-focusser')]"));
            }
            else
            {
                element = GetElement(driver, locator);
            }

            if (!Highlight(driver, element).Enabled == shouldEnabled)
            {
                throw new Exception("Expected State of Element is not Matching");
            }
        }

        /// <summary>
        /// Verifies specified element is visible or invisible
        /// </summary>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="shouldEnabled">True if expectation is 'Visible'</param>
        public static void VerifyVisibleOrInvisible(RemoteWebDriver driver, Locator locator, Boolean shouldVisible = true)
        {
            try
            {
                Highlight(driver, GetWaiter(driver, Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementSyncTimeOut"))).Until(ExpectedConditions.ElementIsVisible(locator.GetBy())));

                bool isDisplayed = driver.FindElement(locator.GetBy()).Displayed;

                if (shouldVisible != isDisplayed)
                {
                    shouldVisible = true;
                    throw new Exception("Element selection state is not matching with expected");
                }
            }
            catch (Exception ex)
            {
                if (shouldVisible)
                {
                    throw new Exception("Element selection state is not matching with expected");
                }
            }
        }

        /// <summary>
        /// Verifies specified element is selected or not
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="selected">True if expectation is 'Checked'</param>        
        public static void VerifyCheckedOrUnchecked(RemoteWebDriver driver, Locator locator, Boolean selected = true)
        {
            if (GetElement(driver, locator).Selected != selected)
            {
                throw new Exception("Element selection state is not matching with expected");
            }
        }

        /// <summary>
        /// Verifies specified element is selected or not
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="selected">True if expectation is 'Checked'</param>        
        public static Boolean GetCheckedStatus(RemoteWebDriver driver, Locator locator)
        {
            return GetElement(driver, locator).Selected;
        }

     
        /// <summary>
        /// Gets element count 
        /// </summary>
        /// <param name="locator"></param>
        /// <returns></returns>
        public static Int32 GetElementCount(RemoteWebDriver driver, Locator locator)
        {
            IList<IWebElement> elements = driver.FindElements(locator.GetBy());
            return elements.Count;
        }

        /// <summary>
        /// Gets element collection 
        /// </summary>        
        /// <returns></returns>
        public static IList<String> GetElementNames(RemoteWebDriver driver, Locator locator)
        {
            List<String> result = new List<string>();
            foreach (IWebElement element in driver.FindElements(locator.GetBy()))
            {
                result.Add(element.Text);
            }
            return result;
        }

        /// <summary>
        /// Gets an element from DOM
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <returns><see cref="IWebElement"/></returns>
        public static IWebElement GetElement(RemoteWebDriver driver, Locator locator)
        {
            return Highlight(driver, driver.FindElement(locator.GetBy()));
        }

        

        public class Query
        {
            /// <summary>
            /// Verifies visibility of an element
            /// </summary>
            /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            /// <param name="by"><see cref="by"/></param>
            /// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            /// <returns>Boolean value representing existence of element</returns>
            public static bool isElementVisible(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
               // int Seconds = Convert.ToInt32(ConfigurationManager.AppSettings.Get("ElementSyncTimeOut"));
                //Highlight(driver, GetWaiter(driver, Seconds).Until(ExpectedConditions.ElementIsVisible(locator.GetBy())));                
                IWebElement element = null;
                try
                {
                    if (wait)
                    {
                        WaitForElementVisible(driver, locator);
                        element = GetElement(driver, locator);
                    }
                    else
                    {
                        element = GetElement(driver, locator);
                    }
                    
                    return element.Displayed;
                }
                catch (Exception ex)
                {
                    throw new Exception("Object is not identified  Exception is : " + ex);
                    
                }
            }


            ///// <summary>
            ///// Verifies visibility of an element
            ///// </summary>
            ///// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            ///// <param name="by"><see cref="by"/></param>
            ///// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            ///// <returns>Boolean value representing existence of element</returns>
            //public static bool isElementVisible(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            //{
            //   
            //    IWebElement element = null;
            //    try
            //    {
            //        if (wait)
            //        {
            //            WaitForElementVisible(driver, locator);
            //            element = GetElement(driver, locator);
            //        }
            //        else
            //        {
            //            element = GetElement(driver, locator);
            //        }
            //        return element.Displayed;
            //    }
            //    catch (Exception)
            //    {
            //        return false;
            //    }
            //}

            //<summary>
            ///// Verifies visibility of an element
            ///// </summary>
            ///// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            ///// <param name="by"><see cref="by"/></param>
            ///// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            ///// <returns>Boolean value representing existence of element</returns>
            public static bool isElementVisibleboolValue(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
                IWebElement element = null;
                try
                {
                    if (wait)
                    {
                        WaitForElementVisible(driver, locator);
                        element = GetElement(driver, locator);
                    }
                    else
                    {
                        element = GetElement(driver, locator);
                    }
                    return element.Displayed || element.Enabled;
                }
                catch (Exception)
                {
                    return false;
                }
            }



            ///<summary>
            /// verify targeted image is broken or alive
            ///</summary>
            public static bool IsImageDisplayed(RemoteWebDriver driver, Locator locator)
            {
                IWebElement element = Highlight(driver, GetElement(driver, locator), false);
                string uriToImage = element.GetAttribute("src");
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uriToImage);

                request.Method = "HEAD";

                try
                {
                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {

                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                catch (WebException e) { throw new Exception("Image is not displayed Exception is : " + e); }
                catch (Exception ex)
                {
                    throw new Exception("Object is not identified  Exception is : " + ex);

                }

            }

           
            /// <summary>
            /// Verifies visibility of an element
            /// </summary>
            /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            /// <param name="by"><see cref="by"/></param>
            /// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            /// <returns>Boolean value representing existence of element</returns>
            public static bool isElementInvisible(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
                IWebElement element = null;
                bool isException = true;

                if (!wait)
                {
                    try
                    {
                        element = GetElement(driver, locator);
                        if (element.Displayed != wait)
                        {
                            isException = false;
                            throw new Exception("Element selection state is not matching with expected");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (isException == false)
                        {
                            throw new Exception("Element selection state is not matching with expected");
                        }

                        isException = true;
                    }
                }
                else
                {
                    try                    
                    {
                        element = GetElement(driver, locator);
                        if (element.Displayed)
                        {
                            isException = false;
                        }
                        else if (element.Displayed == wait)
                        {
                            isException = false;
                            throw new Exception("Element selection state is not matching with expected");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (wait == true)
                        {
                            throw new Exception("Element selection state is not matching with expected");
                        }

                        isException = true;
                    }
                }
                return isException;
            }

            /// <summary>
            /// Verifies visibility of an element for Check Annotation [PDM]
            /// </summary>
            /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            /// <param name="by"><see cref="by"/></param>
            /// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            /// <returns>Boolean value representing existence of element</returns>
            public static bool isElementVisibleForAnnotation(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
                IWebElement element = null;
                bool isException = false;

                if (wait)
                {
                    WaitForElementVisible(driver, locator);
                    element = GetElement(driver, locator);
                }
                else
                {
                    try
                    {
                        element = GetElement(driver, locator);
                        if (element.Displayed != wait)
                        {
                            isException = true;
                            throw new Exception("Element selection state is not matching with expected");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (isException == true)
                        {
                            throw new Exception("Element selection state is not matching with expected");
                        }

                        return true;
                    }
                }
                return element.Displayed;
            }

            /// <summary>
            /// Verifies existence of an element
            /// </summary>
            /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            /// <param name="by"><see cref="by"/></param>
            /// <param name="wait">Flag to indicate whether to wait while looking for element</param>            
            /// <returns>Boolean value representing existence of element</returns>
            public static bool isElementExist(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
                try
                {
                    if (wait)
                    {
                        WaitForElementExist(driver, locator);
                    }
                    else
                    {
                        Highlight(driver, GetElement(driver, locator));
                    }
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            /// <summary>
            /// Verifies existence of an element
            /// </summary>
            /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
            /// <param name="by"><see cref="by"/></param>
            /// <param name="wait">Flag to indicate whether to wait while looking for element</param>
            /// <returns>Boolean value representing whether it is enabled or not</returns>
            public static bool isElementEnabled(RemoteWebDriver driver, Locator locator, Boolean wait = true)
            {
                IWebElement element = null;

                try                
                {
                    if (wait)
                    {
                        WaitForElementVisible(driver, locator);
                        element = Highlight(driver,GetElement(driver,locator));
                    }
                    else
                    {
                        element = Highlight(driver, GetElement(driver, locator));
                    }
                    return element.Enabled;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        public class JS
        {
            /// <summary>
            /// Performs Click with JavaScript
            /// </summary>
            /// <param name="locator"></param>
            public static void Click(RemoteWebDriver driver, Locator locator)
            {
                IWebElement element = Highlight(driver, GetElement(driver, locator), false);
                //JQuery trigger onClick event only. Never Clicks. So, do native JS Click
                ExecuteJS(driver, element, @"$(arguments[0])[0].click();");
            }

            /// <summary>
            /// Sets Text/Value/Selection with JavaScript
            /// </summary>
            public static void SetText(RemoteWebDriver driver, Locator locator, ControlType controlType, String value)
            {
                IWebElement element = Highlight(driver, GetElement(driver, locator), false);
                switch (controlType)
                {
                    case ControlType.Textbox:
                        ExecuteJS(driver, element, String.Format(@"$(arguments[0]).val('{0}');", value));
                        break;
                }

                ExecuteJS(driver, element, @"$(arguments[0]).change();");
            }

            /// <summary>
            /// Executes any JS commands
            /// refer: http://selenium.googlecode.com/svn/trunk/docs/api/dotnet/html/AllMembers_T_OpenQA_Selenium_IJavaScriptExecutor.htm
            /// </summary>
            /// <param name="js">JS to execute</param>
            /// <returns>An object representing the return value</returns>
            public static object GetObject(RemoteWebDriver driver, String js)
            {
                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                return jsExecutor.ExecuteScript(js, new object[] { Locator.Get(LocatorType.XPath, "\\body").Location });
            }

            public static void Enter(RemoteWebDriver driver, Locator locator)
            {
                IWebElement element = Highlight(driver, GetElement(driver, locator), false);
                ExecuteJS(driver, element, @"$(arguments[0]).trigger(jQuery.Event('keypress', {which: 13}));");
            }

            public static void Focus(RemoteWebDriver driver, Locator locator)
            {
                IWebElement element = GetElement(driver, locator);
                ExecuteJS(driver, element, String.Format(@"$(arguments[0].scrollIntoView(true));"));
                Wait(driver, 2, true);
            }
        }

        /// <summary>
        /// Gets text from an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="controlType"><see cref="ControlType"/></param>
        /// <returns>Text of element</returns>
        public static string GetText(RemoteWebDriver driver, Locator locator, ControlType controlType)
        {
            IWebElement element = GetElement(driver, locator);
            String value = String.Empty;

            switch (controlType)
            {
                case ControlType.Select:
                    value = new SelectElement(element).SelectedOption.Text;
                    break;

                case ControlType.Select2:
                    value = element.Text;
                    break;

                case ControlType.Textbox:
                    value = element.GetAttribute("value");
                    break;

                case ControlType.Label:
                    value = element.Text;
                    break;

                case ControlType.Button:
                    value = element.GetAttribute("value");
                    break;
            }

            return value;
        }

        /// <summary>
        /// Sets text to an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="controlType"><see cref="ControlType"/></param>
        /// <param name="text">Text to set</param>
        public static void SetText(RemoteWebDriver driver, Locator locator, ControlType controlType, String text, bool split = false, bool JSForIE = true)
        {
            IWebElement element = Highlight(driver, GetElement(driver, locator), false);

            switch (controlType)
            {
                case ControlType.Select:
                    new SelectElement(element).SelectByText(text);
                    break;

                case ControlType.Index:
                    new SelectElement(element).SelectByIndex(Convert.ToInt16(text));
                    break;

                case ControlType.Listbox:                    
                    SelectElement se = new SelectElement(element);
                    se.SelectByText(text);                    
                    break;

                case ControlType.Textbox:
                    element.Clear();
                    if (text.Contains("'") || text.Contains(@""""))
                    {
                        text = text.Replace("'", "\\'");
                        ExecuteJS(driver, element, String.Format("$(arguments[0]).val('{0}');", text));
                    }
                    else
                    {
                        if (split == true && JSForIE == false)
                        {
                            try
                            {

                                foreach (Char each in text.ToCharArray())
                                    element.SendKeys(each.ToString());
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        else if (driver.Capabilities.BrowserName.ToUpper().Equals("IE") || driver.Capabilities.BrowserName.ToUpper().Equals("INTERNET EXPLORER") &&
                            split == false && JSForIE == true)
                        {
                            JS.SetText(driver, locator, controlType, text);
                        }
                        else
                        {
                            element.SendKeys(text);
                        }
                    }

                    // HACK: SendKeys not raising 'Change' is event in DOM
                    // Having said that, Knockout bound control's model don't get update about changed value
                    ExecuteJS(driver, element, @"$(arguments[0]).change();");
                    break;

                case ControlType.Select2:
                    element.Click();
                    if ((driver.Capabilities.BrowserName.ToUpper().Equals("IE") || driver.Capabilities.BrowserName.ToUpper().Equals("INTERNET EXPLORER")))
                    {
                        SetText(driver, Locator.Get(LocatorType.XPath, "//div[@id='select2-drop']/div[@class='select2-search']/input"), ControlType.Textbox, text, true, false);
                    }
                    else
                    {
                        SetText(driver, Locator.Get(LocatorType.XPath, "//div[@id='select2-drop']/div[@class='select2-search']/input"), ControlType.Textbox, text);
                    }
                    GetElement(driver, Locator.Get(LocatorType.XPath, "//span[@class='select2-match']")).Click();
                    break;

                case ControlType.IFrame:
                    element.Clear();
                    if (text.Contains("'") || text.Contains(@""""))
                    {
                        text = text.Replace("'", "\\'");
                        element.SendKeys(text);
                    }
                    else
                    {
                        element.SendKeys(text);
                    }
                    break;
            }
        }

        /// <summary>
        /// Gets list of options avaialble
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="By"/></param>
        /// <param name="controlType"><see cref="ControlType"/></param>
        /// <returns>Avaialble options as <see cref="List<String>"/></returns>
        public static List<String> GetOptions(RemoteWebDriver driver, Locator locator, ControlType controlType)
        {
            IWebElement element = Highlight(driver, GetElement(driver, locator), false);
            List<String> options = new List<string>();

            switch (controlType)
            {
                case ControlType.Select:
                    foreach (IWebElement each in new SelectElement(element).Options)
                    {
                        options.Add(each.Text);
                    }
                    break;
            }

            return options;
        }

        /// <summary>
        /// Sets specified state to switch (Radio, Checkbox)
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="check">Should check or not</param>
        public static void SetSwitch(RemoteWebDriver driver, Locator locator, bool check = true)
        {
            IWebElement element = GetElement(driver, locator);
            if (element.Selected != check)
            {
                element.Click();
            }
        }

        /// <summary>
        /// Gets state of switch
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <returns>Boolean value representing current state</returns>
        public static bool GetSwitch(RemoteWebDriver driver, Locator locator)
        {
            return GetElement(driver, locator).Selected;
        }

        /// <summary>
        /// Clears value of an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        /// <param name="controlType"><see cref="ControlType"/></param>
        public static void Clear(RemoteWebDriver driver, Locator locator, ControlType controlType)
        {
            IWebElement element = Highlight(driver, GetElement(driver, locator), false);

            switch (controlType)
            {
                case ControlType.Select:
                    new SelectElement(element).DeselectAll();
                    break;

                case ControlType.Textbox:
                    element.Clear();
                    break;
            }
        }

        /// <summary>
        /// Moves Focus to an element
        /// </summary>
        /// <param name="locator"></param>
        public static void Focus(RemoteWebDriver driver, Locator locator)
        {
            // HACK: Control's that out of Window visible area are not receving click (Not always)
            // So, the idea is to set focus to it by calling MoveToElement (Mimics mouse move)
            // However, this kind of stuff not working in Touch devices
            if (driver.Capabilities.BrowserName.ToUpper().Equals("CHROME") ||
                driver.Capabilities.BrowserName.ToUpper().Equals("IE") ||
                driver.Capabilities.BrowserName.ToUpper().Equals("INTERNET EXPLORER"))
            {
                IWebElement element = GetElement(driver, locator);
                new Actions(driver).MoveToElement(element).Perform();
                Wait(driver, 2, true);
            }
        }

        /// <summary>
        /// Clicks an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        public static void Click(RemoteWebDriver driver, Locator locator)
        {
            try
            {
                IWebElement element = GetElement(driver, locator);

                // HACK: Control's that out of Window visible area are not receving click (Not always)
                // So, the idea is to set focus to it by calling MoveToElement (Mimics mouse move)
                // However, this kind of stuff not working in Touch devices
                if (driver.Capabilities.BrowserName.ToUpper().Equals("CHROME") ||
                    driver.Capabilities.BrowserName.ToUpper().Equals("IE") ||
                    driver.Capabilities.BrowserName.ToUpper().Equals("INTERNET EXPLORER"))
                {
                    new Actions(driver).MoveToElement(element).Perform();
                    Wait(driver, 2, true);
                }
                Highlight(driver, element, false).Click();
            }
            catch (Exception e)
            {
                //throw new Exception("Exception is" + e);
            }
        }

        /// <summary>
        /// Switch To Frame Element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        public static void SwitchToFrame(RemoteWebDriver driver, Locator locator)
        {
            driver.SwitchTo().Frame(GetElement(driver, locator));
        }

        /// <summary>
        /// Switch To Default Content
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        public static void SwitchToDefaultConent(RemoteWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
        }

        /// <summary>
        /// Sends Enter Key
        /// </summary>
        /// <param name="locator"></param>
        public static void Enter(RemoteWebDriver driver, Locator locator)
        {
            IWebElement element = Highlight(driver, GetElement(driver, locator), false);
            element.SendKeys(Keys.Enter);
        }

        /// <summary>
        /// Clicks an element
        /// </summary>
        /// <param name="Driver"><see cref="RemoteWebDriver"/></param>
        /// <param name="by"><see cref="by"/></param>
        public static void ClickAddFile(RemoteWebDriver driver, Locator locator, string path)
        {
            //Driver.FindElement(By.XPath("//input[@id='fileupload']")).SendKeys(path);            
            driver.FindElement(locator.GetBy()).SendKeys(path);
        }

        /// <summary>
        /// Get Time Stamp
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="timeFormat"></param>
        public static string GetTimeStamp(string dateTime, string timeFormat)
        {
            DateTime dt = DateTime.Parse(dateTime);
            return dt.ToString(timeFormat);
        }

        /// <summary>
        /// Get Date Time Stamp
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="timeFormat"></param>
        public static string GetDateTimeStamp()
        {
            string dateTimeStamp = String.Empty;
            dateTimeStamp = DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + DateTime.Now.Millisecond.ToString();
            return dateTimeStamp;
        }


   
    }
}
