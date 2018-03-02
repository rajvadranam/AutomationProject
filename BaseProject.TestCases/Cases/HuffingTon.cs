/*
 * Created by SharpDevelop.
 * User: E001011
 * Date: 01-08-2016
 * Time: 17:43
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using BaseProject;
using BaseProject.Pages;
using System.Configuration;
using System.Data;
using OpenQA.Selenium;

namespace BaseProject.Tests.RegressionSuite
{ 
	
    class HuffingTon : BaseCase
    {
        protected override void ExecuteTestCase()
        {
            base.ExecuteTestCase();
            Reporter.Chapter.Title = "Reading stories from the websites";
            Step = "Read the stories for given website";
             Reporter.Chapter.Title = "Today's "+TestData["url"]+" all stories in the site";
            Step = "Navigate to website";
            Common.NavigateTo(Driver, Reporter, TestData["url"]);
            Step = "Read stories on website";
            DataExtractCrawler.Huffstories(Driver, Reporter, TestData["url"]);
            //(//*[contains(@class,'article-meta')])[1]
            
        }

    }
}
