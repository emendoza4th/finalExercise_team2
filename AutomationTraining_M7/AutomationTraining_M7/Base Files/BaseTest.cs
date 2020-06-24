
using AutomationTraining_M7.Reporting;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTraining_M7.Base_Files
{
    class BaseTest
    {
        /*Webdriver Intance*/
        public static clsDriver objclsDriver;
        public static IWebDriver driver;
        /*URL for Webdriver*/
        private static string strBrowserName = ConfigurationManager.AppSettings.Get("url");
        /*Extent Reports Framework*/
        public static clsReportManager objRM;
        public static ExtentV3HtmlReporter objHtmlReporter; //Add information in HTML
        public static ExtentReports objExtent; //Extent Reports Object
        public static ExtentTest objTest; // Test object for Extent Reports

        [OneTimeSetUp]
        public static void fnBeforeClass()
        {
            objRM = new clsReportManager();
            /*Init ExtentHtmlReporter object*/
            if (objHtmlReporter == null)
            {
                objHtmlReporter = new ExtentV3HtmlReporter (objRM.fnReportPath());
            }
            /*Init ExtentReports object*/
            if (objExtent == null)
            {
                objExtent = new ExtentReports();
                objRM.fnReportSetUp(objHtmlReporter, objExtent);
            }
        }

        //OneTimeTearDown after each class test
        [OneTimeTearDown]
        public static void fnAfterClass()
        {
            objExtent.Flush();
        }

        //SetUp Before each test case
        [SetUp]
        public static void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = strBrowserName;
            driver.Manage().Window.Maximize();
            objclsDriver = new clsDriver(driver);

        }

        [TearDown]
        //TearDown After each test case
        public static void AfterTest()
        {
            /*objRM.fnTestCaseResult(objTest, objExtent, driver);
            driver.Close();
            driver.Quit();*/
        }

        /*Clear and Send text to specific field*/
        public static void FnSendkeyAndClear(By by, string pstrText)
        {
            driver.FindElement(by).Clear();
            driver.FindElement(by).SendKeys(pstrText);
        }
    }
}