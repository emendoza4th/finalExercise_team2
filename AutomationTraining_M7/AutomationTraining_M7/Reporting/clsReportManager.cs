using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTraining_M7.Reporting
{
    class clsReportManager
    {
        private DateTime time = DateTime.Now;

        public string fnReportPath()
        {
            var strPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            var strActualPath = strPath.Substring(0, strPath.LastIndexOf("bin"));
            var strProjectPath = new Uri(strActualPath).LocalPath;
            Directory.CreateDirectory(strProjectPath.ToString() + "ExtentReports");
            var strReportPath = strProjectPath + "ExtentReports\\ExtentReports_" + time.ToString("MMddyyyy_HHmmss") + ".html";
            return strReportPath;
        }

        //Function to add a directiory
        public string fnGetCSVPath()
        {
            var strPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            var strActualPath = strPath.Substring(0, strPath.LastIndexOf("bin"));
            var strProjectPath = new Uri(strActualPath).LocalPath;
            Directory.CreateDirectory(strProjectPath.ToString() + "CsvOutput");
            var strReportPath = strProjectPath + "CsvOutput\\";
            return strReportPath;
        }

        public void fnReportSetUp(ExtentV3HtmlReporter phtmlReporter, ExtentReports pExtent)
        {
            phtmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;
            phtmlReporter.Config.DocumentTitle = "Automation Framework Report";
            pExtent.AttachReporter(phtmlReporter);
            pExtent.AddSystemInfo("Project Name:", "Automation Framework");
            pExtent.AddSystemInfo("Application:", "LinkedIn");
            pExtent.AddSystemInfo("Environment:", "QAA");
            pExtent.AddSystemInfo("Browser:", ConfigurationManager.AppSettings.Get("browser"));
            pExtent.AddSystemInfo("Date:", time.ToShortDateString());
            pExtent.AddSystemInfo("Version:", "v1.0");
        }

        public string fnCaptureImage(IWebDriver pobjDriver, string pstrScreenName)
        {
            /*Create SS directory*/
            ITakesScreenshot objITake = (ITakesScreenshot)pobjDriver;
            Screenshot objSS = objITake.GetScreenshot();
            var strSSPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            var strActualPath = strSSPath.Substring(0, strSSPath.LastIndexOf("bin"));
            var strReportPath = new Uri(strActualPath).LocalPath;
            Directory.CreateDirectory(strReportPath.ToString() + "ExtentReports\\Screenshots");
            /*Save Image*/
            var strFullPath = strSSPath.Substring(0, strSSPath.LastIndexOf("bin")) + "ExtentReports\\Screenshots\\" + pstrScreenName;
            var strLocalPath = new Uri(strFullPath).LocalPath;
            objSS.SaveAsFile(strLocalPath, ScreenshotImageFormat.Png);
            return strLocalPath;
        }

        public void fnTestCaseResult(ExtentTest pobjTest, ExtentReports pobjExtent, IWebDriver pobjDriver)
        {
            var status = TestContext.CurrentContext.Result.Outcome.Status;
            var stacktrace = string.IsNullOrEmpty(TestContext.CurrentContext.Result.StackTrace)
           ? "" : string.Format("{0}", TestContext.CurrentContext.Result.StackTrace);
            Status logstatus;
            switch (status)
            {
                case TestStatus.Failed:
                    logstatus = Status.Fail;
                    string strFileName = "Screenshot_" + time.ToString("hh_mm_ss") + ".png";
                    var strImagePath = fnCaptureImage(pobjDriver, strFileName);
                    pobjTest.Log(Status.Fail, "Fail");
                    pobjTest.Fail("Snapshot below: ", MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case TestStatus.Skipped:
                    logstatus = Status.Skip;
                    break;
                case TestStatus.Passed:
                    logstatus = Status.Pass;
                    break;
                default:
                    logstatus = Status.Warning;
                    Console.WriteLine("The status: " + status + " is not supported.");
                    break;
            }
            pobjTest.Log(logstatus, "Test ended with " + logstatus + stacktrace);
            pobjExtent.Flush();
        }


        public void fnAddStepLog(ExtentTest pobjTest, string pstrMessage, string pStatus)
        {
            pobjTest.Log(Status.Info, pstrMessage);
            switch (pStatus.ToUpper())
            {
                case "PASS":
                    pobjTest.Log(Status.Pass, pstrMessage);
                    break;
                case "ERROR":
                    pobjTest.Log(Status.Error, pstrMessage);
                    break;
                case "SKIPT":
                    pobjTest.Log(Status.Skip, pstrMessage);
                    break;
                case "WARNING":
                    pobjTest.Log(Status.Warning, pstrMessage);
                    break;
                case "INFO":
                    pobjTest.Log(Status.Info, pstrMessage);
                    break;
                case "FAIL":
                    pobjTest.Log(Status.Fail, pstrMessage);
                    break;
                case "FATAL":
                    pobjTest.Log(Status.Fatal, pstrMessage);
                    break;
                case "DEBUG":
                    pobjTest.Log(Status.Debug, pstrMessage);
                    break;
                default:
                    pobjTest.Log(Status.Info, pstrMessage);
                    break;
            }

        }

        public void fnAddStepLogWithSnapshot(ExtentTest pobjTest, IWebDriver pobjDriver, string pstrMessage, string pstrImageName, string pStatus)
        {
            var strImagePath = fnCaptureImage(pobjDriver, pstrImageName);
            switch (pStatus.ToUpper())
            {
                case "PASS":
                    pobjTest.Pass(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "ERROR":
                    pobjTest.Error(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "SKIPT":
                    pobjTest.Skip(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "WARNING":
                    pobjTest.Warning(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "INFO":
                    pobjTest.Info(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "FAIL":
                    pobjTest.Fail(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "FATAL":
                    pobjTest.Fatal(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                case "DEBUG":
                    pobjTest.Debug(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
                default:
                    pobjTest.Info(pstrMessage, MediaEntityBuilder.CreateScreenCaptureFromPath(strImagePath).Build());
                    break;
            }
        }


    }
}
