using AutomationTraining_M7.Base_Files;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTraining_M7.Page_Objects
{
    class LinkedIn_LoginPage : BaseTest
    {
        /*DRIVER REFERENCE FOR POM*/
        private static IWebDriver _objDriver;
        private static WebDriverWait objWait;

        /*LOCATORS FOR EACH ELEMENT*/
        readonly static string STR_USERNAME_TEXT = "username";
        readonly static string STR_PASSWORD_TEXT = "password";
        readonly static string STR_SIGNIN_BTN = "//button[@class='btn__primary--large from__button--floating']";
        
        /*CONSTRUCTOR*/
        public  LinkedIn_LoginPage(IWebDriver pobjDriver)
        {
            _objDriver = pobjDriver;
            objWait = new WebDriverWait(_objDriver, new TimeSpan(0, 0, 30));
        }

        /*IWEBELEMEMT OBJECTS*/
        private static IWebElement objUserNameTxt => _objDriver.FindElement(By.Id(STR_USERNAME_TEXT));
        private static IWebElement objPasswordTxt => _objDriver.FindElement(By.Id(STR_PASSWORD_TEXT));
        private static IWebElement objSignInBtn => _objDriver.FindElement(By.XPath(STR_SIGNIN_BTN));

        /*METHODS*/
        //User Name Txt
        private IWebElement GetUserNameField()
        {
            return objUserNameTxt;
        }

        public static void fnEnterUserName(string pstrUserName)
        {
            objWait.Until(ExpectedConditions.ElementExists(By.Id(STR_USERNAME_TEXT)));
            objWait.Until(ExpectedConditions.ElementIsVisible(By.Id(STR_USERNAME_TEXT)));
            objUserNameTxt.Clear();
            objUserNameTxt.SendKeys(pstrUserName);
        }

        //Password Txt
        private IWebElement GetPassword()
        {
            return objPasswordTxt;
        }

        public static void fnEnterPassword(string pstrPassword)
        {
            objWait.Until(ExpectedConditions.ElementExists(By.Id(STR_PASSWORD_TEXT)));
            objWait.Until(ExpectedConditions.ElementIsVisible(By.Id(STR_PASSWORD_TEXT)));
            objPasswordTxt.Clear();
            objPasswordTxt.SendKeys(pstrPassword);
        }

        //SignIn Button
        private IWebElement GetSignInButton()
        {
            return objSignInBtn;
        }

        public static void fnClickSignInButton()
        {
            objWait.Until(ExpectedConditions.ElementExists(By.XPath(STR_SIGNIN_BTN)));
            objWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(STR_SIGNIN_BTN)));
            objWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(STR_SIGNIN_BTN)));
            objSignInBtn.Click();
        }

        public static void fnLoginLikedInPage(DataTable pTable, string pstrSet)
        {
            try
            {
                if (pTable != null && pTable.Rows.Count > 0)
                {
                    //Iterate each row in Table
                    foreach (DataRow row in pTable.Rows)
                    {
                        if (row["SetValue"].ToString().Trim() == pstrSet)
                        {
                            Assert.AreEqual(true, driver.Title.Contains("Login"), "Title not mach");
                            fnEnterUserName("fdgfdgfd");
                            fnEnterUserName("Test23");

                            fnEnterPassword(row["Password"].ToString().Trim());
                            fnEnterPassword(row["Password"].ToString().Trim());
                            fnClickSignInButton();
                        }


                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Eror in fnLoginLikedInPage: \n" + ex.GetBaseException());
                Assert.Fail();
            }

        }

    }
}
