using AutomationTraining_M7.Page_Objects;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutomationTraining_M7.Base_Files;
using AutomationTraining_M7.Data_Model;
using System.IO;


namespace AutomationTraining_M7.Test_Cases
{
    class Test_LinkedInSearch : Test_LinkedIn
    {
        //LinkedIn_LoginPage objLogin; -- DELETE
        public WebDriverWait wait;
        LinkedIn_SearchPage objSearch;
        //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
        
        [Test]
        public void Search_LinkedIn()
        {
            try
            {
                //VARIABLES
                //Report object.
                objTest = objExtent.CreateTest(TestContext.CurrentContext.Test.Name);
                //Path to technologies file
                string userpath = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
                if (Environment.OSVersion.Version.Major >= 6)
                {
                    userpath = Directory.GetParent(userpath).ToString();
                }

                string filepath = userpath + "\\Documents\\technologies.txt";
                string[] arrLines = System.IO.File.ReadAllLines(filepath);
                //Check if the file exists, if not create it and write alert
                if (File.Exists(filepath))
                {
                    Console.WriteLine($"File was found in {filepath} the content is:\n");

                    String fileContent = System.IO.File.ReadAllText(filepath);
                    if (fileContent.Equals("Replace this text with the list of technologies you want to search candidates for."))
                    {
                        Console.WriteLine("File has no technologies to search, please replace them.");
                        return;
                    }
                    else
                    {
                        foreach (string line in arrLines)
                        {
                            Console.WriteLine("\t" + line);
                        }
                    }
                }
                else
                {
                    using (FileStream fs = File.Create(filepath))
                    {
                        byte[] file = new UTF8Encoding(true).GetBytes("Replace this text with the list of technologies you want to search candidates for.");
                        fs.Write(file, 0, file.Length);
                    }
                    Console.WriteLine($"The input file for tech nologies was not found, please go to {filepath} and update the file contents.");

                    return;
                }

                //Step# 1 .- Log In 
                objSearch = new LinkedIn_SearchPage(driver);
                Login_LinkedIn();
                objRM.fnAddStepLogWithSnapshot(objTest, driver, "Login done", $"{ DateTime.Now.ToString("HHmmss")}.png", "Pass");

                //Step# 2 .- Verify if captcha exist
                if (driver.Title.Contains("Verification") | driver.Title.Contains("Verificación"))
                {
                    //Switch to Iframe(0)
                    driver.SwitchTo().DefaultContent();
                    driver.SwitchTo().Frame(driver.FindElement(By.Id("captcha-internal")));
                    //Switch to Iframe that contains captcha.
                    IWebElement objCheckbox;
                    List<IWebElement> frames = new List<IWebElement>(driver.FindElements(By.TagName("iframe")));
                    for (int i = 0; i < frames.Count - 1; i++)
                    {
                        if (frames[i].GetAttribute("role").ToString() == "presentation" | frames[i].GetAttribute("role").ToString() != "")
                        {
                            driver.SwitchTo().Frame(i);
                            wait = new WebDriverWait(driver, new TimeSpan(0, 0, 60));
                            objCheckbox = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//span[@role='checkbox']")));
                            if (objCheckbox.Enabled) { objCheckbox.Click(); }

                        }
                    }
                }

                //Step# 3 .- Set Filters
                for (int i = 0; i < arrLines.Length; i++)
                {
                    objSearch = new LinkedIn_SearchPage(driver);
                    LinkedIn_SearchPage.fnEnterSearchText(arrLines[i]);
                    LinkedIn_SearchPage.fnClickSearchBtn();
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "Technology search.", "Search.png", "Pass");
                    ExportDataCsv file = new ExportDataCsv(arrLines[i]);
                    wait = new WebDriverWait(driver, new TimeSpan(0, 1, 0));
                    wait.Until(ExpectedConditions.UrlContains("search/results"));
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//button[span[text()='People' or text()='Gente']]")));

                    //Step# 4 .- Selecting People button
                    try
                    {
                        LinkedIn_SearchPage.fnSelectPeople();
                    }
                    catch (StaleElementReferenceException)
                    {
                        LinkedIn_SearchPage.fnSelectPeople();
                    }
                    wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@class='search-vertical-filter__dropdown-trigger-text mr1'][text()='People' or text()='Gente']")));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[span[text()='All Filters' or text()='Todos los filtros']]")));
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "People filtered", $"{arrLines[i]}_People_{DateTime.Now.ToString("HHmmss")}.png", "Pass");

                    //Step# 5 .- Locations selection
                    LinkedIn_SearchPage.fnSelectAllFilters();
                    wait.Until(ExpectedConditions.ElementExists(By.XPath("//input[@placeholder='Add a country/region' or @placeholder='Añadir un país o región'][@aria-label='Add a country/region' or @aria-label='Añadir un país o región']")));
                    LinkedIn_SearchPage.fnAddCountry("Mexico");
                    try
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@class='search-basic-typeahead search-vertical-typeahead ember-view']//*[@class='basic-typeahead__selectable ember-view']//span[text()= 'Mexico' or text()='México']")));
                        LinkedIn_SearchPage.fnSelectMexico();
                    }
                    catch (StaleElementReferenceException)
                    {
                        
                        wait.Until(ExpectedConditions.StalenessOf(LinkedIn_SearchPage.SelectMexico()));
                        LinkedIn_SearchPage.fnSelectMexico();
                    }
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "Select Country", $"{arrLines[i]}_Location_{DateTime.Now.ToString("HHmmss")}.png", "Pass");
                    

                    //Step 6 .- Language selection.
                    LinkedIn_SearchPage.fnLanguageEng();
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "Language selection", $"{arrLines[i]}_Language_{DateTime.Now.ToString("HHmmss")}.png", "Pass");

                    //Step# 7 .- Apply the Filters
                    LinkedIn_SearchPage.fnClickApplyBtn();
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "Filters Applied", $"{arrLines[i]}_Filters_{DateTime.Now.ToString("HHmmss")}.png", "Pass");
                    Thread.Sleep(5000);
                    


                    IList<IWebElement> allSearchResults = LinkedIn_SearchPage.fnAllResultPage();

                    

                    //Step 8 .- Get the People information
                    for (int b = 0; b < allSearchResults.Count; b++)
                    {
                        wait.Until(ExpectedConditions.ElementIsVisible (By.XPath($"(//span[@class='actor-name' or @class='name actor-name'])[{b + 1}]")));
                        wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath($"(//span[@class='actor-name' or @class='name actor-name'])[{b + 1}]")));
                        string listXpath = $"(//span[@class='actor-name' or @class='name actor-name'])[{b + 1}]";
                        string STR_TOTAL_RESULTS_WO = listXpath;
                        wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(STR_TOTAL_RESULTS_WO)));
                        IWebElement objSearchResult = driver.FindElement(By.XPath(listXpath));
                        
                        objSearchResult.Click();
                        Thread.Sleep(3000);
                        
                        LinkedIn_SearchPage.fnScrollDownResults();
                        
                        IList<IWebElement> PopUpBtn = driver.FindElements(By.XPath("//span[text()='Volver']"));
                        
                        if (PopUpBtn.Count() > 0)
                            {
                                LinkedIn_SearchPage.fnClickPopUpBtn();
                            }
                        else
                            {
                                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[text()='Información de contacto']")));
                                LinkedIn_SearchPage.fnScrollDownToSkills();
                                file.Member.Add(LinkedIn_SearchPage.fnMemberInfo());
                                if (arrLines[i].Contains("#")) { arrLines[i] = "CSharp"; }
                                objRM.fnAddStepLogWithSnapshot(objTest, driver, "Data from Contact stored", $"{arrLines[i]}_Data_{DateTime.Now.ToString("HHmmss")}.png", "Pass");
                                driver.Navigate().Back();
                            }
                        
 

                    }
                    
                    file.fnCreateFile(file.Member);

                    LinkedIn_SearchPage.fnClearFilters();
                    objRM.fnAddStepLogWithSnapshot(objTest, driver, "Filters cleared", "Home.png", "Pass");

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }
    }
}