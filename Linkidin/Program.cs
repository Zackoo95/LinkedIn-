using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using OpenQA.Selenium.Interactions;

namespace Linkidin
{
    class Program
    {
        private static void NewMethod(IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);
        }

        static void Main(string[] args)
        {

             Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            object misvalue = System.Reflection.Missing.Value;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            // Set a variable to the My Documents path.
            string mydocpath =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            List<IWebElement> optionList = new List<IWebElement>();
            StreamWriter outputFile = new StreamWriter(Path.Combine(mydocpath, "Products.txt"));
            string write;

            var options = new ChromeOptions();
            options.AddArgument("no-sandbox");
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            options.AddAdditionalCapability("capabilityName", "none");
            
            //open chrome web driver
            driver.Navigate().GoToUrl("https://www.linkedin.com");
            IWebElement element = driver.FindElement(By.ClassName("login-email"));
            IWebElement element2 = driver.FindElement(By.ClassName("login-password"));
            element.Click();
            element.SendKeys("");
            element2.Click();
            element2.SendKeys("");
            driver.FindElement(By.Id("login-submit")).Click();

            driver.Navigate().GoToUrl("https://www.linkedin.com/ad-beta/account/507558790/campaign/126895946/details");
            NewMethod(driver);
            driver.FindElement(By.Id("audience-attr-title")).Click();
            driver.FindElement(By.Id("cc-btn-03")).Click();
            NewMethod(driver);
            NewMethod(driver);
            IWebElement element22 = driver.FindElement(By.Id("cc-btn-12"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(element22).Click().Perform();
            element = driver.FindElement(By.ClassName("columns-view__search-input"));
            for (char c = 'A'; c <= 'Z'; c++)
            {
                int x = 1;
                element.SendKeys(c.ToString());
                //element.SendKeys(Keys.Enter);
                System.Threading.Thread.Sleep(5000);
                optionList = driver.FindElements(By.ClassName("u-font__medium--dark")).ToList();
                outputFile.WriteLine(c.ToString() + ":");
                for (int z = 0; z < optionList.Count; z++)
                {
                    oSheet.Cells[x, 1] = optionList[z].Text.ToString();
                    x++;
                }
                for (char i = 'A'; i <= 'Z'; i++)
                {
                    element.SendKeys(Keys.Right);
                    element.SendKeys(i.ToString());
                    //element.SendKeys(Keys.Enter);
                    System.Threading.Thread.Sleep(2000);
                    //optionList = driver.FindElements(By.ClassName("u-font__medium--dark")).ToList();
                    if (driver.FindElements(By.ClassName("u-font__medium--dark")).Count > 0)
                    {
                        optionList = driver.FindElements(By.ClassName("u-font__medium--dark")).ToList();
                        outputFile.WriteLine(c.ToString() + ":");
                        for (int z = 0; z < optionList.Count; z++)
                        {
                            oSheet.Cells[x, 1] = optionList[z].Text.ToString();
                            x++;
                        }
                    }
                    else
                    {
                        element.SendKeys(Keys.Backspace);
                    }
                }
                element.SendKeys(Keys.Backspace);
            }
        }
    }

}