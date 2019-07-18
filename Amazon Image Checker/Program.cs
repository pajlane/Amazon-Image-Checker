using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace ExcelPractice_
{
    class Program
    {

        private static bool IsElementPresent(By by, IWebDriver driver)
        {                                                       
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (ElementNotVisibleException)
            {
                return false;
            }
        }


            public static void Main(string[] args)
        {

            var driverService = ChromeDriverService.CreateDefaultService(@"\\brmpro\MACAPPS\ClickOnce\CustomerServiceAutomationTool");
            driverService.HideCommandPromptWindow = true;
            var driver = new ChromeDriver(driverService, new ChromeOptions());

            


            Excel excel = new Excel(@"C:\Users\plane\Desktop\PDTest.xlsx", 1);
            int cellCount = 2;
            int imageNumber = 0;
            int imageCheck = 3;

            while (cellCount != 1790)
            {
                if (excel.ReadCell(cellCount, 18) == "")
                {
                    cellCount++;
                }
                else
                {
                    //string ASIN = excel.ReadCell(cellCount, 18);
                    string ASIN = excel.ReadCell(6, 18);
                    driver.Url = "https://www.amazon.com/dp/" + ASIN;

                    while (IsElementPresent(By.CssSelector("li.a-spacing-small:nth-child(" + imageCheck + ")"), driver) == false);
                    { 
                        var image = driver.FindElement(By.CssSelector("li.a-spacing-small:nth-child(" + imageCheck + ")"));
                        imageCheck++;
                        imageNumber++;
                        }

                    cellCount++;
                    Console.WriteLine(imageNumber);
                }

                //something that records imageNumber in a column in excel

            }

            Console.ReadKey();

        }
    }
}
