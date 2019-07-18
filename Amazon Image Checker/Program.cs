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
            int imageLocation = 3;

            while (cellCount != 1790)
            {

                string ASIN = excel.ReadCell(cellCount, 18);

                if (excel.ReadCell(cellCount, 18) == "")
                {
                    cellCount++;
                }

                else
                {
                    
                    //string ASIN = "B07MGZKYJ8";
                    driver.Url = "https://www.amazon.com/dp/" + ASIN;

                    while (IsElementPresent(By.CssSelector("li.a-spacing-small:nth-child(" + imageLocation + ")"), driver) == true)
                    { 
                        //var image = driver.FindElement(By.CssSelector("li.a-spacing-small:nth-child(" + imageLocation + ")"));
                        imageLocation++;
                        imageNumber++;
                        
                    }

                    Console.WriteLine(imageNumber);
                    cellCount++;
                    imageNumber = 0;
                    imageLocation = 3;

                    //something that records imageNumber in a column in excel
                }

            }

            Console.ReadKey();

        }
    }
}
