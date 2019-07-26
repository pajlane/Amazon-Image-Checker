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


            Dictionary<int, string> ASINREF = new Dictionary<int, string>();

            for (int i = 2; i < 1790; i++)
            {
                string ASIN = excel.ReadCell(i, 18);
                
                ASINREF.Add(i, ASIN);        
            }

            foreach (KeyValuePair<int, string> bubs in ASINREF)
            {
                Console.WriteLine("Key: {0}, Value: {1}",
                bubs.Key, bubs.Value);
            }

            while (cellCount != 1790)
            {
                imageNumber = 0;

                if (ASINREF[cellCount] != "")
                {
                    driver.Url = "https://www.amazon.com/dp/" + ASINREF[cellCount];

                    while (IsElementPresent(By.CssSelector("li.a-spacing-small:nth-child(" + imageLocation + ")"), driver) == true)
                    { 
                        imageLocation++;
                        imageNumber++; 
                    }

                    Console.WriteLine(imageNumber);
                    
                    imageLocation = 3; 
                }

                excel.WriteToCell(cellCount, 19, imageNumber);

                cellCount++;
            }

            excel.Save();
            excel.Quit();
            driver.Quit();

        }
    }
}
