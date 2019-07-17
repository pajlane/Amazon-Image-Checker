using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPractice_
{
    class Program
    {
        public static void Main(string[] args)
        {

            Excel excel = new Excel(@"C:\Users\plane\Desktop\PDTest.xlsx", 1);
            int cellCount = 2;

            while (cellCount != 1790)
            {
                if (excel.ReadCell(cellCount, 18) == "")
                {
                    cellCount++;
                }
                else
                {


                    Console.WriteLine(excel.ReadCell(cellCount, 18));
                    cellCount++;
                }

            }


            Console.ReadKey();

        }
    }
}
