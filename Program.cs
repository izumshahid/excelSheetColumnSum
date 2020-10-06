using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace BSF
{
    class Program
    {
        static void Main( string[] args )
        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "files\\BSFTest.xlsx";
            //if file dont exist display error and close program
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Requested File not does not exist at path : " + filePath);
                Console.Read();
                Environment.Exit(0);
            }

            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);// place the files with exe in files 
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;
            try
            {                                 
                //give total number of rows, give the colum number you want to calculate sum
                int rowCount = 19, colCount = 3;
                // variable to check of the value coming from the cell is double or not
                double sum = 0, isDouble;

                //starting from 2 because first line is for header.
                for (int i = 2; i <= rowCount; i++)
                {
                    //printing i value just for display so user can see if program is running or not.
                    Console.WriteLine(i);
                    // assigning j = colCount; so only will run for that specific column
                    for (int j = colCount; j <= colCount; j++)// put 6 for BSF file and 3 for testing
                    {
                        //getting cell value
                        var x = xlRange.Cells[i , j].Value2.ToString();
                        //checking if its double or not
                        if (Double.TryParse(x , out isDouble))
                        {
                            //if double calculate sum
                            sum += Convert.ToDouble(x);
                        }
                        else
                        {
                            //if not double value then show it on console
                            Console.WriteLine("line number : " + i + " value is not Digit");
                            Console.Write("\r\n");
                        }
                    }
                }

                Console.WriteLine("Total Sum = " + sum);

                //close the workbook and quit the handler or xcel file will not be even after the program is finished
                xlWorkbook.Close();
                xlApp.Quit();
                Console.Read();
            }
            catch (Exception)
            {
                //close the workbook and quit the handler or xcel file will not be even after the program is finished
                xlWorkbook.Close();
                xlApp.Quit();
                throw;
            }
        }
    }
}
