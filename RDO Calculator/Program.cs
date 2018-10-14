using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace RDO_Calculator
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo TargetFile = new FileInfo
                (@"C:\Users\Deanf\OneDrive\My Documents\Work\### TimeSheets ###\Dean Furley\2017\1-01-2017 Dean Furley.xlsm");
            //DirectoryInfo TargetDirectory = new DirectoryInfo
            //    (@"C:\Users\Deanf\OneDrive\My Documents\Work\### TimeSheets ###\Dean Furley\2018\18-03-2018 Dean Furley.xlsm");
            // WalkDirectoryTree(TargetDirectory);

            Butter.iCount = 0;

            TimeSheet testTimesheet = new TimeSheet();

            // testTimesheet.Load(TargetFile);

            //Console.WriteLine(testTimesheet.ValidFile());

            // Console.WriteLine("Emp Number: " + testTimesheet.EmpNumber.ToString());
            Console.WriteLine("Status: " + testTimesheet.Load(TargetFile).ToString());
            Console.WriteLine("RDO Banked: " + testTimesheet.RdoBanked.ToString());
            Console.WriteLine("RDO Taken: " + testTimesheet.rdoTaken.ToString());
            Console.WriteLine("TimeSheet Date: " + testTimesheet.TsDate.ToString());

            Console.ReadLine();
        }

        static void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {   
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            string sDestinationPath = @"D:\Test\Output";

            // Process all the files in the root directory
            files = root.GetFiles("*.*");

            if (files != null)
            {
                foreach (System.IO.FileInfo element in files)
                {
                    // Console.WriteLine($"{element}");

                    string sFileName = Butter.iCount.ToString() + " " + System.IO.Path.GetFileName(element.Name);
                    string sDestFile = System.IO.Path.Combine(sDestinationPath, sFileName);
                    System.IO.File.Copy(element.FullName, sDestFile, true);

                    // Console.WriteLine(element.FullName);

                    Butter.iCount++;
                }

                // Now find all thne subdirectories under this directory.
                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    // Resursive call for each subdirectory.
                    WalkDirectoryTree(dirInfo);
                }
            }
        }
    }
    static class Butter
    {
        public static int iCount;
    }      
    public class TimeSheet
    {
        private double pRdoBanked;
        private double pRdoTaken;
        //private int pEmpNumber;
        private DateTime pTsDate;
        // System.IO.DirectoryInfo
        public bool Load(System.IO.FileInfo pFilePath)
        {
            Excel.Application excelApp = new Excel.Application();
            // Set Work Book
            Excel.Workbook wb = excelApp.Workbooks.Open(pFilePath.ToString());
            // Set Work Sheet
            Excel.Worksheet ws = wb.Sheets[1];
            // Set Range            
            Excel.Range rng = ws.Cells[6, 1];

            // Check certain cells in the target to prove that the files is a correct time sheet
            //    /*
            //     * The below are other conditions to check via nestled if statements
            //     * to make the program more robust
            //     * 
            //     * 'rng = ws.Cells(6, 1)
            //     * 'ws.Cells(6, 1).Text.ToString.Contains("Employee Name:")
            //     * 'rng = ws.Cells(7, 1)
            //     * 'ws.Cells(7, 1).Text.ToString.Contains("State:")
            //     * 'rng = ws.Cells(8, 1)
            //     * 'ws.Cells(8, 1).Text.ToString.Contains("Employee Number:")
            //    */

            /*
             * Check for "Employee Name in the timetime, this is to firstly check that thi is a valid
             * time sheet and in the correct format, there are additionl conditons above that should
             * be added at a later date so as to make this test more robust
            */
            //rng.ToString.c

            string s = rng.Text.ToString();

            if (s.Contains("Employee Name:"))
            {
                rng = ws.Cells[6, 8];
                pTsDate = DateTime.Parse(rng.Text.ToString());

                int X = 1;
                int Y = 1;
                rng = ws.Cells[Y, X];

                /*
                 * Find the string "Payroll Summary" in the timesheet, this is for the purpose of
                 * locating the correct cells, as the information in not in fixed cell location as 
                 * depending on the amount of entries it is lower down the page, this finds the point
                 * to refernce from
                */
                
                while (rng.Text.ToString().Contains("Payroll Summary:") != (true))
                    {
                    Y = Y + 1;
                    rng = ws.Cells[Y, X];
                    }

                // Increment counter to the line below where the above text is found
                Y = Y + 1;

                // Find the "Normal Hrs" Header, this is to set the Y coodrinate. 
                while (rng.Text.ToString().Contains("Normal Hrs") != (true))
                {
                    X = X + 1;
                    rng = ws.Cells[Y, X];
                }

                rng = ws.Cells[Y + 1, X];

                double temp = double.Parse(rng.Text.ToString());
                pRdoBanked = Math.Ceiling(temp / 8) * 0.4;
                // Console.WriteLine("RDO Hours Banked: " + pRdoBanked.ToString());
                
                // RDO Hours Taken
                while (rng.Text.ToString().Contains("RDO Hrs Taken") != (true))
                {
                    X = X + 1;
                    rng = ws.Cells[Y, X];
                }

                rng = ws.Cells[Y + 1, X];
                // pRdoTaken = double.Parse(rng.Text.ToString());

                // values.valid = rdoBanked;
                //values.hoursBanked = rdoBanked;
                //values.hoursUsed = rdoTaken;
                //values.timesheetDate = tsDate;

                wb.Close();
                return (true);
            }
            else
            {
                wb.Close();
                return (false);
            }
        }

        public double RdoBanked
        {
            get
            {
                return (pRdoBanked);
            }
            set { /* Some logic */ }
        }
        public double rdoTaken
        {
            get
            {
                return (pRdoTaken);
            }
            set { /* Some logic */ }
        }
        
        //public int EmpNumber
        //{
        //    get
        //    {
        //        return (pEmpNumber);
        //    }
        //    set
        //    { /* Some logic */
        //    }
        //}

        public DateTime TsDate
        {
            get
            {
                return (pTsDate);
            }
            set { /* Some logic */ }
        }
    }
}


