using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace RDO_Calculator
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo TargetDirectory = new DirectoryInfo(@"C:\Users\Deanf\OneDrive\### Holding Bay ###\Test\Data");
            WalkDirectoryTree(TargetDirectory);
        }

        static void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            string sDestinationPath = @"C:\Users\Deanf\OneDrive\### Holding Bay ###\Test\Output";
            
            //Process all the files in the root directory
            files = root.GetFiles("*.*");

            if (files != null)
            {
                foreach (System.IO.FileInfo element in files)
                {
                    Console.WriteLine($"{element}");

                    /*
                    Set filename currently uses the current, but is setup this way so that a an index or such can be 
                    added so that the program can handle files with the same name in diffrent locations
                    */

                    string sFileName = System.IO.Path.GetFileName(element.Name);
                    string sDestFile = System.IO.Path.Combine(sDestinationPath, sFileName);
                    System.IO.File.Copy(element.FullName, sDestFile, true);
                    
                    Console.WriteLine(element.FullName);
                    
                }

                // Now find all the subdirectories under this directory.
                subDirs = root.GetDirectories();

                foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                {
                    // Resursive call for each subdirectory.
                    WalkDirectoryTree(dirInfo);
                }
            }
        
        }
    }
}


