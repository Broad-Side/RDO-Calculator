using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace RDO_Calculator
{
    static class Butter
    {
        public static int iCount;
    }

    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo TargetDirectory = new DirectoryInfo(@"D:\Test\Data");
            WalkDirectoryTree(TargetDirectory);

            Butter.iCount = 0;
        }

        static void WalkDirectoryTree(System.IO.DirectoryInfo root)
        {
            System.IO.FileInfo[] files = null;
            System.IO.DirectoryInfo[] subDirs = null;

            string sDestinationPath = @"D:\Test\Output";
            
            //Process all the files in the root directory
            files = root.GetFiles("*.*");

            if (files != null)
            {
                foreach (System.IO.FileInfo element in files)
                {
                    Console.WriteLine($"{element}");

                    string sFileName = Butter.iCount.ToString() + " " + System.IO.Path.GetFileName(element.Name);
                    string sDestFile = System.IO.Path.Combine(sDestinationPath, sFileName);
                    System.IO.File.Copy(element.FullName, sDestFile, true);
                    
                    Console.WriteLine(element.FullName);

                    Butter.iCount++;                                                         
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


