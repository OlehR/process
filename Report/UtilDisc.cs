using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report
{
    public class UtilDisc
    {
        static public bool  MoveAllFilesMask(string sourcePath,string sourceFiles , string targetPath)
        {
            //tring sourcePath = @"C:\Source Folder";
            //string targetPath = @"D:\Destination Folder";

            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }

            string[] sourcefiles = Directory.GetFiles(sourcePath, sourceFiles);

            foreach (string sourcefile in sourcefiles)
            {
                string fileName = Path.GetFileName(sourcefile);
                string destFile = Path.Combine(targetPath, fileName);

                File.Move(sourcefile, destFile);
            }
            return true;

        }
    }
}
