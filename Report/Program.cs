using System;
using System.Reflection;
using System.IO;

namespace Report
{
   
   
    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel();
            excel.ExecuteExcelsMacro(args[0]);           
        }
    }
}
