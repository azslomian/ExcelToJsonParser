using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJson
{
    class Program
    {


        static void Main(string[] args)
        {
            ExcelReader excel = new ExcelReader();
            var list = excel.GetExcelFile();
            excel.ConvertListToJson(list);
            Console.WriteLine("Hello World!");
            Console.ReadKey();

            // Go to http://aka.ms/dotnet-get-started-console to continue learning how to build a console app! 
        }
    }
}
