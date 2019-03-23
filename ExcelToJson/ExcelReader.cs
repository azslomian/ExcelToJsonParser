using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;

namespace ExcelToJson
{
    public class ExcelReader
    {
        public List<Site> GetExcelFile()
        {
            List<Site> sites = new List<Site>();
            Regex regex = new Regex(@"^(http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Adam\source\repos\ExcelToJson\ExcelToJson\bin\Debug\arkusz4.xlsx");

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            //for(int k=1; k<4; k++)
            //{
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            var asdf = xlWorksheet.Rows;
            Console.WriteLine(xlRange.Count);

            //foreach(var row in xlRange.Rows)
            //{
            //    Console.WriteLine(row.ToString());
            //}

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            var worksheet = xlWorksheet.Name;

            for (int i = 1; i <= 416; i++)
            {
                var site = new Site(i, worksheet, xlRange.Rows[i]);
                Console.WriteLine(i);
                if (site.Link != null)
                {
                    if (regex.IsMatch(site.Link))
                    {
                        sites.Add(site);
                    }
                }
            }
            return sites;
        }

        public void ConvertListToJson(List<Site> sites)
        {
            var json = new JavaScriptSerializer().Serialize(sites);
            File.WriteAllText(@"C:\Users\Adam\linkJsonLogo.json", json);
        }
    }
}
