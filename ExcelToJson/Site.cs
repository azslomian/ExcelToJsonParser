using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJson
{
    public class Site
    {
        private dynamic dynamic;

        public Site(int id, string worksheet, dynamic dynamic)
        {
            Worksheet = worksheet;
            this.dynamic = dynamic;

            Worksheet = worksheet;
            Lp = id;
            Level1 = dynamic.Columns[1].Value2;
            Level2 = dynamic.Columns[2].Value2;
            Level3 = dynamic.Columns[3].Value2;
            Level4 = dynamic.Columns[4].Value2;
            Level5 = dynamic.Columns[5].Value2;
            Link = dynamic.Columns[7].Value2;
            Logo = dynamic.Columns[8].Value2;
        }

        public string Worksheet { get; set; }
        public int Lp { get; set; }
        public dynamic Level1 { get; set; }
        public dynamic Level2 { get; set; }
        public dynamic Level3 { get; set; }
        public dynamic Level4 { get; set; }
        public dynamic Level5 { get; set; }
        public dynamic Link { get; set; }
        public dynamic Logo { get; set; }
    }
}
