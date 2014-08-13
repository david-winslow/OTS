using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using LinqToExcel;
using OTS.Domain;

namespace ReportDataCollector
{
 

    public class ExcelDataCollector
    {
        private readonly string _path = "c:\\googledrive\\input.ods";

        public ExcelDataCollector()
        {
        }

        public ReportData GetData()
        {
            var workbook = new Workbook(_path);
            var c = workbook.Worksheets[0].Cells;

            var snippets = GetList(c,"D",9,3);
            return new ReportData() { Date = c["B1"].DateTimeValue, FirstName = c["b3"].StringValue, Surname = c["b2"].StringValue, Therapist = c["b4"].StringValue, TemplateName = c["b5"].StringValue,Snippets = snippets };
        }

        private static IEnumerable<string> GetList(Cells cells,string startColumn,int startRow,int length)
        {

            for (int i = 0; i < length; i++)
            {
                yield return cells[string.Format("{0}{1}", startColumn, startRow + i)].StringValue;
            }
        }
    }

   


}
