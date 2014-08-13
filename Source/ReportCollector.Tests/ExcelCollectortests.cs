using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Download;
using Google.Apis.Requests;
using NUnit.Framework;
using ReportDataCollector;
using Should;

namespace ReportCollector.Tests
{
    [TestFixture]
    public class ExcelCollectortests
    {
        [Test]
        public void ShouldConvertXlsToTreportData()
        {
            //var excelDataCollector = new ExcelDataCollector(@"test1.ods");
            var excelDataCollector = new ExcelDataCollector();
            var reportData = excelDataCollector.GetData();
            reportData.FirstName.ShouldEqual("David");
            reportData.Snippets.Count().ShouldEqual(3);
            reportData.Snippets.First().ShouldEqual("Letterhead");
            reportData.Snippets.Take(2).Last().ShouldEqual("PatientInformation");
            reportData.Snippets.Take(3).Last().ShouldEqual("Signature");
        }
    }
}
