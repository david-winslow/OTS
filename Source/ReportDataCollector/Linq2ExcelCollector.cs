using System.IO;
using System.Linq;
using LinqToExcel;

namespace ReportDataCollector
{
    public interface IExcelCollector
    {
        ReportData FetchData();
    }

    public class Linq2ExcelCollector : IExcelCollector
    {
        private ExcelQueryFactory _excel;

        public Linq2ExcelCollector(string fileName)
        {
            if(!File.Exists(fileName)) throw new FileNotFoundException("could not find input xlsx file.Verify file exist!");
            _excel = new ExcelQueryFactory(fileName);
        }

        public ReportData FetchData()
        {
            var ClientInformation = _excel.WorksheetRange<Client>("A2", "K3","Personal Information").Select(c => c).First();
            var background = _excel.WorksheetRange<Background>("A5", "g6","Personal Information").Select(c => c).First();
            var occupation = _excel.WorksheetRange<Occupation>("A12", "E13","Personal Information").Select(c => c).First();
            var jobDescriptions = _excel.WorksheetRangeNoHeader("A17", "A37", "Personal Information").Select(c => c[0].Value as string).ToList();
            occupation.JobDescription = jobDescriptions.Where(s => !string.IsNullOrEmpty(s)).ToList();
            return new ReportData(ClientInformation,background,occupation);
        }
    }
}