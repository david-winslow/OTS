using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NGS.Templater;
using ReportDataCollector;

namespace ReportGeneration
{
    public class WordGenerator
    {
        private readonly string _wordtemplateFilename;
        private readonly string _spreadSheetFileName;

        

        public WordGenerator(string wordtemplateFilename, string spreadSheetFileName)
        {
            if(!File.Exists(wordtemplateFilename)) throw new FileNotFoundException("Could not find Word template");
            if (!File.Exists(spreadSheetFileName)) throw new FileNotFoundException("Could not find spreadsheet");
            _wordtemplateFilename = wordtemplateFilename;
            _spreadSheetFileName = spreadSheetFileName;
        }

        public void Generate()
        {
            ReportData reportData = new Linq2ExcelCollector(_spreadSheetFileName).FetchData();

            using (ITemplateDocument templateDocument = Configuration.Factory.Open(_wordtemplateFilename))
            {
                templateDocument.Process(reportData);
            }
        }
    }
}
