using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OTS.Domain;
using ReportDataCollector;
using ReportGeneration;

namespace OTSConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileService = new GoogleFileService();
            if (fileService.HasFileToProcess())
            {
                fileService.ProcessFile();
                ReportData data = new ExcelDataCollector().GetData();
                new AsposeWordGenerator(@"C:\googledrive\templates\input").GenerateReport(data);
                fileService.CreateSessionStructure(data);
            }
        }
    }
}
