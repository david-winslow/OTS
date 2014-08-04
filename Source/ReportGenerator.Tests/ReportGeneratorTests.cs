using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using ReportGeneration;

namespace ReportGenerator.Tests
{
    [TestFixture]
    public class ReportGeneratorTests
    {
        [Test]
        public void ShouldLoadTemplate()
        {
            WordGenerator g = new WordGenerator("template.docx","testData.xlsx");
        }

        [ExpectedException(typeof(FileNotFoundException))]
        [Test]
        public void ShouldThrowExceptionIFPAthsAreWrong()
        {

            WordGenerator g = new WordGenerator("templatedddd.docx", "testDataddddd.xlsx");

        }
        [Test,Explicit()]
        public void ShouldGeneratePersonalInformation()
        {
            WordGenerator g = new WordGenerator("template.docx", "testData.xlsx");
            g.Generate();
            Thread.Sleep(500);
          //  Process.Start(@"C:\Projects\ReportGenerator\ReportGenerator.Tests\bin\x86\Debug\template.docx");
        }
    }
}
