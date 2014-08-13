using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OTS.Domain;
using ReportGeneration;

namespace ReportGenerator.Tests
{
    [TestFixture]
   public class WordGeneratorTests
    {
        [Test]
        public void ShouldMergeDocuments()
        {
            AsposeWordGenerator generator = new AsposeWordGenerator(@"c:\googledrive\templates");
            generator.GenerateReport(new ReportData()
            {
                Date = new DateTime(2014, 1, 1),
                Surname = "surname",
                FirstName = "David",
                Therapist = "Aislinn",
                TemplateName = "master",
                Snippets = new List<string>() {"Letterhead","PatientInfo","Signature"}
            });

        }



    }
}
