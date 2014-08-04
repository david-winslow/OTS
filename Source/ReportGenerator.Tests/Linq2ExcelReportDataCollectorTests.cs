using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using NUnit.Framework.Constraints;
using ReportDataCollector;

namespace ReportGenerator.Tests
{

  
  

    [TestFixture]
    public class Linq2ExcelReportDataCollectorTests
    {
        private string _dataPath = "testdata.xlsx";
        private IExcelCollector _collector;

        [Test]
        public void ShouldFetchSpreadSheetBasedOnPath()
        {
            _collector = new Linq2ExcelCollector(_dataPath);
        }



        [ExpectedException(typeof(FileNotFoundException))]
        [Test]
        public void ShouldTHrowExceptionIfSpreadSheetIsMissing()
        {
            _collector = new Linq2ExcelCollector("bogus");

        }

        [Test]
        public void ShouldFetchPersonalInformationFromSpreadSheet()
        {
             _collector = new Linq2ExcelCollector(_dataPath);
            ReportData data  = _collector.FetchData();
            Assert.AreEqual("Mr",data.Client.Title);
            Assert.AreEqual("David",data.Client.FirstName);
            Assert.AreEqual("Winslow",data.Client.LastName);
            Assert.AreEqual("20 elm avenue", data.Client.Address);
            Assert.AreEqual("082 837 4699",data.Client.CellNumber);
            Assert.AreEqual("0116081447",data.Client.LandLine);
            Assert.AreEqual(new DateTime(1975,7,12), data.Client.DOB);
            Assert.AreEqual("7507126057089", data.Client.IDNumber);
            Assert.AreEqual("Afrikaans",data.Client.HomeLanguage);
            Assert.AreEqual("English",data.Client.AssessmentLanguage);
            Assert.AreEqual(true,data.Client.DriverLicense);
        }

        [Test]
        public void ShouldFetchBackgroundInformation()
        {
            _collector = new Linq2ExcelCollector(@"testdata.xlsx");
            var b = _collector.FetchData().Background;
            Assert.AreEqual("married",b.MaritalStatus);
            Assert.AreEqual(2,b.Children);
            Assert.AreEqual(4,b.Siblings);
            Assert.AreEqual("freestanding house",b.ResidenceType);
            Assert.AreEqual(true,b.HasCar);
            Assert.AreEqual(true,b.IsDriving);
            Assert.AreEqual("husband",b.AlternativeDriver);
        }



        [Test]
        public void ShouldFetchOccupationalInformation()
        {
            Linq2ExcelCollector collector = new Linq2ExcelCollector(_dataPath);
            var o = collector.FetchData().Occupation;
            Assert.AreEqual("Senior Personnel Practitioner – Police Officer", o.OccupationAtTimeOfIllness);
            Assert.AreEqual("SAPS", o.EmployerAtTimeOfIllness);
            Assert.AreEqual("Vispol Support Officer ", o.CurrentOccupation);
            Assert.AreEqual("SAPS", o.CurrentEmployer);
        }

        [Test]
        public void ShouldFetchJobDescriptionAndAddToOccupation()
        {
            Linq2ExcelCollector collector = new Linq2ExcelCollector(_dataPath);
            var o = collector.FetchData().Occupation;
            Assert.AreEqual(14,o.JobDescription.Count);
        }

        [Test]
        public void ShouldDetermineGenderBasedOnIdNumber()
        {
            string ID = "7507126057089";
            var client = new Client(){IDNumber = ID };
           Assert.AreEqual("Male",client.Gender);
            Assert.AreEqual("He",client.HeShe);
        }

        [Test]
        public void ShouldMapRowsAsNames()
        {
            
        }
    }
}
