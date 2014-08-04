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
    public class ReportDataCollectorTests
    {
        private string _dataPath = "testdata.xlsx";

        [Test]
        public void ShouldFetchSpreadSheetBasedOnPath()
        {
            ExcelCollector collector = new ExcelCollector(_dataPath);
            
        }

        [ExpectedException(typeof(FileNotFoundException))]
        [Test]
        public void ShouldTHrowExceptionIfSpreadSheetIsMissing()
        {
            ExcelCollector collector = new ExcelCollector("bogus");

        }

        [Test]
        public void ShouldFetchPersonalInformationFromSpreadSheet()
        {
            ExcelCollector collector = new ExcelCollector(_dataPath);
            ReportData data  = collector.FetchData();
            Assert.AreEqual("Mr",data.Client.Title);
            Assert.AreEqual("David",data.Client.FirstName);
            Assert.AreEqual("Piater",data.Client.LastName);
            Assert.AreEqual("25 Willoughby Street, Selcourt, 1559", data.Client.Address);
            Assert.AreEqual("084 580 5881",data.Client.CellNumber);
            Assert.AreEqual("011 818 5551",data.Client.LandLine);
            Assert.AreEqual(new DateTime(1959,8,8), data.Client.DOB);
            Assert.AreEqual("5908080096083", data.Client.IDNumber);
            Assert.AreEqual("Afrikaans",data.Client.HomeLanguage);
            Assert.AreEqual("English",data.Client.AssessmentLanguage);
            Assert.AreEqual(true,data.Client.DriverLicense);
        }

        [Test]
        public void ShouldFetchBackgroundInformation()
        {
            ExcelCollector collector = new ExcelCollector(@"testdata.xlsx");
            var b = collector.FetchData().Background;
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
            ExcelCollector collector = new ExcelCollector(_dataPath);
            var o = collector.FetchData().Occupation;
            Assert.AreEqual("Senior Personnel Practitioner – Police Officer", o.OccupationAtTimeOfIllness);
            Assert.AreEqual("SAPS", o.EmployerAtTimeOfIllness);
            Assert.AreEqual("Vispol Support Officer ", o.CurrentOccupation);
            Assert.AreEqual("SAPS", o.CurrentEmployer);
        }

        [Test]
        public void ShouldFetchJobDescriptionAndAddToOccupation()
        {
            ExcelCollector collector = new ExcelCollector(_dataPath);
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
    }
}
