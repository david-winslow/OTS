using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using ReportDataCollector;

namespace ReportGenerator.Tests
{
    [TestFixture]
    public class ExportGoogleFileTests
    {
        private GoogleFileService _gs;
        private string _localPath;

        [SetUp]
        public void Setup()
        {
            _gs = new GoogleFileService();

            _localPath = "c:\\googledrive";
            if (!Directory.Exists(_localPath))
            {
                Directory.CreateDirectory(_localPath);
            }
        }

        [Test]
        public void ShouldReturnFalseIfNoFileToProcess()
        {            
            Assert.IsFalse(_gs.HasFileToProcess());
        }

        [Test]
        public void ShouldReturnTrueIfHasFileToProcess()
        {
            _gs.UploadFile("input_final.gsheet");
            Assert.IsTrue(_gs.HasFileToProcess());
            _gs.DeleteFile(_gs.InputSheetId);
        }

        [Test]
        public void ShouldSetInputSheetIdIfHasFileToProcess()
        {
            _gs.UploadFile("input_final.gsheet");
            _gs.HasFileToProcess();
            Assert.NotNull(_gs.InputSheetId);
            _gs.DeleteFile(_gs.InputSheetId);
        }

        [Test]
        public void ShouldResetInputSheetIdIfNoFileToProcess()
        {
            _gs.UploadFile("input_final.gsheet");
            _gs.HasFileToProcess();
            _gs.DeleteFile(_gs.InputSheetId);
            _gs.HasFileToProcess();
            Assert.IsNull(_gs.InputSheetId);
        }

        [Ignore,Test]
        public void ShouldHaveDownloadedFileAfterProcessing()
        {
            _gs.UploadFile("input_final.gsheet");
            _gs.HasFileToProcess();
            _gs.ProcessFile();
            Assert.IsTrue(File.Exists(Path.Combine(_localPath,"Input.xlsx")));
            _gs.DeleteFile(_gs.InputSheetId);
        }
    }
}
