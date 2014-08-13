using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v2;
using Google.Apis.Services;
using OTS.Domain;

namespace ReportDataCollector
{
    public class GoogleFileService
    {
        private string clientId = "1057012666404-8lbnbd8cf45dr658c7bqa7jrufr66teo.apps.googleusercontent.com";
        private string secret = "9Mkk3rCzOI2tYLsmmZqfzFec";
        /// <summary>The Drive API scopes.</summary>
        private static readonly string[] Scopes = new[] { DriveService.Scope.DriveFile, DriveService.Scope.Drive };


        public string DownloadFile(string from , string to)
        {
            return string.Empty;

        }

        public void CreateSessionStructure(ReportData data)
        {
            throw new NotImplementedException();
        }

        public string ProcessFile()
        {
            return null;

        }

        public bool HasFileToProcess()
        {
            //search for it
            //save url
            return false;
        }
    }
}
