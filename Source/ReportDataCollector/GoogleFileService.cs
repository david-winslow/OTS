using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Http;
using Google.Apis.Services;
using OTS.Domain;

namespace ReportDataCollector
{
    public class GoogleFileService
    {
        private DriveService _service;
        private static string _accessToken;
        private const string CLIENTID = "1057012666404-8lbnbd8cf45dr658c7bqa7jrufr66teo.apps.googleusercontent.com";
        private const string SECRET = "9Mkk3rCzOI2tYLsmmZqfzFec";
        private const string GDRIVEROOT = "root";
        private const string LOCALPATH = "c:\\googledrive";
        private string _inputSheetId;
        public string InputSheetId { get { return _inputSheetId; } }

        public GoogleFileService()
        {
            _service = GetAuthenticatedDriveService();
        }

        public void DownloadFile(string fileId, string target, string fileFormat)
        {
            using (var gDriveStream = GetGoogleDriveFileStream(fileId, fileFormat))
            {
                using (FileStream fs = new FileStream(target, FileMode.Create, FileAccess.ReadWrite))
                {
                    gDriveStream.CopyTo(fs);
                }
            }
        }

        private static DriveService GetAuthenticatedDriveService()
        {
            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = CLIENTID,
                    ClientSecret = SECRET,
                },
                new[] {DriveService.Scope.Drive, DriveService.Scope.DriveFile},
                "user",
                CancellationToken.None).Result;
            _accessToken = credential.Token.AccessToken;

            // Create the service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "fcereporting"
            });
            return service;
        }

        /// <summary>
        /// Download a file and return a string with its content.
        /// </summary>
        /// <param name="authenticator">
        /// Authenticator responsible for creating authorized web requests.
        /// </param>
        /// <param name="file">Drive File instance.</param>
        /// <returns>File's content if successful, null otherwise.</returns>
        public Stream GetGoogleDriveFileStream(string id, string exportFormat)
        {
            var file = _service.Files.Get(id).Execute();
            var url = file.DownloadUrl;
            if(exportFormat != null)
                url = file.ExportLinks[exportFormat];
            if (!String.IsNullOrEmpty(url))
            {
                
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(
                        new Uri(url));
                    _service.HttpClientInitializer.Initialize(_service.HttpClient);
                    
                    request.Headers.Add(HttpRequestHeader.Authorization, string.Format("Bearer {0}", _accessToken));
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        return response.GetResponseStream();
                    }
                    else
                    {
                        Console.WriteLine("An error occurred: " + response.StatusDescription);
                        return null;
                    }
            }
            else
            {
                // The file doesn't have any content stored on Drive.
                return null;
            }
        }

        public bool HasFileToProcess()
        {
            var filesRequest = _service.Files.List();
            filesRequest.Q = "mimeType!='application/vnd.google-apps.folder' and trashed=false";
            var files = filesRequest.Execute();
            foreach (var file in files.Items)
            {
                if (file.Title.EndsWith("_final"))
                {
                    _inputSheetId = file.Id;
                    return true;
                }
            }
            _inputSheetId = null;
            return false;
        }

        public void UploadFile(string filePath)
        {
            using (var fs = new FileStream("input_final.gsheet", FileMode.Open, FileAccess.Read))
            {
                fs.Seek(0, SeekOrigin.Begin);

                var gfile = new Google.Apis.Drive.v2.Data.File()
                {
                    Title = "input_final",
                    MimeType = "application/vnd.google-apps.spreadsheet"
                };
                var progress = _service.Files.Insert(gfile, fs, "text/plain").Upload();
            }
        }

        public void DeleteFile(string fileId)
        {
            _service.Files.Delete(fileId).Execute();
        }

        public void ProcessFile()
        {
            DownloadFile(InputSheetId, Path.Combine(LOCALPATH, "input.odt"), "application/vnd.oasis.opendocument.text");
        }

        public void CreateSessionStructure(ReportData data)
        {
            throw new NotImplementedException();
        }
    }
}
