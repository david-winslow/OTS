﻿using System;
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

namespace ReportDataCollector
{
    public class GoogleFileService
    {
        private string clientId = "1057012666404-8lbnbd8cf45dr658c7bqa7jrufr66teo.apps.googleusercontent.com";
        private string secret = "9Mkk3rCzOI2tYLsmmZqfzFec";
        /// <summary>The Drive API scopes.</summary>
        private static readonly string[] Scopes = new[] { DriveService.Scope.DriveFile, DriveService.Scope.Drive };

        public async void DownloadFile(string from , string to)
        {
            GoogleWebAuthorizationBroker.Folder = "Drive.Sample";
            UserCredential credential;
            using (var stream = new System.IO.FileStream("google.json",
                System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None);
            }
            // Create the service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "fcereporting",
            });

            var listRequest = service.Files.List().Execute();
            
        }
    }
}