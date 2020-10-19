using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using GoogleSheets.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GoogleSheets.Services
{
    public class GSheetsService : ISheets
    {
        private readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private SheetsService _service;

        public List<string> ReturnName(string number)
        {
            GoogleCredential credential;

            using (var stream = new FileStream(Constants.Credentials, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(_scopes);
            }

            _service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = Constants.ApplicationName,
            });

            var request = _service.Spreadsheets.Values.Get(Constants.SpreadsheetId, Constants.Range);
            var response = request.Execute();
            var values = response.Values;

            var dados = values.Where(v => v.Any())
                                     .Select(v => v[0].ToString())
                                     .ToList();

            return dados;

        }
    }
}
