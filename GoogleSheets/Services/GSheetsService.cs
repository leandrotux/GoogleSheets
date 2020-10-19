using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using GoogleSheets.Models;
using System;
using System.IO;

namespace GoogleSheets.Services
{
    public class GSheetsService : ISheets
    {
        private readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private SheetsService _service;

        public string GetName(string number)
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
            try
            {
                var request = _service.Spreadsheets.Values.Get(Constants.SpreadsheetId, Constants.Range);
                var response = request.Execute();
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        if (row[3].ToString().Equals(number))
                        {
                            return row[0].ToString();
                        }
                    }
                }
                return "Not found";
            }
            catch (Exception ex)
            {
                return "Internal error" + ex.Message.ToString();
            }
        }
    }
}
