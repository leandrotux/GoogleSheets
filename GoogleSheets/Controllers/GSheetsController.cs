using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using GoogleSheets.Models;
using Microsoft.AspNetCore.Mvc;

namespace GoogleSheets.Controllers
{

    [Route("api/sheets")]
    [ApiController]
    public class GSheetsController : Controller
    {
        private readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private SheetsService _service;

        [HttpGet("{number}")]
        public async Task<ActionResult> GetAsync(string number)
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
                            return Ok(row[0].ToString());
                        }
                    }
                }
                return BadRequest("Not found");
            }
            catch (Exception ex)
            {
                return BadRequest("Internal error" + ex.Message.ToString());
            }
        }
    }
}
