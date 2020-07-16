using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Mvc;

namespace GoogleSheets.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class GSheetsController : Controller
    {
        private readonly string[] _scopes = { SheetsService.Scope.Spreadsheets };
        private readonly string _applicationName = "teste2";
        private readonly string _spreadsheetId = "1LCn7Doxyc1P1esyaVKGH4DxoBfnwTG2Gq-FViHxGz_s";
        private readonly string _sheet = "teste";

        private SheetsService _service;

        [HttpGet("{number}")]
        public async Task<ActionResult> GetAsync(string number)
        {
            GoogleCredential credential;

            using (var stream = new FileStream("Credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(_scopes);
            }

            _service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName,
            });
            try
            {
                var range = $"{_sheet}!A2:D10";
                var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
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
                return BadRequest("Dados não encontrado.");
            }
            catch (Exception ex)
            {
                return BadRequest("Internal error" + ex.Message.ToString());
            }
        }
    }
}
