using System;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using GoogleSheets.Models;
using GoogleSheets.Services;
using Microsoft.AspNetCore.Mvc;

namespace GoogleSheets.Controllers
{

    [Route("api/sheets")]
    [ApiController]
    public class GSheetsController : Controller
    {
        private readonly ISheets _sheets;

        public GSheetsController(ISheets sheets)
        {
            _sheets = sheets;
        }

        [HttpGet("{number}")]
        public async Task<ActionResult> GetAsync(string number)
        {
            var data = _sheets.ReturnName(number);

            if(data == null) return BadRequest("Not found");

            return Ok(data);

        }
    }
}
