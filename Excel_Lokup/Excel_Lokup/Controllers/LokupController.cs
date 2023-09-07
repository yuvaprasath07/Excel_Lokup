using DataLayer;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Excel_Lokup.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LokupController : ControllerBase
    {
        public readonly IDatalayer datalayer;
        public LokupController(IDatalayer datalayer)
        {
            this.datalayer = datalayer;
            
        }
        [HttpGet("BatchEnrollment")]
        public IActionResult Excellooup()
        {
            var data = datalayer.GetExcelLokup();
            return Ok();
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            var fileName = datalayer.UploadFileAsync(file);
            if (fileName is List<string> differences)
            {
                return BadRequest(new { Message = "Column name differences detected", Differences = differences });
            }
            else if (fileName != null)
            {
                return Ok(fileName);
            }
            return StatusCode(400, "File Not Match");
        }
    }
}
