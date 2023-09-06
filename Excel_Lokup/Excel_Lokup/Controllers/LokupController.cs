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
            try
            {
                var fileName = datalayer.UploadFileAsync(file);
                return Ok($"File uploaded and saved as {fileName}");
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}
