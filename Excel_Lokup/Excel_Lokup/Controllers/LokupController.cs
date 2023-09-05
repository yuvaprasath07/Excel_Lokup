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
    }
}
