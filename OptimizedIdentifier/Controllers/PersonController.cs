using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OptimizedIdentifier.Services;
using System;
using System.Threading.Tasks;

namespace OptimizedIdentifier.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PersonController : ControllerBase
    {
        private readonly IPersonVerificationService _personVerificationService;

        public PersonController(IPersonVerificationService personVerificationService)
        {
            _personVerificationService = personVerificationService;
        }

        [HttpPost("verify-from-excel")]
        public async Task<IActionResult> VerifyFromExcel(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("Dosya seçilmedi.");
                }

                var result = await _personVerificationService.VerifyFromExcelAsync(file);
                return Content(result, "application/json");
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, $"Bir hata oluştu: {ex.Message}");
            }
        }
    }
}
