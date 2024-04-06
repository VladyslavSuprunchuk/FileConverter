using FileConverter.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace FileConverter.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        private readonly IReportService _reportService;
        private readonly ICsvService _csvService;

        public ReportController(IReportService reportService, ICsvService csvService)
        {
            _reportService = reportService;
            _csvService = csvService;
        }

        [HttpGet("WorldwideReport")]
        public async Task<IResult> Get()
        {
            var content = await _reportService.GetWorldwideReportAsync();

            if (string.IsNullOrEmpty(content))
            {
                return Results.BadRequest("Invalid report");
            }

            var stream = _csvService.CreateCsvReport(content);

            return Results.File(stream, "text/csv", "report.csv");
        }
    }
}
