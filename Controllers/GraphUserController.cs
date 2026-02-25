using ExcelMigrationTool.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System;
using System.IO;
using System.Threading.Tasks;

namespace ExcelMigrationTool.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class GraphUserController : ControllerBase
    {
        private readonly IAzureGraphService _graphService;
        private readonly IWebHostEnvironment _env;

        public GraphUserController(IAzureGraphService graphService, IWebHostEnvironment env)
        {
            _graphService = graphService;
            _env = env;
        }

        [HttpPost("process-users")]
        public async Task<IActionResult> ProcessUsers([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            try
            {
                // Create temp directory if it doesn't exist
                var tempPath = Path.Combine(_env.ContentRootPath, "Temp");
                if (!Directory.Exists(tempPath))
                {
                    Directory.CreateDirectory(tempPath);
                }

                var inputFileName = "users.xlsx";
                var outputFileName = "users_updated.xlsx";
                var inputFilePath = Path.Combine(tempPath, inputFileName);
                var outputFilePath = Path.Combine(tempPath, outputFileName);

                // Save uploaded file to temp path
                using (var stream = new FileStream(inputFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                // Process the file
                await _graphService.ProcessUserExcelAsync(inputFilePath, outputFilePath);

                // Read the updated file for return
                var fileBytes = await System.IO.File.ReadAllBytesAsync(outputFilePath);
                
                // Cleanup
                if (System.IO.File.Exists(inputFilePath)) System.IO.File.Delete(inputFilePath);
                // We keep output for a moment to return it, then it can be cleaned or left in temp

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputFileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet("process-local")]
        public async Task<IActionResult> ProcessLocalFile()
        {
            try
            {
                var rootPath = Path.Combine(_env.ContentRootPath, "Root");
                var inputFilePath = Path.Combine(rootPath, "users.xlsx");
                var outputFilePath = Path.Combine(rootPath, "users_updated.xlsx");

                if (!System.IO.File.Exists(inputFilePath))
                {
                    return NotFound($"The file 'users.xlsx' was not found in the 'Root' directory ({rootPath}).");
                }

                await _graphService.ProcessUserExcelAsync(inputFilePath, outputFilePath);

                return Ok($"Successfully processed users. Updated file saved as 'users_updated.xlsx' in the 'Root' directory.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
