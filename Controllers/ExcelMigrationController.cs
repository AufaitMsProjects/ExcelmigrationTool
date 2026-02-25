using ExcelMigrationTool.Helpers;
using ExcelMigrationTool.Models;
using ExcelMigrationTool.Services;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace ExcelMigrationTool.Controllers;

[ApiController]
[Route("api/excel")]
[Produces("application/json")]
public class ExcelMigrationController : ControllerBase
{
    private readonly IExcelMigrationService _migrationService;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ExcelMigrationController> _logger;

    public ExcelMigrationController(
        IExcelMigrationService migrationService,
        IConfiguration configuration,
        ILogger<ExcelMigrationController> logger)
    {
        _migrationService = migrationService;
        _configuration = configuration;
        _logger = logger;
    }

    /// <summary>
    /// Uploads an Excel file and migrates its data to a SQL Server table
    /// </summary>
    /// <param name="targetSchema">SQL Server schema name (e.g., master, dbo)</param>
    /// <param name="targetTable">SQL Server table name (without schema)</param>
    /// <param name="excelFile">Excel file (.xlsx) to upload</param>
    /// <param name="userId">User ID for tracking</param>
    /// <param name="attachmentRecordType">Attachment record type for BPAttachments table (e.g., Comment, OrderTransmittal)</param>
    /// <returns>Migration result with success/failure counts and error messages</returns>
    [HttpPost("upload")]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(typeof(UploadResponse), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(UploadResponse), StatusCodes.Status400BadRequest)]
    public async Task<ActionResult<UploadResponse>> UploadExcel([FromForm] string targetSchema,[FromForm] string targetTable, [FromForm] IFormFile excelFile,[FromForm] long userId, [FromForm] string? attachmentRecordType = null, CancellationToken cancellationToken = default)
    {
        var response = new UploadResponse();

        try
        {
            _logger.LogInformation("Excel upload started for table {TableName} by user {UserId}, file size: {FileSize} bytes", 
                targetTable, userId, excelFile?.Length ?? 0);
            // Validation
            if (string.IsNullOrWhiteSpace(targetSchema))
            {
                response.ErrorMessages.Add("targetSchema is required.");
                return BadRequest(response);
            }

            if (string.IsNullOrWhiteSpace(targetTable))
            {
                response.ErrorMessages.Add("targetTable is required.");
                return BadRequest(response);
            }

            if (excelFile == null || excelFile.Length == 0)
            {
                response.ErrorMessages.Add("Excel file is required.");
                return BadRequest(response);
            }

            var allowedExtensions = new[] { ".xlsx", ".xls" };
            var fileExtension = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            
            if (!allowedExtensions.Contains(fileExtension))
            {
                response.ErrorMessages.Add($"Invalid file type. Only {string.Join(", ", allowedExtensions)} files are allowed.");
                return BadRequest(response);
            }

            // Get connection string
            var connectionString = _configuration.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                response.ErrorMessages.Add("Database connection string is not configured.");
                return BadRequest(response);
            }

            // Read Excel file
            DataTable excelData;
            try
            {
                _logger.LogInformation("Reading Excel file: {FileName}", excelFile.FileName);
                await using var stream = excelFile.OpenReadStream();
                excelData = ExcelReaderHelper.ReadExcelToDataTable(stream, excelFile.FileName);
                _logger.LogInformation("Excel file read successfully. Rows: {RowCount}, Columns: {ColumnCount}", 
                    excelData.Rows.Count, excelData.Columns.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading Excel file: {FileName}", excelFile.FileName);
                response.ErrorMessages.Add($"Error reading Excel file: {ex.Message}");
                return BadRequest(response);
            }

            // Perform migration
            _logger.LogInformation("Starting migration for {Schema}.{Table} with AttachmentRecordType: {AttachmentRecordType}", 
                targetSchema, targetTable, attachmentRecordType ?? "None");
            response = await _migrationService.MigrateExcelToSqlServerAsync(
                connectionString,
                targetSchema,
                targetTable,
                excelData,
                attachmentRecordType,
                cancellationToken);
            _logger.LogInformation("Migration completed. Inserted: {Inserted}, Updated: {Updated}, Failed: {Failed}", 
                response.RowsInserted, response.RowsUpdated, response.RowsFailed);

            if (response.Success)
            {
                return Ok(response);
            }
            else
            {
                return BadRequest(response);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing Excel upload for user {UserId}", userId);
            response.ErrorMessages.Add($"Unexpected error: {ex.Message}");
            return StatusCode(500, response);
        }
    }
}

