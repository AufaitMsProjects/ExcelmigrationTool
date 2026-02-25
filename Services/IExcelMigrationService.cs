using ExcelMigrationTool.Models;
using System.Data;

namespace ExcelMigrationTool.Services;

public interface IExcelMigrationService
{
    Task<UploadResponse> MigrateExcelToSqlServerAsync(
        string connectionString,
        string schemaName,
        string tableName,
        DataTable excelData,
        string? attachmentRecordType = null,
        CancellationToken cancellationToken = default);
}

