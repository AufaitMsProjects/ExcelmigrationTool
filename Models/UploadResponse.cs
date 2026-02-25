namespace ExcelMigrationTool.Models;

public class UploadResponse
{
    public bool Success { get; set; }
    public int RowsInserted { get; set; }
    public int RowsUpdated { get; set; }
    public int RowsFailed { get; set; }
    public List<string> ErrorMessages { get; set; } = new();
    public List<RowErrorDetail> RowErrors { get; set; } = new();
    public string Message { get; set; } = string.Empty;
}

