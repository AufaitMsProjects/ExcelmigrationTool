namespace ExcelMigrationTool.Models;

public class RowErrorDetail
{
    public int RowNumber { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public object? Value { get; set; }
    public string ErrorMessage { get; set; } = string.Empty;
    public Dictionary<string, object?> RowData { get; set; } = new();
}

