using ExcelDataReader;
using System.Data;

namespace ExcelMigrationTool.Helpers;

public static class ExcelReaderHelper
{
    public static DataTable ReadExcelToDataTable(Stream excelStream, string fileName)
    {
        using var reader = ExcelReaderFactory.CreateReader(excelStream);
        
        var result = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = true
            }
        });

        if (result.Tables.Count == 0)
        {
            throw new InvalidOperationException("Excel file contains no data tables.");
        }

        return result.Tables[0];
    }
}

