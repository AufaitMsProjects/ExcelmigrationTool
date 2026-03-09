using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using System.Data;
using System.Text;

namespace ExcelMigrationTool.Helpers;

public static class ExcelReaderHelper
{
    public static DataTable ReadFileToDataTable(Stream fileStream, string fileName)
    {
        var fileExtension = Path.GetExtension(fileName).ToLowerInvariant();

        return fileExtension switch
        {
            ".xlsx" or ".xls" => ReadExcelToDataTable(fileStream),
            ".csv" => ReadCsvToDataTable(fileStream),
            _ => throw new InvalidOperationException($"Unsupported file type: {fileExtension}")
        };
    }

    private static DataTable ReadExcelToDataTable(Stream excelStream)
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

    private static DataTable ReadCsvToDataTable(Stream csvStream)
    {
        var dataTable = new DataTable();

        using var parser = new TextFieldParser(csvStream, Encoding.UTF8)
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true
        };
        parser.SetDelimiters(",");

        if (parser.EndOfData)
        {
            throw new InvalidOperationException("CSV file contains no data.");
        }

        var headers = parser.ReadFields();
        if (headers == null || headers.Length == 0)
        {
            throw new InvalidOperationException("CSV file header row is missing.");
        }

        var duplicateCounter = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var rawHeader in headers)
        {
            var baseHeader = string.IsNullOrWhiteSpace(rawHeader) ? "Column" : rawHeader.Trim();
            var headerName = baseHeader;

            if (duplicateCounter.TryGetValue(baseHeader, out var existingCount))
            {
                existingCount++;
                duplicateCounter[baseHeader] = existingCount;
                headerName = $"{baseHeader}_{existingCount}";
            }
            else
            {
                duplicateCounter[baseHeader] = 1;
            }

            dataTable.Columns.Add(headerName);
        }

        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields();
            if (fields == null)
            {
                continue;
            }

            var row = dataTable.NewRow();
            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                row[i] = i < fields.Length ? fields[i] : DBNull.Value;
            }

            dataTable.Rows.Add(row);
        }

        return dataTable;
    }
}

