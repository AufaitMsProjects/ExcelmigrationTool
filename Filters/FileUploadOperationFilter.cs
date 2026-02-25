using Microsoft.AspNetCore.Http;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Linq;

namespace ExcelMigrationTool.Filters;

public class FileUploadOperationFilter : IOperationFilter
{
    public void Apply(OpenApiOperation operation, OperationFilterContext context)
    {
        var fileParameters = context.MethodInfo.GetParameters()
            .Where(p => p.ParameterType == typeof(IFormFile) || 
                       p.ParameterType == typeof(IFormFile[]))
            .ToList();

        if (fileParameters.Any())
        {
            operation.RequestBody = new OpenApiRequestBody
            {
                Content = new Dictionary<string, OpenApiMediaType>
                {
                    ["multipart/form-data"] = new OpenApiMediaType
                    {
                        Schema = new OpenApiSchema
                        {
                            Type = "object",
                            Properties = new Dictionary<string, OpenApiSchema>(),
                            Required = new HashSet<string>()
                        }
                    }
                }
            };

            foreach (var param in fileParameters)
            {
                var schema = new OpenApiSchema
                {
                    Type = "string",
                    Format = "binary"
                };

                operation.RequestBody.Content["multipart/form-data"].Schema.Properties.Add(param.Name, schema);
            }

            // Add other form parameters
            var formParams = context.MethodInfo.GetParameters()
                .Where(p => p.ParameterType != typeof(IFormFile) && 
                           p.ParameterType != typeof(IFormFile[]) &&
                           !p.GetCustomAttributes(typeof(System.Runtime.InteropServices.OptionalAttribute), false).Any())
                .ToList();

            foreach (var param in formParams)
            {
                var schema = new OpenApiSchema
                {
                    Type = GetSwaggerType(param.ParameterType),
                    Format = GetSwaggerFormat(param.ParameterType)
                };

                operation.RequestBody.Content["multipart/form-data"].Schema.Properties.Add(param.Name, schema);
                
                if (param.HasDefaultValue == false)
                {
                    operation.RequestBody.Content["multipart/form-data"].Schema.Required.Add(param.Name);
                }
            }
        }
    }

    private string GetSwaggerType(Type type)
    {
        if (type == typeof(string) || type == typeof(char))
            return "string";
        if (type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte))
            return "integer";
        if (type == typeof(float) || type == typeof(double) || type == typeof(decimal))
            return "number";
        if (type == typeof(bool))
            return "boolean";
        if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
            return "string";
        return "string";
    }

    private string? GetSwaggerFormat(Type type)
    {
        if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
            return "date-time";
        if (type == typeof(int))
            return "int32";
        if (type == typeof(long))
            return "int64";
        if (type == typeof(float))
            return "float";
        if (type == typeof(double) || type == typeof(decimal))
            return "double";
        return null;
    }
}

