using ExcelMigrationTool.Filters;
using ExcelMigrationTool.Services;
using System.Text;

var builder = WebApplication.CreateBuilder(args);

// Configure Kestrel server options for longer timeouts (30 minutes for large Excel processing)
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.KeepAliveTimeout = TimeSpan.FromMinutes(30);
    options.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(30);
    options.Limits.MaxRequestBodySize = 209715200; // 100 MB
});

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title = "Excel Migration API",
        Version = "v1",
        Description = "API for migrating Excel data to SQL Server tables"
    });
    
    // Add file upload support
    c.OperationFilter<FileUploadOperationFilter>();
});

// Register services
builder.Services.AddScoped<IExcelMigrationService, ExcelMigrationService>();
builder.Services.AddScoped<IAzureGraphService, AzureGraphService>();

// Configure form options for file uploads
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 209715200; // 100 MB
    options.ValueLengthLimit = int.MaxValue;
    options.ValueCountLimit = int.MaxValue;
    options.KeyLengthLimit = int.MaxValue;
    options.MultipartHeadersLengthLimit = int.MaxValue;
    options.MultipartBoundaryLengthLimit = int.MaxValue;
});

var app = builder.Build();

// Configure the HTTP request pipeline
// Only enable Swagger in Development environment for security
//if (app.Environment.IsDevelopment())
//{
app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Migration API v1");
        c.RoutePrefix = string.Empty; // Set Swagger UI at the app's root
    });
//}

// Only redirect to HTTPS if not in Development or if HTTPS is available
if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}
app.UseAuthorization();
app.MapControllers();

// Register code pages for ExcelDataReader
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

app.Run();

