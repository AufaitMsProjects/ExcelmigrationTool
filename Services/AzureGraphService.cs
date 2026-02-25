using Azure.Identity;
using ClosedXML.Excel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using System.Collections.Generic;
using System.Threading;
using Microsoft.Extensions.Logging;

namespace ExcelMigrationTool.Services
{
    public class AzureGraphService : IAzureGraphService
    {
        private readonly string _tenantId = "5416f9df-a5ff-40b7-be18-c480263c20ef";
        private readonly string _clientId = "ed391a60-bffa-4e17-8c40-b199c207276c";
        private readonly string _clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET")
            ?? throw new InvalidOperationException("Missing required environment variable: AZURE_CLIENT_SECRET");
        private readonly ILogger<AzureGraphService> _logger;

        public AzureGraphService(ILogger<AzureGraphService> logger)
        {
            _logger = logger;
        }

        public async Task ProcessUserExcelAsync(string inputFilePath, string outputFilePath)
        {
            if (!File.Exists(inputFilePath))
            {
                throw new FileNotFoundException("The input Excel file 'users.xlsx' was not found.", inputFilePath);
            }

            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            var clientSecretCredential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret, options);
            var graphClient = new GraphServiceClient(clientSecretCredential);

            using (var workbook = new XLWorkbook(inputFilePath))
            {
                var worksheet = workbook.Worksheets.First();
                
                // Find column indices
                var headerRow = worksheet.Row(1);
                int emailColIndex = 1; // Default to column 1
                int azureIdColIndex = 2; // Default to column 2

                bool foundEmailHeader = false;
                for (int i = 1; i <= headerRow.LastCellUsed().Address.ColumnNumber; i++)
                {
                    string headerValue = headerRow.Cell(i).GetValue<string>().Trim();
                    if (string.Equals(headerValue, "Email", StringComparison.OrdinalIgnoreCase))
                    {
                        emailColIndex = i;
                        foundEmailHeader = true;
                    }
                    else if (string.Equals(headerValue, "Azure ID", StringComparison.OrdinalIgnoreCase) || 
                             string.Equals(headerValue, "Object ID", StringComparison.OrdinalIgnoreCase))
                    {
                        azureIdColIndex = i;
                    }
                }

                if (!foundEmailHeader)
                {
                    _logger.LogWarning("No column named 'Email' found in the header row. Defaulting to column 1.");
                }
                else
                {
                    _logger.LogInformation($"Using column {emailColIndex} ('Email') for lookups.");
                }

                // If Azure ID column doesn't exist, create it in the next available column
                if (worksheet.Cell(1, azureIdColIndex).IsEmpty() && !foundEmailHeader)
                {
                     // If we didn't find specific headers, we stick to column 2 as per original req
                }
                else if (headerRow.Cell(azureIdColIndex).IsEmpty())
                {
                    headerRow.Cell(azureIdColIndex).Value = "Azure ID";
                }

                var rows = worksheet.RowsUsed().Skip(1).ToList(); // Skip header and materialize

                _logger.LogInformation($"Starting parallel processing for {rows.Count} rows.");

                // Use a semaphore to limit concurrency and avoid throttling or overwhelming the system
                var semaphore = new SemaphoreSlim(15); 
                var tasks = new List<Task<(int RowNumber, string ObjectId)>>();

                foreach (var row in rows)
                {
                    string email = row.Cell(emailColIndex).GetValue<string>().Trim();
                    int rowNum = row.RowNumber();

                    if (string.IsNullOrEmpty(email)) continue;

                    tasks.Add(Task.Run(async () =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            return (rowNum, await GetUserObjectId(graphClient, email));
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    }));
                }

                // Wait for all lookups to complete
                var results = await Task.WhenAll(tasks);

                // Update the worksheet sequentially (ClosedXML is not thread-safe)
                foreach (var result in results)
                {
                    var row = worksheet.Row(result.RowNumber);
                    row.Cell(azureIdColIndex).Value = result.ObjectId;
                }

                _logger.LogInformation("Finished processing all rows. Saving workbook.");
                workbook.SaveAs(outputFilePath);
            }
        }

        private async Task<string> GetUserObjectId(GraphServiceClient graphClient, string email)
        {
            try
            {
                // Search for the user by email, UPN, or Aliases (proxyAddresses)
                var users = await graphClient.Users.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = $"mail eq '{email}' or userPrincipalName eq '{email}' or proxyAddresses/any(c:c eq 'smtp:{email}')";
                    requestConfiguration.QueryParameters.Select = new[] { "id" };
                    requestConfiguration.QueryParameters.Top = 1;
                    requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                });

                var user = users?.Value?.FirstOrDefault();

                if (user != null && !string.IsNullOrEmpty(user.Id))
                {
                    return user.Id;
                }

                // Fallback: try getting directly by ID/UPN
                try
                {
                    var directUser = await graphClient.Users[email].GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select = new[] { "id" };
                    });

                    if (directUser != null && !string.IsNullOrEmpty(directUser.Id))
                    {
                        return directUser.Id;
                    }
                }
                catch (ODataError odataErr) when (odataErr.Error?.Code == "Authorization_RequestDenied" || odataErr.ResponseStatusCode == 403)
                {
                    _logger.LogError("Insufficient permissions to fetch user data. Please grant 'User.Read.All' (Application) permission.");
                    return "Error: Insufficient Permissions (403)";
                }
                catch
                {
                    // Ignore other fallback failures
                }

                _logger.LogWarning($"User not found: {email}");
                return "Not Found";
            }
            catch (ODataError odataErr) when (odataErr.Error?.Code == "Authorization_RequestDenied" || odataErr.ResponseStatusCode == 403)
            {
                _logger.LogError("Insufficient permissions to fetch user data. Please grant 'User.Read.All' (Application) permission.");
                return "Error: Insufficient Permissions (403)";
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error fetching user {email}: {ex.Message}");
                return "Error: Fetch Failed";
            }
        }
    }
}
