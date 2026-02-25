using System.Threading.Tasks;

namespace ExcelMigrationTool.Services
{
    public interface IAzureGraphService
    {
        /// <summary>
        /// Reads users from an Excel file, fetches their Azure AD Object IDs via Microsoft Graph,
        /// and updates the Excel file with the IDs.
        /// </summary>
        /// <param name="inputFilePath">The path to the input Excel file (users.xlsx).</param>
        /// <param name="outputFilePath">The path where the updated Excel file (users_updated.xlsx) will be saved.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        Task ProcessUserExcelAsync(string inputFilePath, string outputFilePath);
    }
}
