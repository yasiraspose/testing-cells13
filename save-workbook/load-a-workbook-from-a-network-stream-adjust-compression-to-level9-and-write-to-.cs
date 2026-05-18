using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    static async Task Main()
    {
        // URL of the Excel file to load
        string url = "https://example.com/sample.xlsx";

        using var client = new HttpClient();

        try
        {
            using var response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();

            await using var networkStream = await response.Content.ReadAsStreamAsync();

            // Load the workbook from the network stream
            var workbook = new Workbook(networkStream);

            // Configure OOXML save options with maximum compression (Level9)
            var saveOptions = new OoxmlSaveOptions
            {
                CompressionType = OoxmlCompressionType.Level9
            };

            // Save the workbook to a local file using the specified compression
            workbook.Save("DownloadedCompressed.xlsx", saveOptions);
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Failed to download the file: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}