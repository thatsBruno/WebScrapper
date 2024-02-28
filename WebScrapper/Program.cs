using HtmlAgilityPack;
using ExcelDataReader;
using System.Data;

class Program
{
    static async Task Main(string[] args)
    {
        List<string> initialLinks = ScrapeLinks("https://www.abs.gov.au/statistics/labour/employment-and-unemployment/labour-force-australia");

        string targetLink = initialLinks.Count > 0 ? initialLinks[0] : null;
        string dirPath = "Downloads";
        string defaultPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

        var filePath = Path.Combine(defaultPath, dirPath);

        if (!Directory.Exists(filePath))
        {
            Directory.CreateDirectory(filePath);
        }

        if (targetLink != null)
        {
            List<string> fileLinks = ScrapeLinks("https://www.abs.gov.au/" + targetLink);

            foreach (string fileLink in fileLinks)
            {
                if (fileLink.Contains("01.xls"))
                {
                    await DownloadFileAsync("https://www.abs.gov.au/" + fileLink, filePath); 
                }
            } 
        }
        else
        {
            Console.WriteLine("No links found on the initial page.");
        }

        // read and save the table called Data1 into a .csv from the file
        DataTable data = ReadExcelTable(filePath, "Data1");

        // test 
        foreach (DataRow row in data.Rows)
        {
            foreach (DataColumn col in data.Columns)
            {
                Console.Write(row[col].ToString() + "\t");
            }
            Console.WriteLine();
        }
    }


    #region Helper methods

    private static DataTable ReadExcelTable(string filePath, string tableName)
    {
        // Create a DataTable to hold the data
        DataTable table = new DataTable();

        // Create a FileStream to read the Excel file
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            // Create an instance of ExcelDataReader for .xlsx files
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                // Read the Excel file into a DataSet
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true // Use the first row as column names
                    }
                });

                // Get the specified table from the DataSet
                table = result.Tables[tableName];
            }
        }

        return table;
    }

    private static List<string> ScrapeLinks(string url)
    {
        List<string> links = new List<string>();

        HtmlWeb web = new HtmlWeb();
        HtmlDocument doc = web.Load(url);

        foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a[@href]"))
        {
            string href = link.GetAttributeValue("href", null);
            if (!string.IsNullOrEmpty(href) && href.Contains("2024"))
            {
                links.Add(href);
            }
        }

        return links;
    }

    private static async Task DownloadFileAsync(string url, string directory)
    {
        string fileName = Path.GetFileName(url);
        
        using (HttpClient httpClient = new HttpClient())
        {
            try
            {
                // Set user-agent header to mimic a web browser
                httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");

                // Set referer header if necessary
                httpClient.DefaultRequestHeaders.Referrer = new Uri(url);

                HttpResponseMessage response = await httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();

                    string filePath = Path.Combine(directory, fileName); // TODO store the file inside the project
                    System.IO.File.WriteAllBytes(filePath, fileBytes);
                    Console.WriteLine($"File downloaded successfully to: {filePath}");
                }
                else
                {
                    Console.WriteLine($"Failed to download the file. Status code: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
    #endregion
}