using HtmlAgilityPack;
using ExcelDataReader;
using System.Data;
using System.Text;

class Program
{
    static async Task Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
        DataTable data = ReadExcelTable(filePath + "\\data.xlsx", "Data1");

        ExportToCsv(data, filePath);
    }


    #region Helper methods

    private static void ExportToCsv(DataTable table, string filePath)
    {
        // Create a StringBuilder to hold the CSV content
        StringBuilder stringBuilder = new StringBuilder();

        // Find the row that contains 'series id'
        int seriesIdRow = -1;
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (table.Rows[i][0].ToString() == "Series ID")
            {
                seriesIdRow = i;
                break;
            }
        }

        // If 'series id' was not found, return
        if (seriesIdRow == -1)
        {
            Console.WriteLine("'series id' not found in the first column.");
            return;
        }

        // Extract values from the first column starting from 'series id'
        var columnValues = table.AsEnumerable().Skip(seriesIdRow).Select(r =>
        {
            if (r[0] is DateTime)
            {
                return ((DateTime)r[0]).ToString("MMM-yyyy");
            }
            else
            {
                return r[0].ToString();
            }
        }).ToArray();

        // Append the values as a single row
        stringBuilder.AppendLine(string.Join(",", columnValues));

        // Export the entire row (excluding the first column)
        var rowValues = table.Rows[seriesIdRow].ItemArray.Skip(1).Select(val =>
        {
            return val.ToString();
        }).ToArray();

        stringBuilder.AppendLine(string.Join(",", rowValues));

        // Save to CSV file
        File.WriteAllText(filePath + "\\data.csv", stringBuilder.ToString());
    }


    private static DataTable ReadExcelTable(string filePath, string tableName)
    {
        DataTable table = new DataTable();

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252) // Use the desired encoding (e.g., 1252)
            }))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

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
        string fileName = "data.xlsx";
        
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