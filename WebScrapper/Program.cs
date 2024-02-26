// See https://aka.ms/new-console-template for more information
using HtmlAgilityPack;

class Program
{
    static async Task Main(string[] args)
    {
        // Step 1: Scrape the initial page for links
        List<string> initialLinks = ScrapeLinks("https://www.abs.gov.au/statistics/labour/employment-and-unemployment/labour-force-australia");

        // Step 2: Choose a link to navigate to
        string targetLink = initialLinks.Count > 0 ? initialLinks[0] : null;

        if (targetLink != null)
        {
            // Step 3: Scrape the target page for file download links
            List<string> fileLinks = ScrapeLinks("https://www.abs.gov.au/" + targetLink);

            // Step 4: Download the files
            foreach (string fileLink in fileLinks)
            {
                if (fileLink.Contains("01.xls"))
                {
                    await DownloadFileAsync("https://www.abs.gov.au/" + fileLink, "C:\\Users\\Princcipal\\Desktop\\");
                }
            }
        }
        else
        {
            Console.WriteLine("No links found on the initial page.");
        }
    }

    static List<string> ScrapeLinks(string url)
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

    static async Task DownloadFileAsync(string url, string directory)
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

                    string filePath = Path.Combine(directory, fileName);
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
}
