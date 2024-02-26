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
        string destinationPath = Path.Combine(directory, fileName);

        using (HttpClient httpClient = new HttpClient())
        {
            try
            {
                // Send a GET request to the URL
                HttpResponseMessage response = await httpClient.GetAsync(url);

                // Check if the request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Get the content as a stream
                    using (Stream contentStream = await response.Content.ReadAsStreamAsync())
                    {
                        // Create a FileStream to write the content to a file
                        using (FileStream fileStream = File.Create("downloaded-file.txt"))
                        {
                            // Copy the content from the stream to the file
                            await contentStream.CopyToAsync(fileStream);
                        }
                    }

                    Console.WriteLine("File downloaded successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to download file. Status code: {response.StatusCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
