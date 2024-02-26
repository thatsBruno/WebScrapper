// See https://aka.ms/new-console-template for more information
using HtmlAgilityPack;
using System.Net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

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
                if(fileLink.Contains("01.xls"))
                DownloadFile("https://www.abs.gov.au/" + fileLink, "C:\\Users\\Princcipal\\Desktop\\");

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

    static void DownloadFile(string url, string directory)
    {
        string fileName = Path.GetFileName(url);
        string destinationPath = Path.Combine(directory, fileName);

        using (WebClient client = new WebClient())
        {
            try
            {
                client.DownloadFile(url, destinationPath);
                Console.WriteLine($"File downloaded: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to download {fileName}: {ex.Message}");
            }
        }
    }
}
