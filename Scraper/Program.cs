using OfficeOpenXml;
using System.Text.RegularExpressions;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Scraper
{
    public class Progam
    {
        private static void Main(string[] args)
        {
            //parameters
            Console.Write("Excel-file location: ");
            string xlFile = Console.ReadLine();
            Console.Write("\rSheet-name: ");
            string sheetName = Console.ReadLine();
            Console.Write("\rRegex for finding matches: ");
            string rgx = Console.ReadLine();
            Console.Write("\rDelete empty rows? (yes/y): ");
            string dlEmptyRows = Console.ReadLine();


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //extract data from sheet
            using (var sheet = new ExcelPackage(xlFile))
            {
                List<string> results = new();
                int counter = 0;
                int i = 1;
                var ws = sheet.Workbook.Worksheets[sheetName];
                while (i <= ws.Dimension.End.Row)
                {
                    if (ws.Cells[i, 1].Value == null && (String.Equals(dlEmptyRows, "y") || String.Equals(dlEmptyRows, "yes")))
                    {
                        ws.DeleteRow(i);
                        continue;
                    }
                    else
                    {
                        Console.Write("\rMatches scraped: {0}", counter);
                        //get email from site
                        if (ws.Cells[i, 1].Value == null)
                        {
                            i++;
                            continue;
                        }
                        if (ws.Cells[i, 1].Text != null
                            || ws.Cells[i, 1].Hyperlink != null && Uri.IsWellFormedUriString(ws.Cells[i, 1].Hyperlink.ToString(), UriKind.Absolute))
                        {
                            string uri;
                            if (Uri.IsWellFormedUriString(ws.Cells[i, 1].Text, UriKind.Absolute))
                                uri = ws.Cells[i, 1].Text;
                            else if (ws.Cells[i, 1].Hyperlink != null && Uri.IsWellFormedUriString(ws.Cells[i, 1].Hyperlink.ToString(), UriKind.Absolute))
                                uri = ws.Cells[i, 1].Hyperlink.ToString();
                            else
                            {
                                i++;
                                continue;
                            }
                            try
                            {
                                string? result = rgxScrape(uri, rgx);
                                if (result == null)
                                {
                                    i++;
                                    continue;
                                }
                                else
                                {
                                    results.Add(result);
                                    ws.Cells[i, 2].Value = result;
                                    counter++;
                                    i++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error for cell {ws.Cells[i, 1].Address}: " + ex.Message);
                                i++;
                            }
                        }
                        else
                            i++;
                    }
                }
                sheet.Save();
                Console.WriteLine("\nSaved to Worksheet");
                Console.WriteLine("Emails scraped: " + results.Count.ToString());
            }
        }

        private static string? rgxScrape(string uri, string rgx)
        {
            HttpClient client = new();
            string html = client.GetStringAsync(uri).Result;
            var match = Regex.Match(html, rgx);
            if (!match.Success)
                return null;
            else
                return match.Groups[1].Value;
        }
    }
}