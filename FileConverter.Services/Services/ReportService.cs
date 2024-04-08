using ClosedXML.Excel;
using FileConverter.Core.Keywords;
using FileConverter.Services.Interfaces;
using HtmlAgilityPack;
using System.Text;

namespace FileConverter.Services.Services
{
    public class ReportService : IReportService
    {
        private readonly HttpClient _client;

        private const int StartIndexOfTable = 2;
        private const int EndIndexOfTable = 30;

        public ReportService(IHttpClientFactory httpClientFactory)
        {
            _client = httpClientFactory.CreateClient(ClientKeywords.Title);
        }

        public async Task<string> GetWorldwideReportAsync()
        {
            var htmlcontent = await GetReportsAsync();
            var linkToFile = ParseLink(htmlcontent, "Worldwide Rig Count");
            var report = await GetReportContentByLinkAsync(linkToFile);

            return report;
        }

        private static string ParseLink(string htmlContent, string linkName)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);
            var currentYear = DateTime.Now.Year;

            var link = doc.DocumentNode.Descendants("a")
                .Where(a =>
                {
                    var text = a.InnerText;
                    return text.StartsWith(linkName) && text.EndsWith(currentYear.ToString());
                }).FirstOrDefault();

            if (link != null)
            {
                var text = link.GetAttributeValue("href", "");

                return text;
            }

            return string.Empty;
        }

        private async Task<string> GetReportsAsync()
        {
            var url = $"{_client.BaseAddress}{ClientKeywords.FilesPath}";

            using var requestMessage = new HttpRequestMessage(HttpMethod.Get, url);
            requestMessage.Headers.Add(ClientKeywords.ConnectionHeader, ClientKeywords.KeepAliveHeaderValue);
            requestMessage.Headers.Add(ClientKeywords.CookieHeader, ClientKeywords.CookieHeaderValue);

            var response = await _client.SendAsync(requestMessage);

            if (response.IsSuccessStatusCode)
            {
                var htmlContent = await response.Content.ReadAsStringAsync();

                return htmlContent;
            }

            return string.Empty;
        }

        private async Task<string> GetReportContentByLinkAsync(string linkToFile)
        {
            var url = $"{_client.BaseAddress}{linkToFile}";
            var content = string.Empty;

            using var requestMessage = new HttpRequestMessage(HttpMethod.Get, url);
            requestMessage.Headers.Add(ClientKeywords.ConnectionHeader, ClientKeywords.KeepAliveHeaderValue);
            requestMessage.Headers.Add(ClientKeywords.CookieHeader, ClientKeywords.CookieHeaderValue);

            var response = await _client.SendAsync(requestMessage);

            if (response.IsSuccessStatusCode)
            {
                var stream = await response.Content.ReadAsStreamAsync();

                using (var workbook = new XLWorkbook(stream))
                {
                    var workSheet = workbook.Worksheets.FirstOrDefault();

                    if (workSheet is null)
                    {
                        return content;
                    }

                    var pastYear = DateTime.Now.Year - 2;
                    var pastYearRow = FindRowWithYear(workSheet, pastYear);
                    var currentYearRow = FindRowWithYear(workSheet, DateTime.Now.Year);

                    if (pastYearRow != null && currentYearRow != null)
                    {
                        content = ReadTables(workSheet, pastYearRow.Value, currentYearRow.Value);
                    }
                }
            }

            return content;
        }

        private static int? FindRowWithYear(IXLWorksheet worksheet, int year)
        {
            for (var row = 1; row <= worksheet.LastRowUsed().RowNumber(); row++)
            {
                var cell = worksheet.Cell(row, "B");
                int cellYear;
                if (int.TryParse(cell.Value.ToString(), out cellYear) && cellYear == year)
                {
                    return row;
                }
            }
            return null;
        }

        private static string ReadTables(IXLWorksheet worksheet, int endRow, int startRow)
        {
            using (var sw = new StringWriter())
            {
                for (var row = startRow; row < endRow; row++)
                {
                    foreach (var cell in worksheet.Row(row).Cells())
                    {
                        sw.Write(cell.Value);
                        sw.Write(" ");
                    }
                    sw.WriteLine();
                }

                return sw.ToString();
            }
        }
    }
}
