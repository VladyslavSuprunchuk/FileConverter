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

                    var tableData = ReadTables(workSheet);
                    content = GetFirstTablesContent(tableData);
                }
            }

            return content;
        }

        private static string GetFirstTablesContent(string tableData)
        {
            using (var reader = new StringReader(tableData))
            {
                var extractedData = new StringBuilder();

                for (var i = 1; i <= Int32.MaxValue; i++)
                {
                    var line = reader.ReadLine();
                    if (i >= StartIndexOfTable && i < EndIndexOfTable && line != null)
                    {
                        extractedData.AppendLine(line);

                        if (i == EndIndexOfTable - 1)
                        {
                            break;
                        }
                    }
                }

                return extractedData.ToString();
            }
        }

        private static string ReadTables(IXLWorksheet worksheet)
        {
            using (var sw = new StringWriter())
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.Cells())
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
