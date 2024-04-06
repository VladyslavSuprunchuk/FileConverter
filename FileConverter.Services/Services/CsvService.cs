using CsvHelper;
using FileConverter.Services.Interfaces;
using System.Globalization;

namespace FileConverter.Services.Services
{
    public class CsvService : ICsvService
    {
        public Stream CreateCsvReport(string content)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

            var lines = content.Split(" ").ToList();

            foreach (var item in lines)
            {
                csv.WriteField(item);
            }

            csv.Flush();
            writer.Flush();
            stream.Position = 0;

            return stream;
        }
    }
}
