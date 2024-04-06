using FileConverter.Services.Services;

namespace FileConverter.Services.Test
{
    public class CsvServiceTests
    {
        [Fact]
        public void CreateCsvReport_ReturnsStreamWithCorrectContent()
        {
            var csvService = new CsvService();
            var content = "test1 test2";

            using (var stream = csvService.CreateCsvReport(content))
            using (var reader = new StreamReader(stream))
            {
                var result = reader.ReadToEnd();

                Assert.Contains("test1,", result);
                Assert.Contains("test2", result);
            }
        }

        [Fact]
        public void CreateCsvReportReturnsStreamWithCorrectPosition()
        {
            var csvService = new CsvService();
            var content = "test1 test2";

            using (var stream = csvService.CreateCsvReport(content))
            {
                Assert.Equal(0, stream.Position);
            }
        }
    }
}