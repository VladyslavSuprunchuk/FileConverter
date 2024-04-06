namespace FileConverter.Services.Interfaces
{
    public interface ICsvService
    {
        Stream CreateCsvReport(string content);
    }
}
