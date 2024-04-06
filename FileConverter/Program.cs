using FileConverter.Core.Keywords;
using FileConverter.Services.Interfaces;
using FileConverter.Services.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IReportService, ReportService>();
builder.Services.AddScoped<ICsvService, CsvService>();
builder.Services.AddHttpClient(ClientKeywords.Title,
    client => 
    {
        client.BaseAddress = new Uri(ClientKeywords.BaseUrl);
    });

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
