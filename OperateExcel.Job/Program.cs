using Hangfire;
using Hangfire.SqlServer;
using OperateExcel.Job;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<ExcelImportOptions>(builder.Configuration.GetSection("ExcelImport"));
builder.Services.Configure<FeishuOptions>(builder.Configuration.GetSection("Feishu"));
builder.Services.PostConfigure<ExcelImportOptions>(options =>
{
    foreach (var arg in args)
    {
        if (arg.StartsWith("--date=", StringComparison.OrdinalIgnoreCase))
        {
            options.ProcessingDateOverride = arg["--date=".Length..];
        }
        else if (arg.StartsWith("--template=", StringComparison.OrdinalIgnoreCase))
        {
            options.TemplateFilePath = arg["--template=".Length..];
        }
        else if (arg.StartsWith("--output-dir=", StringComparison.OrdinalIgnoreCase))
        {
            options.OutputDirectory = arg["--output-dir=".Length..];
        }
    }
});
builder.Services.AddHttpClient<FeishuApiClient>();
builder.Services.AddSingleton<ExcelImportJob>();

var hangfireConnectionString = builder.Configuration.GetConnectionString("Hangfire");
if (!string.IsNullOrWhiteSpace(hangfireConnectionString))
{
    builder.Services.AddHangfire(config => config
        .SetDataCompatibilityLevel(CompatibilityLevel.Version_180)
        .UseSimpleAssemblyNameTypeSerializer()
        .UseRecommendedSerializerSettings()
        .UseSqlServerStorage(hangfireConnectionString, new SqlServerStorageOptions
        {
            PrepareSchemaIfNecessary = true
        }));
    builder.Services.AddHangfireServer();
}

var app = builder.Build();

if (args.Any(arg => string.Equals(arg, "--run-once", StringComparison.OrdinalIgnoreCase)))
{
    await app.Services.GetRequiredService<ExcelImportJob>().RunAsync();
    return;
}

if (!string.IsNullOrWhiteSpace(hangfireConnectionString))
{
    app.UseHangfireDashboard();

    var options = app.Services.GetRequiredService<Microsoft.Extensions.Options.IOptions<ExcelImportOptions>>().Value;
    RecurringJob.AddOrUpdate<ExcelImportJob>(
        "store-data-to-temp-xlsx",
        job => job.RunAsync(),
        options.DailyCron,
        new RecurringJobOptions { TimeZone = TimeZoneInfo.Local });
}

app.MapGet("/", () => "OperateExcel job is running.");
app.MapPost("/run", async (ExcelImportJob job) =>
{
    var result = await job.RunAsync();
    return Results.Ok(result);
});

await app.RunAsync();
