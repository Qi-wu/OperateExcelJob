using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OperateExcel.Job;

var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
{
    Args = args,
    ContentRootPath = AppContext.BaseDirectory
});

builder.Services.Configure<ExcelImportOptions>(builder.Configuration.GetSection("ExcelImport"));
builder.Services.Configure<FileLogOptions>(builder.Configuration.GetSection("FileLog"));
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
        else if (arg.StartsWith("--sku-owner-mapping=", StringComparison.OrdinalIgnoreCase))
        {
            options.SkuOwnerMappingFilePath = arg["--sku-owner-mapping=".Length..];
        }
    }
});

builder.Logging.ClearProviders();
builder.Logging.AddConsole();
builder.Logging.AddProvider(new FileLoggerProvider(builder.Configuration.GetSection("FileLog").Get<FileLogOptions>() ?? new FileLogOptions()));

builder.Services.AddWindowsService(options =>
{
    var serviceName = builder.Configuration.GetValue<string>("WindowsService:ServiceName");
    options.ServiceName = string.IsNullOrWhiteSpace(serviceName)
        ? "OperateExcelJob"
        : serviceName;
});

builder.Services.AddHttpClient<FeishuApiClient>();
builder.Services.AddSingleton<ExcelImportJob>();

if (args.Any(arg => string.Equals(arg, "--run-once", StringComparison.OrdinalIgnoreCase)))
{
    using var runOnceHost = builder.Build();
    await runOnceHost.Services.GetRequiredService<ExcelImportJob>().RunAsync();
    return;
}

builder.Services.AddHostedService<DailyExcelImportWorker>();

await builder.Build().RunAsync();
