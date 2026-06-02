using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Globalization;

namespace OperateExcel.Job;

public sealed class DailyExcelImportWorker : BackgroundService
{
    private readonly ExcelImportJob _job;
    private readonly ExcelImportOptions _options;
    private readonly ILogger<DailyExcelImportWorker> _logger;

    public DailyExcelImportWorker(
        ExcelImportJob job,
        IOptions<ExcelImportOptions> options,
        ILogger<DailyExcelImportWorker> logger)
    {
        _job = job;
        _options = options.Value;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var schedule = DailySchedule.Parse(_options.DailyCron);
        _logger.LogInformation(
            "OperateExcel daily worker started. Schedule={DailyCron}, LocalRunTime={RunTime}.",
            _options.DailyCron,
            schedule.RunTime);

        while (!stoppingToken.IsCancellationRequested)
        {
            var nextRun = schedule.GetNextRun(DateTimeOffset.Now);
            var delay = nextRun - DateTimeOffset.Now;
            if (delay < TimeSpan.Zero)
            {
                delay = TimeSpan.Zero;
            }

            _logger.LogInformation("Next Excel import run scheduled at {NextRun:yyyy-MM-dd HH:mm:ss zzz}.", nextRun);

            try
            {
                await Task.Delay(delay, stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }

            if (stoppingToken.IsCancellationRequested)
            {
                break;
            }

            await RunJobAsync(stoppingToken);
        }

        _logger.LogInformation("OperateExcel daily worker stopped.");
    }

    private async Task RunJobAsync(CancellationToken stoppingToken)
    {
        try
        {
            _logger.LogInformation("Excel import job started.");
            var result = await _job.RunAsync();
            _logger.LogInformation(
                "Excel import job completed. Date={ProcessingDate:yyyy-MM-dd}, Output={OutputFile}, Fulfillment={FulfillmentRows}, Payments={PaymentRows}, Advertising={AdvertisingRows}.",
                result.ProcessingDate,
                result.TemplateFilePath,
                result.FulfillmentRows,
                result.PaymentRows,
                result.AdvertisingRows);

            foreach (var message in result.Messages)
            {
                _logger.LogInformation("Excel import detail: {Message}", message);
            }
        }
        catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception exception)
        {
            _logger.LogError(exception, "Excel import job failed.");
        }
    }

    private readonly record struct DailySchedule(TimeOnly RunTime)
    {
        public static DailySchedule Parse(string? cron)
        {
            var text = string.IsNullOrWhiteSpace(cron) ? "0 2 * * *" : cron.Trim();
            var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 5
                || parts[2] != "*"
                || parts[3] != "*"
                || parts[4] != "*"
                || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var minute)
                || !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var hour)
                || hour is < 0 or > 23
                || minute is < 0 or > 59)
            {
                throw new InvalidOperationException(
                    $"Unsupported DailyCron '{cron}'. Expected a daily cron expression like '0 2 * * *'.");
            }

            return new DailySchedule(new TimeOnly(hour, minute));
        }

        public DateTimeOffset GetNextRun(DateTimeOffset now)
        {
            var todayRun = new DateTimeOffset(
                now.Year,
                now.Month,
                now.Day,
                RunTime.Hour,
                RunTime.Minute,
                0,
                now.Offset);

            return todayRun > now
                ? todayRun
                : todayRun.AddDays(1);
        }
    }
}
