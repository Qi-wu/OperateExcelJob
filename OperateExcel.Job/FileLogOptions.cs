using Microsoft.Extensions.Logging;

namespace OperateExcel.Job;

public sealed class FileLogOptions
{
    public string Directory { get; set; } = "logs";
    public string FileNamePrefix { get; set; } = "operate-excel-";
    public LogLevel MinimumLevel { get; set; } = LogLevel.Information;
}
