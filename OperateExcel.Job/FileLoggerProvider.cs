using Microsoft.Extensions.Logging;
using System.Text;

namespace OperateExcel.Job;

public sealed class FileLoggerProvider : ILoggerProvider
{
    private readonly FileLogOptions _options;
    private readonly object _syncRoot = new();

    public FileLoggerProvider(FileLogOptions options)
    {
        _options = options;
    }

    public ILogger CreateLogger(string categoryName)
    {
        return new FileLogger(categoryName, _options, _syncRoot);
    }

    public void Dispose()
    {
    }

    private sealed class FileLogger : ILogger
    {
        private readonly string _categoryName;
        private readonly FileLogOptions _options;
        private readonly object _syncRoot;

        public FileLogger(string categoryName, FileLogOptions options, object syncRoot)
        {
            _categoryName = categoryName;
            _options = options;
            _syncRoot = syncRoot;
        }

        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull
        {
            return null;
        }

        public bool IsEnabled(LogLevel logLevel)
        {
            return logLevel != LogLevel.None && logLevel >= _options.MinimumLevel;
        }

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel))
            {
                return;
            }

            var message = formatter(state, exception);
            if (string.IsNullOrWhiteSpace(message) && exception is null)
            {
                return;
            }

            var now = DateTimeOffset.Now;
            var builder = new StringBuilder()
                .Append(now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz"))
                .Append(" [")
                .Append(logLevel)
                .Append("] ")
                .Append(_categoryName)
                .Append(": ")
                .AppendLine(message);

            if (exception is not null)
            {
                builder.AppendLine(exception.ToString());
            }

            var logPath = ResolveLogPath(now);
            lock (_syncRoot)
            {
                System.IO.Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
                File.AppendAllText(logPath, builder.ToString(), Encoding.UTF8);
            }
        }

        private string ResolveLogPath(DateTimeOffset now)
        {
            var directory = Path.IsPathRooted(_options.Directory)
                ? _options.Directory
                : Path.Combine(AppContext.BaseDirectory, _options.Directory);
            return Path.Combine(directory, $"{_options.FileNamePrefix}{now:yyyyMMdd}.log");
        }
    }
}
