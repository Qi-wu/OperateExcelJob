using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OperateExcel.Job;

public sealed class ExcelImportJob
{
    private const string AdvertisingSheetName = "\u5e7f\u544a";

    private static readonly IReadOnlyDictionary<string, string> SheetSourceExtensions =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["fulfillment"] = ".txt",
            ["payments"] = ".csv",
            [AdvertisingSheetName] = ".xlsx"
        };

    private static readonly IReadOnlyList<string> StoreReadOrder =
    [
        "\u65e0\u5fe7\u65e0\u8651",
        "Oyumoents",
        "Yue an Company",
        "DUX"
    ];

    private static readonly IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> SheetHeaderAliases =
        new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            [AdvertisingSheetName] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["\u70b9\u51fb\u7387(CTR)"] = "\u70b9\u51fb\u7387 (CTR)",
                ["\u6bcf\u6b21\u70b9\u51fb\u6210\u672c(CPC)"] = "\u5355\u6b21\u70b9\u51fb\u6210\u672c (CPC)"
            }
        };

    private static readonly Regex CurrencyTokenRegex = new(
        "(?i)(US\\$|USD|CNY|RMB|EUR|GBP|[$\u20ac\u00a3\u00a5\uffe5])",
        RegexOptions.Compiled);

    private static readonly Regex ExcelCurrencyFormatRegex = new(
        "\\[\\$([^\\]-]+)-[^\\]]+\\]",
        RegexOptions.Compiled);

    private readonly ExcelImportOptions _options;
    private readonly ILogger<ExcelImportJob> _logger;

    public ExcelImportJob(IOptions<ExcelImportOptions> options, ILogger<ExcelImportJob> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public async Task<ImportResult> RunAsync()
    {
        return await Task.Run(Run);
    }

    private ImportResult Run()
    {
        var messages = new List<string>();
        var processingDate = ResolveProcessingDate();
        var dateFolderName = processingDate.ToString("yyyy-MM-dd");
        var sourceDirectory = Path.Combine(_options.RootDirectory, dateFolderName, dateFolderName);

        if (!Directory.Exists(sourceDirectory))
        {
            throw new DirectoryNotFoundException($"Source directory not found: {sourceDirectory}");
        }

        if (!File.Exists(_options.TemplateFilePath))
        {
            throw new FileNotFoundException("Template file not found.", _options.TemplateFilePath);
        }

        var storeDirectories = Directory.GetDirectories(sourceDirectory);
        if (storeDirectories.Length == 0)
        {
            throw new DirectoryNotFoundException($"No store directories found under: {sourceDirectory}");
        }

        IWorkbook workbook;
        using (var templateStream = File.Open(_options.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            workbook = WorkbookFactory.Create(templateStream);
        }
        var formatter = new DataFormatter();
        var cellStyleCache = new CellStyleCache(workbook);

        var importedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["fulfillment"] = 0,
            ["payments"] = 0,
            [AdvertisingSheetName] = 0
        };

        foreach (var (sheetName, extension) in SheetSourceExtensions)
        {
            var targetSheet = workbook.GetSheet(sheetName)
                ?? throw new InvalidOperationException($"Sheet not found in template: {sheetName}");

            var targetHeaderRowIndex = FindHeaderRow(targetSheet, formatter);
            if (targetHeaderRowIndex < 0)
            {
                messages.Add($"Skipped sheet {sheetName}: no header row found.");
                continue;
            }

            if (_options.ClearExistingData)
            {
                ClearRowsBelow(targetSheet, targetHeaderRowIndex);
            }

            var targetHeaders = ReadRow(targetSheet.GetRow(targetHeaderRowIndex), formatter)
                .Select(DelimitedTableReader.CleanHeader)
                .ToList();

            foreach (var storeDirectory in OrderStoreDirectories(storeDirectories))
            {
                var sourceFile = Directory.GetFiles(storeDirectory, $"*{extension}", SearchOption.TopDirectoryOnly)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault();

                if (sourceFile is null)
                {
                    messages.Add($"Store {Path.GetFileName(storeDirectory)} has no {extension} file for sheet {sheetName}.");
                    continue;
                }

                var sourceTable = extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                    ? ReadExcelTable(sourceFile, formatter)
                    : extension.Equals(".csv", StringComparison.OrdinalIgnoreCase)
                        ? DelimitedTableReader.ReadCsv(sourceFile)
                        : DelimitedTableReader.ReadTxt(sourceFile);

                var sourceIndex = BuildHeaderIndex(sourceTable.Headers);
                var writableColumns = targetHeaders
                    .Select((header, columnIndex) => new
                    {
                        Header = header,
                        TargetColumnIndex = columnIndex,
                        SourceColumnIndex = sourceIndex.TryGetValue(ResolveSourceHeader(sheetName, header), out var sourceColumnIndex)
                            ? sourceColumnIndex
                            : -1
                    })
                    .Where(column => column.Header.Length > 0 && column.SourceColumnIndex >= 0)
                    .ToList();

                if (writableColumns.Count == 0)
                {
                    messages.Add($"Skipped {sourceFile}: no matching columns for sheet {sheetName}.");
                    continue;
                }

                foreach (var sourceRow in sourceTable.Rows)
                {
                    var targetRow = targetSheet.CreateRow(targetSheet.LastRowNum + 1);
                    foreach (var column in writableColumns)
                    {
                        var value = column.SourceColumnIndex < sourceRow.Count
                            ? sourceRow[column.SourceColumnIndex]
                            : string.Empty;

                        SetCellValue(targetRow.CreateCell(column.TargetColumnIndex), value, cellStyleCache);
                    }

                    importedCounts[sheetName]++;
                }

                messages.Add($"Imported {sourceTable.Rows.Count} rows from {sourceFile} to sheet {sheetName}.");
            }
        }

        var tempOutput = Path.Combine(
            Path.GetDirectoryName(_options.TemplateFilePath)!,
            $"{Path.GetFileNameWithoutExtension(_options.TemplateFilePath)}.{DateTime.Now:yyyyMMddHHmmss}.tmp.xlsx");

        using (var outputStream = File.Create(tempOutput))
        {
            workbook.Write(outputStream);
        }

        workbook.Close();
        RemoveReadOnlyAttribute(_options.TemplateFilePath);
        File.Move(tempOutput, _options.TemplateFilePath, overwrite: true);

        var result = new ImportResult(
            processingDate,
            sourceDirectory,
            _options.TemplateFilePath,
            importedCounts["fulfillment"],
            importedCounts["payments"],
            importedCounts[AdvertisingSheetName],
            messages);

        _logger.LogInformation(
            "Excel import finished. Fulfillment={FulfillmentRows}, Payments={PaymentRows}, Advertising={AdvertisingRows}",
            result.FulfillmentRows,
            result.PaymentRows,
            result.AdvertisingRows);

        return result;
    }

    private DateOnly ResolveProcessingDate()
    {
        if (!string.IsNullOrWhiteSpace(_options.ProcessingDateOverride)
            && DateOnly.TryParse(_options.ProcessingDateOverride, out var overrideDate))
        {
            return overrideDate;
        }

        return DateOnly.FromDateTime(DateTime.Today.AddDays(_options.DateOffsetDays));
    }

    private static IEnumerable<string> OrderStoreDirectories(IEnumerable<string> storeDirectories)
    {
        var order = StoreReadOrder
            .Select((storeName, index) => new { storeName, index })
            .ToDictionary(item => item.storeName, item => item.index, StringComparer.OrdinalIgnoreCase);

        return storeDirectories
            .OrderBy(path => order.TryGetValue(Path.GetFileName(path), out var index) ? index : int.MaxValue)
            .ThenBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase);
    }

    private static string ResolveSourceHeader(string sheetName, string targetHeader)
    {
        return SheetHeaderAliases.TryGetValue(sheetName, out var aliases)
            && aliases.TryGetValue(targetHeader, out var sourceHeader)
                ? sourceHeader
                : targetHeader;
    }

    private static TableData ReadExcelTable(string path, DataFormatter formatter)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = WorkbookFactory.Create(stream);
        var sheet = workbook.GetSheetAt(0);
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        if (headerRowIndex < 0)
        {
            workbook.Close();
            return new TableData([], []);
        }

        var headers = ReadRow(sheet.GetRow(headerRowIndex), formatter)
            .Select(DelimitedTableReader.CleanHeader)
            .ToList();

        var rows = new List<IReadOnlyList<string>>();
        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var values = ReadRow(row, formatter, headers.Count);
            if (values.Any(value => !string.IsNullOrWhiteSpace(value)))
            {
                rows.Add(values);
            }
        }

        workbook.Close();
        return new TableData(headers, rows);
    }

    private static int FindHeaderRow(ISheet sheet, DataFormatter formatter)
    {
        var maxRowsToScan = Math.Min(sheet.LastRowNum, 30);
        for (var rowIndex = sheet.FirstRowNum; rowIndex <= maxRowsToScan; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var nonEmptyCellCount = ReadRow(row, formatter)
                .Count(value => !string.IsNullOrWhiteSpace(value));

            if (nonEmptyCellCount >= 2)
            {
                return rowIndex;
            }
        }

        return -1;
    }

    private static IReadOnlyList<string> ReadRow(IRow row, DataFormatter formatter, int? maxColumns = null)
    {
        var columnCount = maxColumns ?? row.LastCellNum;
        var values = new List<string>();
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            values.Add(formatter.FormatCellValue(row.GetCell(columnIndex)).Trim());
        }

        return values;
    }

    private static Dictionary<string, int> BuildHeaderIndex(IReadOnlyList<string> headers)
    {
        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Count; i++)
        {
            var header = DelimitedTableReader.CleanHeader(headers[i]);
            if (header.Length > 0 && !index.ContainsKey(header))
            {
                index.Add(header, i);
            }
        }

        return index;
    }

    private static void SetCellValue(ICell cell, string value, CellStyleCache styleCache)
    {
        if (TryParseNumericCell(value, out var parsed))
        {
            cell.SetCellValue(parsed.Value);
            if (parsed.FormatKey is not null)
            {
                cell.CellStyle = styleCache.Get(parsed.FormatKey);
            }

            return;
        }

        cell.SetCellValue(value);
    }

    private static bool TryParseNumericCell(string value, out ParsedNumericCell parsed)
    {
        parsed = default;
        var text = value.Trim();
        if (text.Length == 0)
        {
            return false;
        }

        text = NormalizeExcelCurrencyFormat(text);

        var isPercent = text.EndsWith('%');
        var currencySymbol = ResolveCurrencySymbol(text);
        var numberText = CurrencyTokenRegex.Replace(text, string.Empty)
            .Replace("%", string.Empty)
            .Trim();

        var isParenthesizedNegative = numberText.StartsWith('(') && numberText.EndsWith(')');
        if (isParenthesizedNegative)
        {
            numberText = "-" + numberText[1..^1];
        }

        numberText = numberText.Replace(",", string.Empty);
        if (numberText.Length > 1 && numberText[0] == '0' && char.IsDigit(numberText[1]))
        {
            return false;
        }

        if (!double.TryParse(numberText, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var number))
        {
            return false;
        }

        var formatKey = currencySymbol is not null
            ? $"currency:{currencySymbol}"
            : isPercent
                ? "percent"
                : null;

        parsed = new ParsedNumericCell(isPercent ? number / 100D : number, formatKey);
        return true;
    }

    private static string? ResolveCurrencySymbol(string text)
    {
        text = NormalizeExcelCurrencyFormat(text);

        var match = CurrencyTokenRegex.Match(text);
        if (!match.Success)
        {
            return null;
        }

        return match.Value.ToUpperInvariant() switch
        {
            "USD" or "US$" or "$" => "$",
            "CNY" or "RMB" or "\u00a5" or "\uffe5" => "\u00a5",
            "EUR" or "\u20ac" => "\u20ac",
            "GBP" or "\u00a3" => "\u00a3",
            _ => match.Value
        };
    }

    private static string NormalizeExcelCurrencyFormat(string text)
    {
        return ExcelCurrencyFormatRegex.Replace(text, match => match.Groups[1].Value);
    }

    private static void ClearRowsBelow(ISheet sheet, int headerRowIndex)
    {
        for (var rowIndex = sheet.LastRowNum; rowIndex > headerRowIndex; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is not null)
            {
                sheet.RemoveRow(row);
            }
        }
    }

    private static void RemoveReadOnlyAttribute(string path)
    {
        var attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
        {
            File.SetAttributes(path, attributes & ~FileAttributes.ReadOnly);
        }
    }

    private readonly record struct ParsedNumericCell(double Value, string? FormatKey);

    private sealed class CellStyleCache
    {
        private readonly IWorkbook _workbook;
        private readonly IDataFormat _dataFormat;
        private readonly Dictionary<string, ICellStyle> _styles = new(StringComparer.OrdinalIgnoreCase);

        public CellStyleCache(IWorkbook workbook)
        {
            _workbook = workbook;
            _dataFormat = workbook.CreateDataFormat();
        }

        public ICellStyle Get(string key)
        {
            if (_styles.TryGetValue(key, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            style.DataFormat = _dataFormat.GetFormat(ResolveFormat(key));
            _styles.Add(key, style);
            return style;
        }

        private static string ResolveFormat(string key)
        {
            if (string.Equals(key, "percent", StringComparison.OrdinalIgnoreCase))
            {
                return "0.00%";
            }

            if (key.StartsWith("currency:", StringComparison.OrdinalIgnoreCase))
            {
                var symbol = key["currency:".Length..];
                return $"{symbol}#,##0.00";
            }

            return "General";
        }
    }
}
