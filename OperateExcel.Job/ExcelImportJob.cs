using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;

namespace OperateExcel.Job;

public sealed class ExcelImportJob
{
    private static readonly IReadOnlyDictionary<string, string> SheetSourceExtensions =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["fulfillment"] = ".txt",
            ["payments"] = ".csv",
            ["广告"] = ".xlsx"
        };

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

        var importedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["fulfillment"] = 0,
            ["payments"] = 0,
            ["广告"] = 0
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

            foreach (var storeDirectory in storeDirectories.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
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
                        SourceColumnIndex = sourceIndex.TryGetValue(header, out var sourceColumnIndex)
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

                        targetRow.CreateCell(column.TargetColumnIndex, CellType.String).SetCellValue(value);
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
            importedCounts["广告"],
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
}
