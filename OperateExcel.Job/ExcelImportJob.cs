using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OperateExcel.Job;

public sealed class ExcelImportJob
{
    private const string AdvertisingSheetName = "\u5e7f\u544a";
    private const string FulfillmentSheetName = "fulfillment";
    private const string PaymentsSheetName = "payments";
    private const string FulfillmentTemplateSheetName = "\u6a21\u7248F";
    private const string PaymentTemplateSheetName = "\u6a21\u677fP";
    private const string WaitingOrderFileName = "\u7b49\u5f85\u4e2d\u7684\u8ba2\u5355\u53f7.xlsx";
    private const string AmazonOrderIdHeader = "amazon-order-id";
    private const string OrderStatusHeader = "order-status";
    private const int FulfillmentCopyStartColumnIndex = 11; // Excel column L.
    private const int FulfillmentCopyEndColumnIndex = 19; // Excel column T.

    private static readonly IReadOnlyDictionary<string, string> SheetSourceExtensions =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [FulfillmentSheetName] = ".txt",
            [PaymentsSheetName] = ".csv",
            [AdvertisingSheetName] = ".xlsx"
        };

    private static readonly IReadOnlyDictionary<string, string> PaymentTemplateColumnMap =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["sku"] = "sku",
            ["quantity"] = "quantity",
            ["product sales"] = "item-price",
            ["product sales tax"] = "item-tax",
            ["shipping credits"] = "shipping-price",
            ["gift wrap credits"] = "gift wrap credits",
            ["total"] = "total"
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

    private static readonly Regex ChineseMonthDateRegex = new(
        @"^(?<month>\d{1,2})\u6708\s+(?<day>\d{1,2}),?\s+(?<year>\d{4})$",
        RegexOptions.Compiled);

    private static readonly Regex CellReferenceRegex = new(
        @"(?<![A-Za-z0-9_])(?<column>\$?[A-Za-z]{1,3})(?<absoluteRow>\$?)(?<row>\d+)(?![A-Za-z0-9_])",
        RegexOptions.Compiled);

    private readonly ExcelImportOptions _options;
    private readonly FeishuOptions _feishuOptions;
    private readonly FeishuApiClient _feishuClient;
    private readonly ILogger<ExcelImportJob> _logger;

    public ExcelImportJob(
        IOptions<ExcelImportOptions> options,
        IOptions<FeishuOptions> feishuOptions,
        FeishuApiClient feishuClient,
        ILogger<ExcelImportJob> logger)
    {
        _options = options.Value;
        _feishuOptions = feishuOptions.Value;
        _feishuClient = feishuClient;
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
        var sourceDirectory = Path.Combine(_options.RootDirectory, dateFolderName);

        if (!Directory.Exists(sourceDirectory))
        {
            throw new DirectoryNotFoundException($"Source directory not found: {sourceDirectory}");
        }

        if (!File.Exists(_options.TemplateFilePath))
        {
            throw new FileNotFoundException("Template file not found.", _options.TemplateFilePath);
        }

        Directory.CreateDirectory(_options.OutputDirectory);
        var outputFilePath = Path.Combine(_options.OutputDirectory, BuildReportFileName(processingDate));
        File.Copy(_options.TemplateFilePath, outputFilePath, overwrite: true);
        RemoveReadOnlyAttribute(outputFilePath);

        var storeDirectories = Directory.GetDirectories(sourceDirectory);
        if (storeDirectories.Length == 0)
        {
            throw new DirectoryNotFoundException($"No store directories found under: {sourceDirectory}");
        }

        IWorkbook workbook;
        using (var templateStream = File.Open(outputFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            workbook = WorkbookFactory.Create(templateStream);
        }
        var formatter = new DataFormatter();
        var cellStyleCache = new CellStyleCache(workbook);

        ImportB2BOlFromFeishu(workbook, processingDate, formatter, cellStyleCache, messages);

        var importedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            [FulfillmentSheetName] = 0,
            [PaymentsSheetName] = 0,
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

            var targetHeaders = ReadRow(targetSheet.GetRow(targetHeaderRowIndex), formatter)
                .Select(DelimitedTableReader.CleanHeader)
                .ToList();

            var nextTargetRowIndex = targetHeaderRowIndex + 1;

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

                if (_options.ClearExistingData && nextTargetRowIndex == targetHeaderRowIndex + 1)
                {
                    ClearWritableCells(targetSheet, targetHeaderRowIndex, writableColumns.Select(column => column.TargetColumnIndex));
                }

                if (writableColumns.Count == 0)
                {
                    messages.Add($"Skipped {sourceFile}: no matching columns for sheet {sheetName}.");
                    continue;
                }

                foreach (var sourceRow in sourceTable.Rows)
                {
                    var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
                    foreach (var column in writableColumns)
                    {
                        var value = column.SourceColumnIndex < sourceRow.Count
                            ? sourceRow[column.SourceColumnIndex]
                            : string.Empty;

                        SetCellValue(targetRow.CreateCell(column.TargetColumnIndex), value, cellStyleCache);
                    }

                    importedCounts[sheetName]++;
                    nextTargetRowIndex++;
                }

                messages.Add($"Imported {sourceTable.Rows.Count} rows from {sourceFile} to sheet {sheetName}.");
            }
        }

        ApplyPostImportRules(workbook, processingDate, formatter, cellStyleCache, messages);

        var tempOutput = Path.Combine(
            _options.OutputDirectory,
            $"{Path.GetFileNameWithoutExtension(outputFilePath)}.{DateTime.Now:yyyyMMddHHmmss}.tmp.xlsx");

        using (var outputStream = File.Create(tempOutput))
        {
            workbook.Write(outputStream);
        }

        RemoveStaleCalculationMetadata(tempOutput);
        workbook.Close();
        RemoveReadOnlyAttribute(outputFilePath);
        File.Move(tempOutput, outputFilePath, overwrite: true);

        var result = new ImportResult(
            processingDate,
            sourceDirectory,
            outputFilePath,
            importedCounts[FulfillmentSheetName],
            importedCounts[PaymentsSheetName],
            importedCounts[AdvertisingSheetName],
            messages);

        _logger.LogInformation(
            "Excel import finished. Fulfillment={FulfillmentRows}, Payments={PaymentRows}, Advertising={AdvertisingRows}",
            result.FulfillmentRows,
            result.PaymentRows,
            result.AdvertisingRows);

        return result;
    }

    private void ImportB2BOlFromFeishu(
        IWorkbook targetWorkbook,
        DateOnly processingDate,
        DataFormatter formatter,
        CellStyleCache cellStyleCache,
        ICollection<string> messages)
    {
        if (!_feishuOptions.Enabled)
        {
            messages.Add("Skipped Feishu B2B（ol） import: Feishu import is disabled.");
            _logger.LogWarning("Skipped Feishu B2B（ol） import because Feishu import is disabled.");
            return;
        }

        if (!_feishuClient.IsConfigured)
        {
            throw new InvalidOperationException("Feishu B2B（ol） import is enabled, but AppId/AppSecret or table configuration is incomplete.");
        }

        var attachmentBytes = _feishuClient
            .DownloadProfitAttachmentAsync(processingDate)
            .GetAwaiter()
            .GetResult();

        using var stream = new MemoryStream(attachmentBytes);
        var sourceWorkbook = WorkbookFactory.Create(stream);
        try
        {
            var sourceTable = ReadExcelTable(sourceWorkbook, _feishuOptions.SourceSheetName, formatter);
            var targetSheet = FindSheet(targetWorkbook, _feishuOptions.TargetSheetName)
                ?? throw new InvalidOperationException($"Sheet not found in template: {_feishuOptions.TargetSheetName}");

            var importedRows = CopyTableDataToSheet(sourceTable, targetSheet, formatter, cellStyleCache);
            messages.Add($"Imported {importedRows} rows from Feishu sheet {_feishuOptions.SourceSheetName} to sheet {_feishuOptions.TargetSheetName}.");
        }
        finally
        {
            sourceWorkbook.Close();
        }
    }

    private void ApplyPostImportRules(
        IWorkbook workbook,
        DateOnly processingDate,
        DataFormatter formatter,
        CellStyleCache cellStyleCache,
        ICollection<string> messages)
    {
        var waitingOrderIds = ReadWaitingOrderIds(processingDate, formatter);
        var highlightedOrderIds = HighlightWaitingFulfillmentOrders(workbook, waitingOrderIds, formatter, cellStyleCache);

        var copiedFulfillmentRows = CopyFilteredFulfillmentRowsToTemplate(workbook, formatter, cellStyleCache);
        var copiedPaymentRows = CopyOrderPaymentRowsToTemplate(workbook, formatter, cellStyleCache);
        MarkWorkbookForFormulaRecalculation(workbook);

        messages.Add($"Highlighted {highlightedOrderIds} fulfillment rows from {WaitingOrderFileName}.");
        messages.Add($"Copied {copiedFulfillmentRows} fulfillment rows to sheet {FulfillmentTemplateSheetName}.");
        messages.Add($"Copied {copiedPaymentRows} payment rows to sheet {PaymentTemplateSheetName}.");
    }

    private HashSet<string> ReadWaitingOrderIds(DateOnly processingDate, DataFormatter formatter)
    {
        var waitingOrderPath = Path.Combine(
            _options.RootDirectory,
            processingDate.ToString("yyyy-MM-dd"),
            WaitingOrderFileName);

        if (!File.Exists(waitingOrderPath))
        {
            _logger.LogWarning("Waiting order file not found: {WaitingOrderPath}", waitingOrderPath);
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        using var stream = File.Open(waitingOrderPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = WorkbookFactory.Create(stream);
        try
        {
            var sheet = workbook.GetSheetAt(0);
            var headerRowIndex = FindHeaderRow(sheet, formatter);
            if (headerRowIndex < 0)
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            var headers = ReadRow(sheet.GetRow(headerRowIndex), formatter)
                .Select(DelimitedTableReader.CleanHeader)
                .ToList();
            var headerIndex = BuildHeaderIndex(headers);
            var orderColumnIndex = ResolveWaitingOrderColumnIndex(headerIndex);
            if (orderColumnIndex < 0)
            {
                orderColumnIndex = headers.Count > 1 ? 1 : 0;
            }

            var orderIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                {
                    continue;
                }

                var orderId = NormalizeOrderId(formatter.FormatCellValue(row.GetCell(orderColumnIndex)));
                if (orderId.Length > 0)
                {
                    orderIds.Add(orderId);
                }
            }

            return orderIds;
        }
        finally
        {
            workbook.Close();
        }
    }

    private static int HighlightWaitingFulfillmentOrders(
        IWorkbook workbook,
        ISet<string> waitingOrderIds,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        if (waitingOrderIds.Count == 0)
        {
            return 0;
        }

        var sheet = workbook.GetSheet(FulfillmentSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentSheetName}");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var orderColumnIndex = ResolveRequiredColumn(headerIndex, AmazonOrderIdHeader, FulfillmentSheetName);

        var highlightedRows = 0;
        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var orderCell = row.GetCell(orderColumnIndex);
            var orderId = NormalizeOrderId(formatter.FormatCellValue(orderCell));
            if (orderId.Length == 0 || !waitingOrderIds.Contains(orderId))
            {
                continue;
            }

            orderCell ??= row.CreateCell(orderColumnIndex);
            orderCell.CellStyle = cellStyleCache.GetRedFontStyle(orderCell.CellStyle);
            highlightedRows++;
        }

        return highlightedRows;
    }

    private static int CopyFilteredFulfillmentRowsToTemplate(
        IWorkbook workbook,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var sourceSheet = workbook.GetSheet(FulfillmentSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentSheetName}");
        var targetSheet = workbook.GetSheet(FulfillmentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentTemplateSheetName}");

        var sourceHeaderRowIndex = FindHeaderRow(sourceSheet, formatter);
        var targetHeaderRowIndex = FindHeaderRow(targetSheet, formatter);
        var sourceHeaders = ReadHeaders(sourceSheet, sourceHeaderRowIndex, formatter);
        var targetHeaderIndex = ReadHeaderIndex(targetSheet, targetHeaderRowIndex, formatter, preferLastDuplicate: true);
        var sourceHeaderIndex = BuildHeaderIndex(sourceHeaders);
        var orderColumnIndex = ResolveRequiredColumn(sourceHeaderIndex, AmazonOrderIdHeader, FulfillmentSheetName);
        var statusColumnIndex = ResolveRequiredColumn(sourceHeaderIndex, OrderStatusHeader, FulfillmentSheetName);

        var copyColumns = Enumerable.Range(
                FulfillmentCopyStartColumnIndex,
                FulfillmentCopyEndColumnIndex - FulfillmentCopyStartColumnIndex + 1)
            .Where(sourceColumnIndex => sourceColumnIndex < sourceHeaders.Count)
            .Select(sourceColumnIndex => new CopyColumn(
                sourceColumnIndex,
                targetHeaderIndex.TryGetValue(sourceHeaders[sourceColumnIndex], out var targetColumnIndex)
                    ? targetColumnIndex
                    : -1,
                sourceHeaders[sourceColumnIndex]))
            .Where(column => column.TargetColumnIndex >= 0 && column.Header.Length > 0)
            .ToList();

        ClearWritableCells(targetSheet, targetHeaderRowIndex, copyColumns.Select(column => column.TargetColumnIndex));

        var formulaTemplateRow = targetSheet.GetRow(targetHeaderRowIndex + 1);
        var dataColumnIndexes = copyColumns.Select(column => column.TargetColumnIndex).ToHashSet();
        var nextTargetRowIndex = targetHeaderRowIndex + 1;
        for (var rowIndex = sourceHeaderRowIndex + 1; rowIndex <= sourceSheet.LastRowNum; rowIndex++)
        {
            var sourceRow = sourceSheet.GetRow(rowIndex);
            if (sourceRow is null)
            {
                continue;
            }

            var orderCell = sourceRow.GetCell(orderColumnIndex);
            var orderStatus = formatter.FormatCellValue(sourceRow.GetCell(statusColumnIndex));
            if (IsCellFontRed(workbook, orderCell)
                || orderStatus.Contains("cancelled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
            CopyFormulaTemplateCells(formulaTemplateRow, targetRow, dataColumnIndexes);
            foreach (var column in copyColumns)
            {
                CopyFormattedCellValue(
                    sourceRow.GetCell(column.SourceColumnIndex),
                    targetRow.CreateCell(column.TargetColumnIndex),
                    formatter,
                    cellStyleCache);
            }

            nextTargetRowIndex++;
        }

        return nextTargetRowIndex - targetHeaderRowIndex - 1;
    }

    private static int CopyOrderPaymentRowsToTemplate(
        IWorkbook workbook,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var sourceSheet = workbook.GetSheet(PaymentsSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentsSheetName}");
        var targetSheet = workbook.GetSheet(PaymentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentTemplateSheetName}");

        var sourceHeaderRowIndex = FindHeaderRow(sourceSheet, formatter);
        var targetHeaderRowIndex = FindHeaderRow(targetSheet, formatter);
        var sourceHeaderIndex = ReadHeaderIndex(sourceSheet, sourceHeaderRowIndex, formatter);
        var targetHeaderIndex = ReadHeaderIndex(targetSheet, targetHeaderRowIndex, formatter, preferLastDuplicate: true);
        var typeColumnIndex = ResolveRequiredColumn(sourceHeaderIndex, "type", PaymentsSheetName);

        var copyColumns = PaymentTemplateColumnMap
            .Select(mapping => new CopyColumn(
                ResolveRequiredColumn(sourceHeaderIndex, mapping.Key, PaymentsSheetName),
                ResolveRequiredColumn(targetHeaderIndex, mapping.Value, PaymentTemplateSheetName),
                mapping.Key))
            .ToList();

        ClearWritableCells(targetSheet, targetHeaderRowIndex, copyColumns.Select(column => column.TargetColumnIndex));

        var formulaTemplateRow = targetSheet.GetRow(targetHeaderRowIndex + 1);
        var dataColumnIndexes = copyColumns.Select(column => column.TargetColumnIndex).ToHashSet();
        var nextTargetRowIndex = targetHeaderRowIndex + 1;
        for (var rowIndex = sourceHeaderRowIndex + 1; rowIndex <= sourceSheet.LastRowNum; rowIndex++)
        {
            var sourceRow = sourceSheet.GetRow(rowIndex);
            if (sourceRow is null)
            {
                continue;
            }

            var type = formatter.FormatCellValue(sourceRow.GetCell(typeColumnIndex));
            if (!string.Equals(type.Trim(), "order", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
            CopyFormulaTemplateCells(formulaTemplateRow, targetRow, dataColumnIndexes);
            foreach (var column in copyColumns)
            {
                CopyFormattedCellValue(
                    sourceRow.GetCell(column.SourceColumnIndex),
                    targetRow.CreateCell(column.TargetColumnIndex),
                    formatter,
                    cellStyleCache);
            }

            nextTargetRowIndex++;
        }

        return nextTargetRowIndex - targetHeaderRowIndex - 1;
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

    private static string BuildReportFileName(DateOnly processingDate)
    {
        return $"{processingDate.Year}.{processingDate.Month}\u6708\u65e5\u62a5{processingDate.Day}\u65e5.xlsx";
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
        var table = ReadExcelTable(workbook, workbook.GetSheetAt(0), formatter);
        workbook.Close();
        return table;
    }

    private static TableData ReadExcelTable(IWorkbook workbook, string sheetName, DataFormatter formatter)
    {
        var sheet = FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException($"Sheet not found in source workbook: {sheetName}");

        return ReadExcelTable(workbook, sheet, formatter);
    }

    private static TableData ReadExcelTable(IWorkbook workbook, ISheet sheet, DataFormatter formatter)
    {
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        if (headerRowIndex < 0)
        {
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

        return new TableData(headers, rows);
    }

    private static int CopyTableDataToSheet(
        TableData sourceTable,
        ISheet targetSheet,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var targetHeaderRowIndex = FindHeaderRow(targetSheet, formatter);
        if (targetHeaderRowIndex < 0)
        {
            throw new InvalidOperationException($"No header row found in sheet: {targetSheet.SheetName}");
        }

        var targetHeaders = ReadRow(targetSheet.GetRow(targetHeaderRowIndex), formatter)
            .Select(DelimitedTableReader.CleanHeader)
            .ToList();
        var sourceIndex = BuildHeaderIndex(sourceTable.Headers);
        var writableColumns = targetHeaders
            .Select((header, columnIndex) => new CopyColumn(
                sourceIndex.TryGetValue(header, out var sourceColumnIndex) ? sourceColumnIndex : -1,
                columnIndex,
                header))
            .Where(column => column.Header.Length > 0 && column.SourceColumnIndex >= 0)
            .ToList();

        if (writableColumns.Count == 0)
        {
            throw new InvalidOperationException($"No matching columns for sheet: {targetSheet.SheetName}");
        }

        ClearWritableCells(targetSheet, targetHeaderRowIndex, writableColumns.Select(column => column.TargetColumnIndex));

        var nextTargetRowIndex = targetHeaderRowIndex + 1;
        foreach (var sourceRow in sourceTable.Rows)
        {
            var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
            foreach (var column in writableColumns)
            {
                var value = column.SourceColumnIndex < sourceRow.Count
                    ? sourceRow[column.SourceColumnIndex]
                    : string.Empty;

                SetCellValue(targetRow.CreateCell(column.TargetColumnIndex), value, cellStyleCache);
            }

            nextTargetRowIndex++;
        }

        return sourceTable.Rows.Count;
    }

    private static ISheet? FindSheet(IWorkbook workbook, string sheetName)
    {
        var sheet = workbook.GetSheet(sheetName);
        if (sheet is not null)
        {
            return sheet;
        }

        for (var i = 0; i < workbook.NumberOfSheets; i++)
        {
            sheet = workbook.GetSheetAt(i);
            if (string.Equals(sheet.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return sheet;
            }
        }

        return null;
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

    private static IReadOnlyList<string> ReadHeaders(ISheet sheet, int headerRowIndex, DataFormatter formatter)
    {
        if (headerRowIndex < 0)
        {
            throw new InvalidOperationException($"No header row found in sheet: {sheet.SheetName}");
        }

        return ReadRow(sheet.GetRow(headerRowIndex), formatter)
            .Select(DelimitedTableReader.CleanHeader)
            .ToList();
    }

    private static Dictionary<string, int> ReadHeaderIndex(
        ISheet sheet,
        int headerRowIndex,
        DataFormatter formatter,
        bool preferLastDuplicate = false)
    {
        return BuildHeaderIndex(ReadHeaders(sheet, headerRowIndex, formatter), preferLastDuplicate);
    }

    private static Dictionary<string, int> BuildHeaderIndex(IReadOnlyList<string> headers, bool preferLastDuplicate = false)
    {
        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Count; i++)
        {
            var header = DelimitedTableReader.CleanHeader(headers[i]);
            if (header.Length == 0)
            {
                continue;
            }

            if (preferLastDuplicate)
            {
                index[header] = i;
            }
            else if (!index.ContainsKey(header))
            {
                index.Add(header, i);
            }
        }

        return index;
    }

    private static int ResolveRequiredColumn(
        IReadOnlyDictionary<string, int> headerIndex,
        string header,
        string sheetName)
    {
        if (headerIndex.TryGetValue(header, out var columnIndex))
        {
            return columnIndex;
        }

        throw new InvalidOperationException($"Column '{header}' not found in sheet: {sheetName}");
    }

    private static int ResolveWaitingOrderColumnIndex(IReadOnlyDictionary<string, int> headerIndex)
    {
        foreach (var candidate in new[] { "\u8ba2\u5355\u53f7", AmazonOrderIdHeader, "order id", "order-id" })
        {
            if (headerIndex.TryGetValue(candidate, out var columnIndex))
            {
                return columnIndex;
            }
        }

        return -1;
    }

    private static string NormalizeOrderId(string? value)
    {
        return (value ?? string.Empty).Trim().Trim('\'');
    }

    private static void CopyFormattedCellValue(
        ICell? sourceCell,
        ICell targetCell,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        if (sourceCell is null)
        {
            targetCell.SetCellValue(string.Empty);
            return;
        }

        SetCellValue(targetCell, formatter.FormatCellValue(sourceCell), cellStyleCache);
    }

    private static void CopyFormulaTemplateCells(IRow? templateRow, IRow targetRow, ISet<int> dataColumnIndexes)
    {
        if (templateRow is null)
        {
            return;
        }

        var rowOffset = targetRow.RowNum - templateRow.RowNum;
        foreach (var templateCell in templateRow.Cells)
        {
            if (templateCell.CellType != CellType.Formula || dataColumnIndexes.Contains(templateCell.ColumnIndex))
            {
                continue;
            }

            var targetCell = targetRow.GetCell(templateCell.ColumnIndex) ?? targetRow.CreateCell(templateCell.ColumnIndex);
            targetCell.CellStyle = templateCell.CellStyle;
            targetCell.SetCellFormula(AdjustFormulaRows(templateCell.CellFormula, rowOffset));
        }
    }

    private static string AdjustFormulaRows(string formula, int rowOffset)
    {
        if (rowOffset == 0)
        {
            return formula;
        }

        return CellReferenceRegex.Replace(formula, match =>
        {
            if (match.Groups["absoluteRow"].Value == "$")
            {
                return match.Value;
            }

            var rowNumber = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture);
            return $"{match.Groups["column"].Value}{rowNumber + rowOffset}";
        });
    }

    private static bool IsCellFontRed(IWorkbook workbook, ICell? cell)
    {
        if (cell is null)
        {
            return false;
        }

        var font = workbook.GetFontAt(cell.CellStyle.FontIndex);
        return font.Color == IndexedColors.Red.Index;
    }

    private static void MarkWorkbookForFormulaRecalculation(IWorkbook workbook)
    {
        for (var i = 0; i < workbook.NumberOfSheets; i++)
        {
            workbook.GetSheetAt(i).ForceFormulaRecalculation = true;
        }

        if (workbook is XSSFWorkbook xssfWorkbook)
        {
            xssfWorkbook.SetForceFormulaRecalculation(true);
        }
    }

    private static void SetCellValue(ICell cell, string value, CellStyleCache styleCache)
    {
        if (TryParseDateCell(value, out var dateValue))
        {
            cell.SetCellValue(dateValue);
            cell.CellStyle = styleCache.Get("date");
            return;
        }

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

    private static bool TryParseDateCell(string value, out DateTime dateValue)
    {
        dateValue = default;
        var text = value.Trim();
        var match = ChineseMonthDateRegex.Match(text);
        if (match.Success
            && int.TryParse(match.Groups["year"].Value, out var year)
            && int.TryParse(match.Groups["month"].Value, out var month)
            && int.TryParse(match.Groups["day"].Value, out var day))
        {
            dateValue = new DateTime(year, month, day);
            return true;
        }

        return DateTime.TryParseExact(
            text,
            ["MMM d yyyy", "MMMM d yyyy", "M/d/yyyy", "yyyy-MM-dd"],
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out dateValue);
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

    private static void ClearWritableCells(ISheet sheet, int headerRowIndex, IEnumerable<int> writableColumnIndexes)
    {
        var columns = writableColumnIndexes.Distinct().ToList();
        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            foreach (var columnIndex in columns)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null)
                {
                    row.RemoveCell(cell);
                }
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

    private static void RemoveStaleCalculationMetadata(string xlsxPath)
    {
        using var archive = ZipFile.Open(xlsxPath, ZipArchiveMode.Update);
        archive.GetEntry("xl/calcChain.xml")?.Delete();
        RemoveWorkbookCalcChainRelationship(archive);
        RemoveCalcChainContentType(archive);
        MarkWorkbookXmlForFullRecalculation(archive);
    }

    private static void RemoveWorkbookCalcChainRelationship(ZipArchive archive)
    {
        const string relationshipsPath = "xl/_rels/workbook.xml.rels";
        var entry = archive.GetEntry(relationshipsPath);
        if (entry is null)
        {
            return;
        }

        var document = ReadXmlDocument(entry);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        document.Root?
            .Elements(ns + "Relationship")
            .Where(element => string.Equals((string?)element.Attribute("Target"), "calcChain.xml", StringComparison.OrdinalIgnoreCase))
            .Remove();

        ReplaceXmlEntry(archive, relationshipsPath, document);
    }

    private static void RemoveCalcChainContentType(ZipArchive archive)
    {
        const string contentTypesPath = "[Content_Types].xml";
        var entry = archive.GetEntry(contentTypesPath);
        if (entry is null)
        {
            return;
        }

        var document = ReadXmlDocument(entry);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        document.Root?
            .Elements(ns + "Override")
            .Where(element => string.Equals((string?)element.Attribute("PartName"), "/xl/calcChain.xml", StringComparison.OrdinalIgnoreCase))
            .Remove();

        ReplaceXmlEntry(archive, contentTypesPath, document);
    }

    private static void MarkWorkbookXmlForFullRecalculation(ZipArchive archive)
    {
        const string workbookPath = "xl/workbook.xml";
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        var document = ReadXmlDocument(entry);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var root = document.Root;
        if (root is null)
        {
            return;
        }

        var calcPr = root.Element(ns + "calcPr");
        if (calcPr is null)
        {
            calcPr = new XElement(ns + "calcPr");
            root.Add(calcPr);
        }

        calcPr.SetAttributeValue("calcMode", "auto");
        calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        calcPr.SetAttributeValue("forceFullCalc", "1");

        ReplaceXmlEntry(archive, workbookPath, document);
    }

    private static XDocument ReadXmlDocument(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static void ReplaceXmlEntry(ZipArchive archive, string path, XDocument document)
    {
        archive.GetEntry(path)?.Delete();
        var replacement = archive.CreateEntry(path);
        using var stream = replacement.Open();
        document.Save(stream, SaveOptions.DisableFormatting);
    }

    private readonly record struct ParsedNumericCell(double Value, string? FormatKey);

    private readonly record struct CopyColumn(int SourceColumnIndex, int TargetColumnIndex, string Header);

    private sealed class CellStyleCache
    {
        private readonly IWorkbook _workbook;
        private readonly IDataFormat _dataFormat;
        private readonly Dictionary<string, ICellStyle> _styles = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<short, ICellStyle> _redFontStyles = new();

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

        public ICellStyle GetRedFontStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_redFontStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
            }

            var font = _workbook.CreateFont();
            font.Color = IndexedColors.Red.Index;
            style.SetFont(font);
            _redFontStyles.Add(baseStyleIndex, style);
            return style;
        }

        private static string ResolveFormat(string key)
        {
            if (string.Equals(key, "percent", StringComparison.OrdinalIgnoreCase))
            {
                return "0.00%";
            }

            if (string.Equals(key, "date", StringComparison.OrdinalIgnoreCase))
            {
                return "m/d/yyyy";
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
