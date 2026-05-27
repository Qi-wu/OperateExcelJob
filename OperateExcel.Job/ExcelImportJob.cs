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
    private const string MappingSheetName = "\u6620\u5c04\u8868";
    private const string RmaSheetName = "RMA\u7533\u8bf7\u8868";
    private const string FulfillmentTemplateSheetName = "\u6a21\u7248F";
    private const string PaymentTemplateSheetName = "\u6a21\u677fP";
    private const string SummarySheetName = "\u6c47\u603b";
    private const string WaitingOrderFileName = "\u7b49\u5f85\u4e2d\u7684\u8ba2\u5355\u53f7.xlsx";
    private const string AmazonOrderIdHeader = "amazon-order-id";
    private const string OrderStatusHeader = "order-status";
    private const int FulfillmentCopyStartColumnIndex = 11; // Excel column L.
    private const int FulfillmentCopyEndColumnIndex = 19; // Excel column T.
    private const int FulfillmentSummaryStartRowIndex = 1334; // Excel row 1335.
    private const int FulfillmentSummaryStartColumnIndex = 2; // Excel column C.
    private const int FulfillmentSummaryColumnCount = 9;
    private static readonly short WaitingOrderFontColorIndex = IndexedColors.Red.Index;
    private const int PaymentStoreSummaryStartRowIndex = 1500; // Excel row 1501.
    private const int PaymentStoreSummaryStartColumnIndex = 2; // Excel column C.
    private const int PaymentStoreSummaryColumnCount = 5;
    private const int PaymentFirstDailySummaryStartColumnIndex = 0; // Excel column A.
    private const int PaymentFirstDailySummaryColumnCount = 10; // Excel columns A:J.
    private const int PaymentSecondDailySummaryStartColumnIndex = 14; // Excel column O.
    private const int PaymentSecondDailySummaryColumnCount = 19; // Excel columns O:AG.

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

    private static readonly IReadOnlyList<string> FulfillmentSummaryPeople =
    [
        "\u83ab\u7f8e\u7389",
        "\u4e01\u82b3\u82b3",
        "\u6b27\u9633\u535a\u6587",
        "\u5176\u4ed6",
        "\u8bb8\u6893\u6e1d",
        "\u8c2d\u71b9\u6770",
        "\u5510\u7487\u6dd1",
        "\u674e\u5c0f\u5a49",
        "\u6731\u5c0f\u71d5"
    ];

    private static readonly IReadOnlyList<string> FulfillmentSummaryStores =
    [
        "\u65e0\u5fe7\u65e0\u8651",
        "AN",
        "OYU",
        "DUX"
    ];

    private static readonly IReadOnlyList<string> FulfillmentSummaryHeaders =
    [
        "\u8ba2\u5355",
        "\u9500\u552e\u603b\u989d",
        "\u6ea2\u4ef7",
        "\u5e7f\u544a\u82b1\u8d39",
        "\u6700\u7ec8\u6ea2\u4ef7",
        "payments",
        "Refund",
        "payment\u6ea2\u4ef7"
    ];

    private static readonly IReadOnlyList<string> PaymentStoreSummaryHeaders =
    [
        "\u672c\u65e5\u9500\u552e\u989d",
        "\u672c\u65e5\u8ba2\u5355\u6570",
        "\u6ea2\u4ef7",
        "\u5e7f\u544a\u8d39"
    ];

    private static readonly IReadOnlyList<string> PaymentDailySummaryHeaders =
    [
        "\u65e5\u671f",
        "\u59d3\u540d",
        "\u8ba2\u5355",
        "\u9500\u552e\u603b\u989d",
        "\u6ea2\u4ef7",
        "\u5e7f\u544a\u82b1\u8d39",
        "\u6700\u7ec8\u6ea2\u4ef7",
        "payments",
        "Refund",
        "payment\u6ea2\u4ef7"
    ];

    private static readonly IReadOnlyDictionary<string, double> PaymentMonthlyBudgetByPerson =
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            ["\u83ab\u7f8e\u7389"] = 522425D,
            ["\u4e01\u82b3\u82b3"] = 428121D,
            ["\u6b27\u9633\u535a\u6587"] = 504162D,
            ["\u5176\u4ed6"] = 0D,
            ["\u8bb8\u6893\u6e1d"] = 0D,
            ["\u8c2d\u71b9\u6770"] = 221994D,
            ["\u5510\u7487\u6dd1"] = 145832D,
            ["\u674e\u5c0f\u5a49"] = 0D,
            ["\u6731\u5c0f\u71d5"] = 0D
        };

    private static readonly IReadOnlyDictionary<string, string> StoreAccountNames =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["\u65e0\u5fe7\u65e0\u8651"] = "\u65e0\u5fe7\u65e0\u8651",
            ["Oyumoents"] = "OYU",
            ["Yue an Company"] = "AN",
            ["Dux"] = "DUX",
            ["DUX"] = "DUX"
        };

    private static readonly IReadOnlyList<MappingSourceSheet> MappingSourceSheets =
    [
        new("\u65e0\u5fe7\u65e0\u8651", ["WUYOU", "\u65e0\u5fe7\u65e0\u8651"]),
        new("OYU", ["OYUMOENTS \u6620\u5c04", "OYUMOENTS\u6620\u5c04", "OYUMOENTS \u6620\u5c04\u8868", "OYUMOENTS\u6620\u5c04\u8868"]),
        new("AN", ["AN\u6620\u5c04\u8868", "AN \u6620\u5c04\u8868", "AN\u6620\u5c04", "AN \u6620\u5c04"]),
        new("DUX", ["DUX \u6620\u5c04", "DUX\u6620\u5c04", "DUX \u6620\u5c04\u8868", "DUX\u6620\u5c04\u8868"])
    ];

    private static readonly IReadOnlySet<string> MappingImportHeaders =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "\u5e73\u53f0SKU",
            "Asin",
            "B2B Item Code",
            "\u8fd0\u8425"
        };

    private static readonly IReadOnlyDictionary<string, string> RmaColumnMap =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["date/time"] = "\u65e5\u671f",
            ["order id"] = "order id",
            ["sku"] = "sku",
            ["quantity"] = "\u6570\u91cf",
            ["code"] = "Item Code",
            ["total"] = "total",
            ["\u5f52\u5c5e"] = "\u5f52\u5c5e",
            ["\u5382\u5bb6"] = "\u5382\u5bb6",
            ["\u8d26\u53f7"] = "\u5e97\u94fa"
        };

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
        ImportMappingFromFeishu(workbook, processingDate, formatter, cellStyleCache, messages);

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
            var accountColumnIndex = ResolveAccountColumnIndex(targetHeaders);

            var nextTargetRowIndex = targetHeaderRowIndex + 1;

            foreach (var storeDirectory in OrderStoreDirectories(storeDirectories))
            {
                var storeName = Path.GetFileName(storeDirectory);
                var accountName = ResolveAccountName(storeName);
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
                    ClearWritableCells(
                        targetSheet,
                        targetHeaderRowIndex,
                        writableColumns
                            .Select(column => column.TargetColumnIndex)
                            .Concat(accountColumnIndex >= 0 ? [accountColumnIndex] : []));
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

                    if (accountColumnIndex >= 0)
                    {
                        SetCellValue(targetRow.CreateCell(accountColumnIndex), accountName, cellStyleCache);
                    }

                    importedCounts[sheetName]++;
                    nextTargetRowIndex++;
                }

                messages.Add($"Imported {sourceTable.Rows.Count} rows from {sourceFile} to sheet {sheetName}.");
            }
        }

        UpdateRmaApplicationFromRefundPayments(workbook, processingDate, formatter, cellStyleCache, messages);

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

    private void ImportMappingFromFeishu(
        IWorkbook targetWorkbook,
        DateOnly processingDate,
        DataFormatter formatter,
        CellStyleCache cellStyleCache,
        ICollection<string> messages)
    {
        if (!_feishuOptions.Enabled)
        {
            messages.Add("Skipped Feishu mapping import: Feishu import is disabled.");
            _logger.LogWarning("Skipped Feishu mapping import because Feishu import is disabled.");
            return;
        }

        if (!_feishuClient.IsConfigured)
        {
            throw new InvalidOperationException("Feishu mapping import is enabled, but AppId/AppSecret or table configuration is incomplete.");
        }

        var attachmentBytes = _feishuClient
            .DownloadMappingAttachmentAsync(processingDate)
            .GetAwaiter()
            .GetResult();

        using var stream = new MemoryStream(attachmentBytes);
        var sourceWorkbook = WorkbookFactory.Create(stream);
        try
        {
            var targetSheet = FindSheet(targetWorkbook, _feishuOptions.MappingTargetSheetName)
                ?? FindSheet(targetWorkbook, MappingSheetName)
                ?? throw new InvalidOperationException($"Sheet not found in template: {_feishuOptions.MappingTargetSheetName}");

            var importedRows = CopyMappingSheetsToTarget(sourceWorkbook, targetSheet, formatter, cellStyleCache);
            messages.Add($"Imported {importedRows} rows from Feishu mapping attachment to sheet {targetSheet.SheetName}.");
        }
        finally
        {
            sourceWorkbook.Close();
        }
    }

    private void UpdateRmaApplicationFromRefundPayments(
        IWorkbook importedWorkbook,
        DateOnly processingDate,
        DataFormatter formatter,
        CellStyleCache cellStyleCache,
        ICollection<string> messages)
    {
        if (!_feishuOptions.Enabled)
        {
            messages.Add("Skipped Feishu RMA update: Feishu import is disabled.");
            _logger.LogWarning("Skipped Feishu RMA update because Feishu import is disabled.");
            return;
        }

        if (!_feishuClient.IsConfigured)
        {
            throw new InvalidOperationException("Feishu RMA update is enabled, but AppId/AppSecret or table configuration is incomplete.");
        }

        var refundRows = ReadRefundPaymentRows(importedWorkbook, formatter);
        if (refundRows.Count == 0)
        {
            messages.Add("Skipped Feishu RMA update: no Refund rows found in payments.");
            return;
        }

        var sourceRecordDate = processingDate.AddDays(-1);
        var rmaBytes = _feishuClient
            .DownloadRmaAttachmentAsync(sourceRecordDate)
            .GetAwaiter()
            .GetResult();

        using var inputStream = new MemoryStream(rmaBytes);
        var rmaWorkbook = WorkbookFactory.Create(inputStream);
        try
        {
            var targetSheetName = BuildRmaMonthSheetName(processingDate);
            var targetSheet = FindSheet(rmaWorkbook, targetSheetName) ?? rmaWorkbook.CreateSheet(targetSheetName);
            var appendedRows = AppendRefundRowsToRmaSheet(refundRows, targetSheet, formatter, cellStyleCache);

            using var outputStream = new MemoryStream();
            rmaWorkbook.Write(outputStream, leaveOpen: true);
            var outputBytes = outputStream.ToArray();
            var uploadFileDate = processingDate.AddDays(1);
            _feishuClient
                .UploadRmaAttachmentAsync(processingDate, $"{RmaSheetName}{uploadFileDate.Month}-{uploadFileDate.Day}.xlsx", outputBytes)
                .GetAwaiter()
                .GetResult();

            messages.Add($"Appended {appendedRows} Refund payment rows to RMA sheet {targetSheet.SheetName} and uploaded it to Feishu date {processingDate:yyyy-MM-dd}.");
        }
        finally
        {
            rmaWorkbook.Close();
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
        var generatedFulfillmentSummary = GenerateFulfillmentSummaryTemplates(workbook);
        var generatedPaymentSummaryBlocks = GeneratePaymentSummaryTemplates(workbook);
        var generatedDailySummaryRows = AppendDailySummaryTemplates(workbook, processingDate, generatedFulfillmentSummary, formatter);
        MarkWorkbookForFormulaRecalculation(workbook);

        messages.Add($"Highlighted {highlightedOrderIds} fulfillment rows from {WaitingOrderFileName}.");
        messages.Add($"Copied {copiedFulfillmentRows} fulfillment rows to sheet {FulfillmentTemplateSheetName}.");
        messages.Add($"Copied {copiedPaymentRows} payment rows to sheet {PaymentTemplateSheetName}.");
        messages.Add($"Generated {generatedFulfillmentSummary.BlockCount} fulfillment summary templates from row {FulfillmentSummaryStartRowIndex + 1}.");
        messages.Add($"Generated {generatedPaymentSummaryBlocks} payment summary templates from row {PaymentStoreSummaryStartRowIndex + 1}.");
        messages.Add($"Appended {generatedDailySummaryRows} rows to sheet {SummarySheetName}.");
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

    private static FulfillmentSummaryGenerationResult GenerateFulfillmentSummaryTemplates(IWorkbook workbook)
    {
        var sheet = workbook.GetSheet(FulfillmentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentTemplateSheetName}");

        ClearGeneratedFulfillmentSummaryArea(sheet);

        var storeHeaderStyleRow = sheet.GetRow(1302); // Existing DUX header row, Excel row 1303.
        var storePersonStyleRow = sheet.GetRow(1303); // Existing DUX first person row, Excel row 1304.
        var storeTotalStyleRow = sheet.GetRow(1314); // Existing DUX total row, Excel row 1315.
        var summaryTitleStyleRow = sheet.GetRow(1321); // Existing summary title row, Excel row 1322.
        var summaryHeaderStyleRow = sheet.GetRow(1322); // Existing summary header row, Excel row 1323.
        var summaryPersonStyleRow = sheet.GetRow(1323); // Existing summary first person row, Excel row 1324.
        var summaryTotalStyleRow = sheet.GetRow(1333); // Existing summary total row, Excel row 1334.

        var nextRowIndex = FulfillmentSummaryStartRowIndex;
        foreach (var store in FulfillmentSummaryStores)
        {
            nextRowIndex = CreateStoreSummaryTemplate(
                sheet,
                nextRowIndex,
                FulfillmentSummaryStartColumnIndex,
                store,
                storeHeaderStyleRow,
                storePersonStyleRow,
                storeTotalStyleRow);
            nextRowIndex++;
        }

        var allStoreTitleRowIndex = nextRowIndex;
        var allStoreHeaderRowIndex = allStoreTitleRowIndex + 1;
        var allStoreFirstPersonRowIndex = allStoreHeaderRowIndex + 1;

        CreateAllStoreSummaryTemplate(
            sheet,
            allStoreTitleRowIndex,
            FulfillmentSummaryStartColumnIndex,
            summaryTitleStyleRow,
            summaryHeaderStyleRow,
            summaryPersonStyleRow,
            summaryTotalStyleRow);

        return new FulfillmentSummaryGenerationResult(
            FulfillmentSummaryStores.Count + 1,
            allStoreHeaderRowIndex,
            allStoreFirstPersonRowIndex,
            allStoreFirstPersonRowIndex + FulfillmentSummaryPeople.Count - 1);
    }

    private static int CreateStoreSummaryTemplate(
        ISheet sheet,
        int headerRowIndex,
        int startColumnIndex,
        string storeName,
        IRow? headerStyleRow,
        IRow? personStyleRow,
        IRow? totalStyleRow)
    {
        var headerRow = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        CopyRowStyle(headerStyleRow, headerRow, FulfillmentSummaryStartColumnIndex, startColumnIndex, FulfillmentSummaryColumnCount);
        SetStringCell(headerRow, startColumnIndex, storeName);
        for (var i = 0; i < FulfillmentSummaryHeaders.Count; i++)
        {
            SetStringCell(headerRow, startColumnIndex + i + 1, FulfillmentSummaryHeaders[i]);
        }

        var firstPersonRowIndex = headerRowIndex + 1;
        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var rowIndex = firstPersonRowIndex + i;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            CopyRowStyle(personStyleRow, row, FulfillmentSummaryStartColumnIndex, startColumnIndex, FulfillmentSummaryColumnCount);
            SetStringCell(row, startColumnIndex, FulfillmentSummaryPeople[i]);
            SetStoreSummaryFormulas(row, startColumnIndex, headerRowIndex, storeName);
        }

        var totalRowIndex = firstPersonRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, FulfillmentSummaryStartColumnIndex, startColumnIndex, FulfillmentSummaryColumnCount);
        SetStringCell(totalRow, startColumnIndex, "\u5408\u8ba1");
        SetSumFormulas(totalRow, startColumnIndex, firstPersonRowIndex, totalRowIndex - 1);

        return totalRowIndex + 1;
    }

    private static void CreateAllStoreSummaryTemplate(
        ISheet sheet,
        int titleRowIndex,
        int startColumnIndex,
        IRow? titleStyleRow,
        IRow? headerStyleRow,
        IRow? personStyleRow,
        IRow? totalStyleRow)
    {
        var titleRow = sheet.GetRow(titleRowIndex) ?? sheet.CreateRow(titleRowIndex);
        CopyRowStyle(titleStyleRow, titleRow, 26, startColumnIndex, FulfillmentSummaryColumnCount); // AA:AI -> C:K.
        SetStringCell(titleRow, startColumnIndex, "\u6c47\u603b\u603b\u8ba1");

        var headerRowIndex = titleRowIndex + 1;
        var headerRow = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        CopyRowStyle(headerStyleRow, headerRow, 26, startColumnIndex, FulfillmentSummaryColumnCount);
        SetStringCell(headerRow, startColumnIndex, "\u5168\u90e8\u5e97\u94fa");
        for (var i = 0; i < FulfillmentSummaryHeaders.Count; i++)
        {
            SetStringCell(headerRow, startColumnIndex + i + 1, FulfillmentSummaryHeaders[i]);
        }

        var firstPersonRowIndex = headerRowIndex + 1;
        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var rowIndex = firstPersonRowIndex + i;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            CopyRowStyle(personStyleRow, row, 26, startColumnIndex, FulfillmentSummaryColumnCount);
            SetStringCell(row, startColumnIndex, FulfillmentSummaryPeople[i]);
            SetAllStoreSummaryFormulas(row, startColumnIndex, headerRowIndex);
        }

        var totalRowIndex = firstPersonRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, 26, startColumnIndex, FulfillmentSummaryColumnCount);
        SetStringCell(totalRow, startColumnIndex, "\u5408\u8ba1");
        SetSumFormulas(totalRow, startColumnIndex, firstPersonRowIndex, totalRowIndex - 1);
    }

    private static void SetStoreSummaryFormulas(IRow row, int startColumnIndex, int headerRowIndex, string storeName)
    {
        var rowNumber = row.RowNum + 1;
        var personRef = CellReference(startColumnIndex, rowNumber);
        var refundHeaderRef = CellReference(startColumnIndex + 7, headerRowIndex + 1);
        var storeCriterion = QuoteExcelString(storeName);

        SetFormulaCell(row, startColumnIndex + 1, $"SUMIFS(D:D,A:A,{personRef},O:O,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 2, $"SUMIFS(M:M,A:A,{personRef},O:O,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 3, $"SUMIFS(K:K,A:A,{personRef},O:O,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 4, $"SUMIFS('\u5e7f\u544a'!M:M,'\u5e7f\u544a'!N:N,{personRef},'\u5e7f\u544a'!P:P,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 5, $"{CellReference(startColumnIndex + 3, rowNumber)}-{CellReference(startColumnIndex + 4, rowNumber)}");
        SetFormulaCell(row, startColumnIndex + 6, $"SUMIFS('\u6a21\u677fP'!L:L,'\u6a21\u677fP'!A:A,{personRef},'\u6a21\u677fP'!N:N,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 7, $"SUMIFS(payments!AC:AC,payments!C:C,{refundHeaderRef},payments!AE:AE,{personRef},payments!AG:AG,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 8, $"SUMIFS('\u6a21\u677fP'!J:J,'\u6a21\u677fP'!A:A,{personRef},'\u6a21\u677fP'!N:N,{storeCriterion})");
    }

    private static void SetAllStoreSummaryFormulas(IRow row, int startColumnIndex, int headerRowIndex)
    {
        var rowNumber = row.RowNum + 1;
        var personRef = CellReference(startColumnIndex, rowNumber);
        var refundHeaderRef = CellReference(startColumnIndex + 7, headerRowIndex + 1);

        SetFormulaCell(row, startColumnIndex + 1, $"SUMIFS(D:D,A:A,{personRef})");
        SetFormulaCell(row, startColumnIndex + 2, $"SUMIFS(M:M,A:A,{personRef})");
        SetFormulaCell(row, startColumnIndex + 3, $"SUMIFS(K:K,A:A,{personRef})");
        SetFormulaCell(row, startColumnIndex + 4, $"SUMIFS('\u5e7f\u544a'!M:M,'\u5e7f\u544a'!N:N,{personRef})");
        SetFormulaCell(row, startColumnIndex + 5, $"{CellReference(startColumnIndex + 3, rowNumber)}-{CellReference(startColumnIndex + 4, rowNumber)}");
        SetFormulaCell(row, startColumnIndex + 6, $"SUMIFS('\u6a21\u677fP'!L:L,'\u6a21\u677fP'!A:A,{personRef})");
        SetFormulaCell(row, startColumnIndex + 7, $"SUMIFS(payments!AC:AC,payments!C:C,{refundHeaderRef},payments!AE:AE,{personRef})");
        SetFormulaCell(row, startColumnIndex + 8, $"SUMIFS('\u6a21\u677fP'!J:J,'\u6a21\u677fP'!A:A,{personRef})");
    }

    private static void SetSumFormulas(IRow row, int startColumnIndex, int firstRowIndex, int lastRowIndex)
    {
        var firstRowNumber = firstRowIndex + 1;
        var lastRowNumber = lastRowIndex + 1;
        for (var columnIndex = startColumnIndex + 1; columnIndex < startColumnIndex + FulfillmentSummaryColumnCount; columnIndex++)
        {
            var columnName = ColumnIndexToName(columnIndex);
            SetFormulaCell(row, columnIndex, $"SUM({columnName}{firstRowNumber}:{columnName}{lastRowNumber})");
        }
    }

    private static void SetColumnSumFormulas(IRow row, int startColumnIndex, int columnCount, int firstRowIndex, int lastRowIndex)
    {
        var firstRowNumber = firstRowIndex + 1;
        var lastRowNumber = lastRowIndex + 1;
        for (var columnIndex = startColumnIndex; columnIndex < startColumnIndex + columnCount; columnIndex++)
        {
            var columnName = ColumnIndexToName(columnIndex);
            SetFormulaCell(row, columnIndex, $"SUM({columnName}{firstRowNumber}:{columnName}{lastRowNumber})");
        }
    }

    private static void ClearGeneratedFulfillmentSummaryArea(ISheet sheet)
    {
        for (var rowIndex = FulfillmentSummaryStartRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            for (var columnIndex = FulfillmentSummaryStartColumnIndex;
                 columnIndex < FulfillmentSummaryStartColumnIndex + FulfillmentSummaryColumnCount;
                 columnIndex++)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null)
                {
                    row.RemoveCell(cell);
                }
            }
        }
    }

    private static int GeneratePaymentSummaryTemplates(IWorkbook workbook)
    {
        var sheet = workbook.GetSheet(PaymentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentTemplateSheetName}");

        ClearGeneratedPaymentSummaryArea(sheet);

        var headerStyleRow = sheet.GetRow(1486); // Existing DUX header row, Excel row 1487.
        var personStyleRow = sheet.GetRow(1487); // Existing DUX first person row, Excel row 1488.
        var totalStyleRow = sheet.GetRow(1498); // Existing DUX total row, Excel row 1499.

        var nextRowIndex = PaymentStoreSummaryStartRowIndex;
        foreach (var store in FulfillmentSummaryStores)
        {
            nextRowIndex = CreatePaymentStoreSummaryTemplate(
                sheet,
                nextRowIndex,
                PaymentStoreSummaryStartColumnIndex,
                store,
                headerStyleRow,
                personStyleRow,
                totalStyleRow);
            nextRowIndex++;
        }

        return FulfillmentSummaryStores.Count;
    }

    private static int CreatePaymentStoreSummaryTemplate(
        ISheet sheet,
        int headerRowIndex,
        int startColumnIndex,
        string storeName,
        IRow? headerStyleRow,
        IRow? personStyleRow,
        IRow? totalStyleRow)
    {
        var headerRow = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        CopyRowStyle(headerStyleRow, headerRow, 3, startColumnIndex, PaymentStoreSummaryColumnCount); // D:H -> C:G.
        SetStringCell(headerRow, startColumnIndex, storeName);
        for (var i = 0; i < PaymentStoreSummaryHeaders.Count; i++)
        {
            SetStringCell(headerRow, startColumnIndex + i + 1, PaymentStoreSummaryHeaders[i]);
        }

        var firstPersonRowIndex = headerRowIndex + 1;
        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var rowIndex = firstPersonRowIndex + i;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            CopyRowStyle(personStyleRow, row, 3, startColumnIndex, PaymentStoreSummaryColumnCount);
            SetStringCell(row, startColumnIndex, FulfillmentSummaryPeople[i]);
            SetPaymentStoreSummaryFormulas(row, startColumnIndex, storeName);
        }

        var totalRowIndex = firstPersonRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, 3, startColumnIndex, PaymentStoreSummaryColumnCount);
        SetStringCell(totalRow, startColumnIndex, "\u5408\u8ba1");
        SetColumnSumFormulas(totalRow, startColumnIndex + 1, PaymentStoreSummaryColumnCount - 1, firstPersonRowIndex, totalRowIndex - 1);

        return totalRowIndex + 1;
    }

    private static void SetPaymentStoreSummaryFormulas(IRow row, int startColumnIndex, string storeName)
    {
        var rowNumber = row.RowNum + 1;
        var personRef = CellReference(startColumnIndex, rowNumber);
        var storeCriterion = QuoteExcelString(storeName);

        SetFormulaCell(row, startColumnIndex + 1, $"SUMIFS(L:L,A:A,{personRef},N:N,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 2, $"SUMIFS(D:D,A:A,{personRef},N:N,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 3, $"SUMIFS(J:J,A:A,{personRef},N:N,{storeCriterion})");
        SetFormulaCell(row, startColumnIndex + 4, $"SUMIFS('\u5e7f\u544a'!M:M,'\u5e7f\u544a'!N:N,{personRef},'\u5e7f\u544a'!P:P,{storeCriterion})");
    }

    private static void ClearGeneratedPaymentSummaryArea(ISheet sheet)
    {
        for (var rowIndex = PaymentStoreSummaryStartRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            for (var columnIndex = PaymentStoreSummaryStartColumnIndex;
                 columnIndex < PaymentStoreSummaryStartColumnIndex + PaymentStoreSummaryColumnCount;
                 columnIndex++)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null)
                {
                    row.RemoveCell(cell);
                }
            }
        }
    }

    private static int AppendDailySummaryTemplates(
        IWorkbook workbook,
        DateOnly processingDate,
        FulfillmentSummaryGenerationResult fulfillmentSummary,
        DataFormatter formatter)
    {
        var fulfillmentSheet = workbook.GetSheet(FulfillmentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentTemplateSheetName}");
        var summarySheet = workbook.GetSheet(SummarySheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {SummarySheetName}");

        var firstSummaryRowIndex = FindNextAppendRowIndex(summarySheet, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryColumnCount);
        AppendFirstDailySummaryRows(
            summarySheet,
            fulfillmentSheet,
            processingDate,
            fulfillmentSummary,
            firstSummaryRowIndex);

        var secondSummaryFirstRowIndex = FindNextSecondDailySummaryRowIndex(summarySheet);
        AppendSecondDailySummaryRows(
            summarySheet,
            processingDate,
            secondSummaryFirstRowIndex,
            formatter);

        return FulfillmentSummaryPeople.Count + FulfillmentSummaryPeople.Count + 1;
    }

    private static void AppendFirstDailySummaryRows(
        ISheet summarySheet,
        ISheet fulfillmentSheet,
        DateOnly processingDate,
        FulfillmentSummaryGenerationResult fulfillmentSummary,
        int startRowIndex)
    {
        var styleRow = summarySheet.GetRow(Math.Max(1, startRowIndex - 1)) ?? summarySheet.GetRow(1);
        var sourceHeaderIndex = ReadHeaderIndexAtColumns(
            fulfillmentSheet.GetRow(fulfillmentSummary.AllStoreHeaderRowIndex),
            FulfillmentSummaryStartColumnIndex,
            FulfillmentSummaryColumnCount);
        var targetHeaderIndex = ReadHeaderIndexAtColumns(
            summarySheet.GetRow(0),
            PaymentFirstDailySummaryStartColumnIndex,
            PaymentFirstDailySummaryColumnCount);

        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var targetRowIndex = startRowIndex + i;
            var targetRow = summarySheet.GetRow(targetRowIndex) ?? summarySheet.CreateRow(targetRowIndex);
            CopyRowStyle(styleRow, targetRow, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryColumnCount);

            SetDateCell(targetRow, ResolveRequiredColumn(targetHeaderIndex, "\u65e5\u671f", SummarySheetName), processingDate);
            SetStringCell(targetRow, ResolveRequiredColumn(targetHeaderIndex, "\u59d3\u540d", SummarySheetName), FulfillmentSummaryPeople[i]);

            var sourceRow = fulfillmentSheet.GetRow(fulfillmentSummary.AllStoreFirstPersonRowIndex + i)
                ?? throw new InvalidOperationException($"Summary source row not found in sheet: {FulfillmentTemplateSheetName}");

            foreach (var header in PaymentDailySummaryHeaders.Skip(2))
            {
                if (!targetHeaderIndex.TryGetValue(header, out var targetColumnIndex)
                    || !sourceHeaderIndex.TryGetValue(header, out var sourceColumnIndex))
                {
                    continue;
                }

                SetFormulaCell(
                    targetRow,
                    targetColumnIndex,
                    $"'{FulfillmentTemplateSheetName}'!{CellReference(sourceColumnIndex, sourceRow.RowNum + 1)}");
            }
        }
    }

    private static void AppendSecondDailySummaryRows(
        ISheet summarySheet,
        DateOnly processingDate,
        int startRowIndex,
        DataFormatter formatter)
    {
        var templatePersonRow = FindSecondDailySummaryTemplatePersonRow(summarySheet, formatter)
            ?? throw new InvalidOperationException($"No second summary template row found in sheet: {SummarySheetName}");
        var lastTotalRowIndex = FindLastTotalRowIndex(summarySheet, PaymentSecondDailySummaryStartColumnIndex);

        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var personName = FulfillmentSummaryPeople[i];
            var targetRowIndex = startRowIndex + i;
            var targetRow = summarySheet.GetRow(targetRowIndex) ?? summarySheet.CreateRow(targetRowIndex);
            CopyRowStyle(templatePersonRow, targetRow, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryColumnCount);
            SetDateCell(targetRow, PaymentSecondDailySummaryStartColumnIndex, processingDate);
            SetStringCell(targetRow, PaymentSecondDailySummaryStartColumnIndex + 1, personName);
            SetSecondDailySummaryPersonFormulas(targetRow, personName);
        }

        var totalRowIndex = startRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = summarySheet.GetRow(totalRowIndex) ?? summarySheet.CreateRow(totalRowIndex);
        var templateTotalRow = summarySheet.GetRow(lastTotalRowIndex) ?? templatePersonRow;
        CopyRowStyle(templateTotalRow, totalRow, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryColumnCount);
        SetStringCell(totalRow, PaymentSecondDailySummaryStartColumnIndex, "\u5408\u8ba1");
        SetSecondDailySummaryTotalFormulas(totalRow, startRowIndex, totalRowIndex - 1);
    }

    private static void SetSecondDailySummaryPersonFormulas(IRow row, string personName)
    {
        var rowNumber = row.RowNum + 1;
        var personRef = CellReference(PaymentSecondDailySummaryStartColumnIndex + 1, rowNumber);

        SetFormulaCell(row, 16, $"SUMIFS($C:$C,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 17, $"SUMIFS($D:$D,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 18, $"SUMIFS($E:$E,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 19, $"SUMIFS($F:$F,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 21, $"IFERROR(T{rowNumber}/R{rowNumber},0)");
        SetFormulaCell(row, 22, $"IFERROR(U{rowNumber}/R{rowNumber},0)");
        SetFormulaCell(row, 23, $"SUMIFS($I:$I,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetFormulaCell(row, 25, $"SUMIFS($H:$H,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 26, $"SUMIFS($J:$J,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 27, $"AA{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
        SetFormulaCell(row, 29, $"AF{rowNumber}/30");
        SetFormulaCell(row, 30, $"X{rowNumber}+Z{rowNumber}+SUMIFS($AE:$AE,$O:$O,O{rowNumber}-1,$P:$P,{personRef})");
        SetNumericCell(row, 31, PaymentMonthlyBudgetByPerson.TryGetValue(personName, out var budget) ? budget : 0D);
        SetFormulaCell(row, 32, $"R{rowNumber}-AD{rowNumber}");
    }

    private static void SetSecondDailySummaryTotalFormulas(IRow row, int firstRowIndex, int lastRowIndex)
    {
        var rowNumber = row.RowNum + 1;
        var firstRowNumber = firstRowIndex + 1;
        var lastRowNumber = lastRowIndex + 1;

        foreach (var columnIndex in new[] { 16, 17, 18, 19, 23, 25, 26, 27, 29, 30, 31 })
        {
            var columnName = ColumnIndexToName(columnIndex);
            SetFormulaCell(row, columnIndex, $"SUM({columnName}{firstRowNumber}:{columnName}{lastRowNumber})");
        }

        SetFormulaCell(row, 21, $"T{rowNumber}/R{rowNumber}");
        SetFormulaCell(row, 22, $"U{rowNumber}/R{rowNumber}");
        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
    }

    private static IRow? FindSecondDailySummaryTemplatePersonRow(ISheet sheet, DataFormatter formatter)
    {
        for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var date = formatter.FormatCellValue(row.GetCell(PaymentSecondDailySummaryStartColumnIndex)).Trim();
            var name = formatter.FormatCellValue(row.GetCell(PaymentSecondDailySummaryStartColumnIndex + 1)).Trim();
            if (date.Length > 0 && name.Length > 0 && !string.Equals(date, "\u5408\u8ba1", StringComparison.OrdinalIgnoreCase))
            {
                return row;
            }
        }

        return null;
    }

    private static int FindNextSecondDailySummaryRowIndex(ISheet sheet)
    {
        return FindLastTotalRowIndex(sheet, PaymentSecondDailySummaryStartColumnIndex) + 1;
    }

    private static int FindLastTotalRowIndex(ISheet sheet, int columnIndex)
    {
        for (var rowIndex = sheet.LastRowNum; rowIndex >= 1; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            var cell = row?.GetCell(columnIndex);
            if (cell is not null && string.Equals(cell.ToString()?.Trim(), "\u5408\u8ba1", StringComparison.OrdinalIgnoreCase))
            {
                return rowIndex;
            }
        }

        throw new InvalidOperationException($"No total row found in sheet: {sheet.SheetName}");
    }

    private static int FindNextAppendRowIndex(ISheet sheet, int startColumnIndex, int columnCount)
    {
        for (var rowIndex = sheet.LastRowNum; rowIndex >= 0; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            for (var columnIndex = startColumnIndex; columnIndex < startColumnIndex + columnCount; columnIndex++)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null && !string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    return rowIndex + 1;
                }
            }
        }

        return 1;
    }

    private static Dictionary<string, int> ReadHeaderIndexAtColumns(IRow? row, int startColumnIndex, int columnCount)
    {
        if (row is null)
        {
            throw new InvalidOperationException("Header row not found.");
        }

        var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var columnIndex = startColumnIndex; columnIndex < startColumnIndex + columnCount; columnIndex++)
        {
            var header = row.GetCell(columnIndex)?.ToString()?.Trim() ?? string.Empty;
            if (header.Length > 0 && !index.ContainsKey(header))
            {
                index.Add(header, columnIndex);
            }
        }

        return index;
    }

    private static void SetDateCell(IRow row, int columnIndex, DateOnly date)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellValue(date.ToDateTime(TimeOnly.MinValue));
    }

    private static void CopyRowStyle(IRow? sourceRow, IRow targetRow, int sourceStartColumnIndex, int targetStartColumnIndex, int columnCount)
    {
        if (sourceRow is null)
        {
            return;
        }

        targetRow.Height = sourceRow.Height;
        for (var i = 0; i < columnCount; i++)
        {
            var sourceCell = sourceRow.GetCell(sourceStartColumnIndex + i);
            if (sourceCell is null)
            {
                continue;
            }

            var targetCell = targetRow.GetCell(targetStartColumnIndex + i) ?? targetRow.CreateCell(targetStartColumnIndex + i);
            targetCell.CellStyle = sourceCell.CellStyle;
        }
    }

    private static void SetStringCell(IRow row, int columnIndex, string value)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellValue(value);
    }

    private static void SetFormulaCell(IRow row, int columnIndex, string formula)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellFormula(formula);
    }

    private static void SetNumericCell(IRow row, int columnIndex, double value)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellValue(value);
    }

    private static string CellReference(int zeroBasedColumnIndex, int oneBasedRowNumber)
    {
        return $"{ColumnIndexToName(zeroBasedColumnIndex)}{oneBasedRowNumber}";
    }

    private static string ColumnIndexToName(int zeroBasedColumnIndex)
    {
        var columnNumber = zeroBasedColumnIndex + 1;
        var name = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            name = (char)('A' + columnNumber % 26) + name;
            columnNumber /= 26;
        }

        return name;
    }

    private static string QuoteExcelString(string value)
    {
        return $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
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

    private static string ResolveAccountName(string storeName)
    {
        return StoreAccountNames.TryGetValue(storeName, out var accountName)
            ? accountName
            : storeName;
    }

    private static int ResolveAccountColumnIndex(IReadOnlyList<string> targetHeaders)
    {
        for (var columnIndex = 0; columnIndex < targetHeaders.Count; columnIndex++)
        {
            if (string.Equals(targetHeaders[columnIndex], "\u8d26\u53f7", StringComparison.OrdinalIgnoreCase))
            {
                return columnIndex;
            }
        }

        return -1;
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
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        if (headerRowIndex < 0)
        {
            return new TableData([], []);
        }

        var headers = ReadRow(sheet.GetRow(headerRowIndex), formatter, evaluator: evaluator)
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

            var values = ReadRow(row, formatter, headers.Count, evaluator);
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

    private static int CopyMappingSheetsToTarget(
        IWorkbook sourceWorkbook,
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
        var targetHeaderIndex = BuildHeaderIndex(targetHeaders);
        var accountColumnIndex = ResolveRequiredColumn(targetHeaderIndex, "\u8d26\u53f7", targetSheet.SheetName);
        var copyColumns = MappingImportHeaders
            .Select(header => new CopyColumn(-1, ResolveRequiredColumn(targetHeaderIndex, header, targetSheet.SheetName), header))
            .ToList();

        ClearWritableCells(
            targetSheet,
            targetHeaderRowIndex,
            copyColumns.Select(column => column.TargetColumnIndex).Concat([accountColumnIndex]));

        var nextTargetRowIndex = targetHeaderRowIndex + 1;
        foreach (var sourceSheet in MappingSourceSheets)
        {
            var matchedSheet = FindSheet(sourceWorkbook, sourceSheet.SheetNameCandidates)
                ?? throw new InvalidOperationException(
                    $"Sheet not found in source workbook. Expected one of: {string.Join(", ", sourceSheet.SheetNameCandidates)}. Actual sheets: {string.Join(", ", ReadSheetNames(sourceWorkbook))}");

            var sourceTable = ReadExcelTable(sourceWorkbook, matchedSheet, formatter);
            var sourceHeaderIndex = BuildHeaderIndex(sourceTable.Headers);
            var resolvedColumns = copyColumns
                .Select(column => column with
                {
                    SourceColumnIndex = ResolveRequiredColumn(sourceHeaderIndex, column.Header, matchedSheet.SheetName)
                })
                .ToList();

            foreach (var sourceRow in sourceTable.Rows)
            {
                var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
                foreach (var column in resolvedColumns)
                {
                    var value = column.SourceColumnIndex < sourceRow.Count
                        ? sourceRow[column.SourceColumnIndex]
                        : string.Empty;

                    SetCellValue(targetRow.CreateCell(column.TargetColumnIndex), value, cellStyleCache);
                }

                SetCellValue(targetRow.CreateCell(accountColumnIndex), sourceSheet.AccountName, cellStyleCache);
                nextTargetRowIndex++;
            }
        }

        return nextTargetRowIndex - targetHeaderRowIndex - 1;
    }

    private static IReadOnlyList<IReadOnlyDictionary<string, string>> ReadRefundPaymentRows(
        IWorkbook workbook,
        DataFormatter formatter)
    {
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        var paymentSheet = workbook.GetSheet(PaymentsSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentsSheetName}");
        var headerRowIndex = FindHeaderRow(paymentSheet, formatter);
        var headerIndex = ReadHeaderIndex(paymentSheet, headerRowIndex, formatter);
        var typeColumnIndex = ResolveRequiredColumn(headerIndex, "type", PaymentsSheetName);
        var sourceColumns = RmaColumnMap
            .Select(mapping => new
            {
                SourceHeader = mapping.Key,
                TargetHeader = mapping.Value,
                SourceColumnIndex = ResolveRequiredColumn(headerIndex, mapping.Key, PaymentsSheetName)
            })
            .ToList();

        var refundRows = new List<IReadOnlyDictionary<string, string>>();
        for (var rowIndex = headerRowIndex + 1; rowIndex <= paymentSheet.LastRowNum; rowIndex++)
        {
            var row = paymentSheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var type = formatter.FormatCellValue(row.GetCell(typeColumnIndex)).Trim();
            if (!string.Equals(type, "Refund", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in sourceColumns)
            {
                values[column.TargetHeader] = FormatCellResolvedValue(
                    row.GetCell(column.SourceColumnIndex),
                    formatter,
                    evaluator).Trim();
            }

            refundRows.Add(values);
        }

        return refundRows;
    }

    private static int AppendRefundRowsToRmaSheet(
        IReadOnlyList<IReadOnlyDictionary<string, string>> refundRows,
        ISheet targetSheet,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var targetHeaders = RmaColumnMap.Values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var headerRowIndex = FindHeaderRow(targetSheet, formatter);
        IRow headerRow;
        if (headerRowIndex < 0)
        {
            headerRowIndex = 0;
            headerRow = targetSheet.GetRow(headerRowIndex) ?? targetSheet.CreateRow(headerRowIndex);
            for (var columnIndex = 0; columnIndex < targetHeaders.Count; columnIndex++)
            {
                headerRow.CreateCell(columnIndex).SetCellValue(targetHeaders[columnIndex]);
            }
        }
        else
        {
            headerRow = targetSheet.GetRow(headerRowIndex)
                ?? throw new InvalidOperationException($"Header row not found in sheet: {targetSheet.SheetName}");
        }

        var targetHeaderIndex = BuildHeaderIndex(ReadRow(headerRow, formatter));
        foreach (var header in targetHeaders)
        {
            if (targetHeaderIndex.ContainsKey(header))
            {
                continue;
            }

            var columnIndex = targetHeaderIndex.Count == 0
                ? 0
                : targetHeaderIndex.Values.Max() + 1;
            headerRow.CreateCell(columnIndex).SetCellValue(header);
            targetHeaderIndex[header] = columnIndex;
        }

        var nextTargetRowIndex = FindNextAppendRowIndex(targetSheet, headerRowIndex, formatter);
        foreach (var refundRow in refundRows)
        {
            var targetRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
            foreach (var header in targetHeaders)
            {
                var value = refundRow.TryGetValue(header, out var cellValue) ? cellValue : string.Empty;
                SetCellValue(targetRow.CreateCell(targetHeaderIndex[header]), value, cellStyleCache);
            }

            nextTargetRowIndex++;
        }

        return refundRows.Count;
    }

    private static int FindNextAppendRowIndex(ISheet sheet, int headerRowIndex, DataFormatter formatter)
    {
        for (var rowIndex = sheet.LastRowNum; rowIndex > headerRowIndex; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is not null && ReadRow(row, formatter).Any(value => !string.IsNullOrWhiteSpace(value)))
            {
                return rowIndex + 1;
            }
        }

        return headerRowIndex + 1;
    }

    private static string BuildRmaMonthSheetName(DateOnly date)
    {
        return $"{date.Year % 100}\u5e74{date.Month}\u6708";
    }

    private static ISheet? FindSheet(IWorkbook workbook, string sheetName)
    {
        return FindSheet(workbook, [sheetName]);
    }

    private static ISheet? FindSheet(IWorkbook workbook, IReadOnlyList<string> sheetNameCandidates)
    {
        foreach (var sheetName in sheetNameCandidates)
        {
            var sheet = workbook.GetSheet(sheetName);
            if (sheet is not null)
            {
                return sheet;
            }
        }

        foreach (var sheetName in sheetNameCandidates)
        {
            var normalizedCandidate = NormalizeSheetName(sheetName);
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if (string.Equals(NormalizeSheetName(sheet.SheetName), normalizedCandidate, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
        }

        return null;
    }

    private static IReadOnlyList<string> ReadSheetNames(IWorkbook workbook)
    {
        var sheetNames = new List<string>();
        for (var i = 0; i < workbook.NumberOfSheets; i++)
        {
            sheetNames.Add(workbook.GetSheetAt(i).SheetName);
        }

        return sheetNames;
    }

    private static string NormalizeSheetName(string sheetName)
    {
        var normalized = new string(sheetName.Where(character => !char.IsWhiteSpace(character)).ToArray());
        return normalized.EndsWith("\u8868", StringComparison.OrdinalIgnoreCase)
            ? normalized[..^1]
            : normalized;
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

    private static IReadOnlyList<string> ReadRow(
        IRow row,
        DataFormatter formatter,
        int? maxColumns = null,
        IFormulaEvaluator? evaluator = null)
    {
        var columnCount = maxColumns ?? row.LastCellNum;
        var values = new List<string>();
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            var cell = row.GetCell(columnIndex);
            values.Add(evaluator is not null && cell?.CellType == CellType.Formula
                ? formatter.FormatCellValue(cell, evaluator).Trim()
                : formatter.FormatCellValue(cell).Trim());
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

    private static string FormatCellResolvedValue(
        ICell? cell,
        DataFormatter formatter,
        IFormulaEvaluator evaluator)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        return cell.CellType == CellType.Formula
            ? formatter.FormatCellValue(cell, evaluator)
            : formatter.FormatCellValue(cell);
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
        if (font.Color == WaitingOrderFontColorIndex)
        {
            return true;
        }

        var color = font.GetType().GetMethod("GetXSSFColor", Type.EmptyTypes)?.Invoke(font, null)
            ?? font.GetType().GetProperty("XSSFColor")?.GetValue(font);

        return IsRedRgbColor(ReadColorBytes(color, "RGB"))
            || IsRedRgbColor(ReadColorBytes(color, "ARGB"));
    }

    private static byte[]? ReadColorBytes(object? color, string propertyName)
    {
        var value = color?.GetType().GetProperty(propertyName)?.GetValue(color);
        return value switch
        {
            byte[] bytes => bytes,
            sbyte[] signedBytes => signedBytes.Select(item => unchecked((byte)item)).ToArray(),
            _ => null
        };
    }

    private static bool IsRedRgbColor(byte[]? bytes)
    {
        if (bytes is null || bytes.Length < 3)
        {
            return false;
        }

        var offset = bytes.Length - 3;
        return bytes[offset] >= 240 && bytes[offset + 1] <= 32 && bytes[offset + 2] <= 32;
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

    private sealed record MappingSourceSheet(string AccountName, IReadOnlyList<string> SheetNameCandidates);

    private readonly record struct CopyColumn(int SourceColumnIndex, int TargetColumnIndex, string Header);

    private readonly record struct FulfillmentSummaryGenerationResult(
        int BlockCount,
        int AllStoreHeaderRowIndex,
        int AllStoreFirstPersonRowIndex,
        int AllStoreLastPersonRowIndex);

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
            font.Color = WaitingOrderFontColorIndex;
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
