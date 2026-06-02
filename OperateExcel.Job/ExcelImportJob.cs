using Microsoft.Extensions.Logging;
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
    private const string StorePersonRelationSheetName = "\u5e97\u94fa\u4eba\u5458\u5173\u7cfb";
    private const string WaitingOrderFileName = "\u7b49\u5f85\u4e2d\u7684\u8ba2\u5355\u53f7.xlsx";
    private const string B2BOlSourceFileNameKeyword = "\u4ea7\u54c1\u4fe1\u606f\u4e0b\u8f7d_\u57fa\u7840";
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
    private const int StoreDailySummaryStartColumnIndex = 14; // Excel column O.
    private const int StoreDailySummaryColumnCount = 16; // Excel columns O:AD.
    private const int DateColumnWidth = 12 * 256;

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
        "OYU",
        "AN",
        "DUX"
    ];

    private static readonly IReadOnlyDictionary<string, IReadOnlyList<string>> StorePeople =
        new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["\u65e0\u5fe7\u65e0\u8651"] =
            [
                "\u83ab\u7f8e\u7389",
                "\u4e01\u82b3\u82b3",
                "\u6b27\u9633\u535a\u6587",
                "\u5176\u4ed6",
                "\u8bb8\u6893\u6e1d",
                "\u8c2d\u71b9\u6770"
            ],
            ["AN"] =
            [
                "\u83ab\u7f8e\u7389",
                "\u5176\u4ed6",
                "\u8c2d\u71b9\u6770",
                "\u6731\u5c0f\u71d5"
            ],
            ["OYU"] =
            [
                "\u8bb8\u6893\u6e1d",
                "\u8c2d\u71b9\u6770",
                "\u674e\u5c0f\u5a49"
            ],
            ["DUX"] =
            [
                "\u4e01\u82b3\u82b3",
                "\u6b27\u9633\u535a\u6587",
                "\u8bb8\u6893\u6e1d",
                "\u8c2d\u71b9\u6770",
                "\u5510\u7487\u6dd1"
            ]
        };

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

    private static readonly IReadOnlyList<string> StoreDailySummaryHeaders =
    [
        "\u65e5\u671f",
        "\u59d3\u540d",
        "\u8ba2\u5355",
        "\u9500\u552e\u603b\u989d",
        "\u6ea2\u4ef7",
        "\u5e7f\u544a\u82b1\u8d39",
        "\u6700\u7ec8\u6ea2\u4ef7",
        "\u5e7f\u544a\u5360\u6bd4",
        "\u5229\u6da6\u7387",
        "Refund",
        "\u9000\u6b3e\u5360\u6bd4",
        "payments",
        "payment\u6ea2\u4ef7",
        "payment\u6700\u7ec8\u6ea2\u4ef7",
        "payment\u6bdb\u5229\u6bd4",
        "\u7d2f\u79efpayments"
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

        Directory.CreateDirectory(_options.OutputDirectory);
        var outputFilePath = Path.Combine(_options.OutputDirectory, BuildReportFileName(processingDate));
        CreateReportBaseFile(outputFilePath, processingDate, messages);

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

        ImportB2BOlFromLocalFile(workbook, sourceDirectory, formatter, cellStyleCache, messages);
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

                if (nextTargetRowIndex == targetHeaderRowIndex + 1)
                {
                    ClearDataRows(targetSheet, targetHeaderRowIndex + 1, targetSheet.LastRowNum, preserveFormulas: true);
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
                        value = NormalizeImportedCellValue(sheetName, column.Header, value);

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
        UploadDailyReportToFeishu(outputFilePath, processingDate, messages);

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

    private void CreateReportBaseFile(
        string outputFilePath,
        DateOnly processingDate,
        ICollection<string> messages)
    {
        if (!_feishuOptions.Enabled)
        {
            if (!File.Exists(_options.TemplateFilePath))
            {
                throw new FileNotFoundException("Template file not found.", _options.TemplateFilePath);
            }

            File.Copy(_options.TemplateFilePath, outputFilePath, overwrite: true);
            RemoveReadOnlyAttribute(outputFilePath);
            messages.Add($"Created daily report from local template: {_options.TemplateFilePath}.");
            return;
        }

        if (!_feishuClient.IsConfigured)
        {
            throw new InvalidOperationException("Feishu daily report template import is enabled, but AppId/AppSecret or table configuration is incomplete.");
        }

        var templateReportDate = processingDate.AddDays(-1);
        var reportBytes = _feishuClient
            .DownloadDailyReportAttachmentAsync(templateReportDate)
            .GetAwaiter()
            .GetResult();

        File.WriteAllBytes(outputFilePath, reportBytes);
        RemoveReadOnlyAttribute(outputFilePath);
        messages.Add($"Created daily report from Feishu {_feishuOptions.DailyReportAttachmentFieldName} attachment for {templateReportDate:yyyy-MM-dd}.");
    }

    private void UploadDailyReportToFeishu(
        string outputFilePath,
        DateOnly processingDate,
        ICollection<string> messages)
    {
        if (!_feishuOptions.Enabled)
        {
            messages.Add("Skipped Feishu daily report upload: Feishu import is disabled.");
            _logger.LogWarning("Skipped Feishu daily report upload because Feishu import is disabled.");
            return;
        }

        if (!_feishuClient.IsConfigured)
        {
            throw new InvalidOperationException("Feishu daily report upload is enabled, but AppId/AppSecret or table configuration is incomplete.");
        }

        var reportBytes = File.ReadAllBytes(outputFilePath);
        _feishuClient
            .UploadDailyReportAttachmentAndMarkCompletedAsync(processingDate, Path.GetFileName(outputFilePath), reportBytes)
            .GetAwaiter()
            .GetResult();

        _logger.LogInformation("Uploaded daily report to Feishu for {ProcessingDate}.", processingDate);
        messages.Add($"Uploaded daily report to Feishu date {processingDate:yyyy-MM-dd} and set {_feishuOptions.CompletionFieldName} to {_feishuOptions.CompletionValue}.");
    }

    private void ImportB2BOlFromLocalFile(
        IWorkbook targetWorkbook,
        string sourceDirectory,
        DataFormatter formatter,
        CellStyleCache cellStyleCache,
        ICollection<string> messages)
    {
        var sourceFile = FindB2BOlSourceFile(sourceDirectory);
        var sourceTable = ReadExcelTableWithTwoRowHeaders(sourceFile, formatter);
        var targetSheet = FindSheet(targetWorkbook, _feishuOptions.TargetSheetName)
            ?? throw new InvalidOperationException($"Sheet not found in template: {_feishuOptions.TargetSheetName}");

        var importedRows = CopyTableDataToSheet(sourceTable, targetSheet, formatter, cellStyleCache);
        messages.Add($"Imported {importedRows} rows from local file {Path.GetFileName(sourceFile)} to sheet {_feishuOptions.TargetSheetName}.");
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

        var sourceSheets = _feishuClient
            .ReadMappingSpreadsheetSheetsAsync()
            .GetAwaiter()
            .GetResult();

        var targetSheet = FindSheet(targetWorkbook, _feishuOptions.MappingTargetSheetName)
            ?? FindSheet(targetWorkbook, MappingSheetName)
            ?? throw new InvalidOperationException($"Sheet not found in template: {_feishuOptions.MappingTargetSheetName}");

        var importedRows = CopyMappingSpreadsheetSheetsToTarget(sourceSheets, targetSheet, formatter, cellStyleCache);
        messages.Add($"Imported {importedRows} rows from Feishu mapping spreadsheet to sheet {targetSheet.SheetName}.");
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
        UpsertStorePersonRelationSheet(workbook, cellStyleCache);
        if (processingDate.Day == 1)
        {
            ClearMonthlyDailySummarySheets(workbook);
            messages.Add($"Cleared monthly daily summary data for {processingDate:yyyy-MM-dd}.");
        }

        var generatedFulfillmentSummary = GenerateFulfillmentSummaryTemplates(workbook);
        var storeDailyMetrics = BuildStoreDailySummaryMetrics(workbook, formatter);
        var appendedStoreDailySummaryRows = AppendStoreDailySummarySheets(workbook, processingDate, storeDailyMetrics, cellStyleCache);
        var generatedPaymentSummaryBlocks = GeneratePaymentSummaryTemplates(workbook);
        var generatedDailySummaryRows = AppendDailySummaryTemplates(workbook, processingDate, generatedFulfillmentSummary, formatter, cellStyleCache);
        MarkWorkbookForFormulaRecalculation(workbook);

        messages.Add($"Highlighted {highlightedOrderIds} fulfillment rows from {WaitingOrderFileName}.");
        messages.Add($"Copied {copiedFulfillmentRows} fulfillment rows to sheet {FulfillmentTemplateSheetName}.");
        messages.Add($"Copied {copiedPaymentRows} payment rows to sheet {PaymentTemplateSheetName}.");
        messages.Add($"Updated sheet {StorePersonRelationSheetName}.");
        messages.Add($"Generated {generatedFulfillmentSummary.BlockCount} fulfillment summary templates from row {FulfillmentSummaryStartRowIndex + 1}.");
        messages.Add($"Appended {appendedStoreDailySummaryRows} rows to store daily summary sheets.");
        messages.Add($"Generated {generatedPaymentSummaryBlocks} payment summary templates from row {PaymentStoreSummaryStartRowIndex + 1}.");
        messages.Add($"Appended {generatedDailySummaryRows} rows to sheet {SummarySheetName}.");
    }

    private static void ClearMonthlyDailySummarySheets(IWorkbook workbook)
    {
        var summarySheet = workbook.GetSheet(SummarySheetName);
        if (summarySheet is not null)
        {
            ClearSheetDataRange(
                summarySheet,
                1,
                PaymentFirstDailySummaryStartColumnIndex,
                PaymentFirstDailySummaryColumnCount);
            ClearSheetDataRange(
                summarySheet,
                1,
                PaymentSecondDailySummaryStartColumnIndex,
                PaymentSecondDailySummaryColumnCount);
        }

        foreach (var store in FulfillmentSummaryStores)
        {
            var storeSheet = FindSheet(workbook, store);
            if (storeSheet is null)
            {
                continue;
            }

            ClearSheetDataRange(
                storeSheet,
                1,
                PaymentFirstDailySummaryStartColumnIndex,
                PaymentFirstDailySummaryColumnCount);
            ClearSheetDataRange(
                storeSheet,
                1,
                StoreDailySummaryStartColumnIndex,
                StoreDailySummaryColumnCount);
        }
    }

    private static void ClearSheetDataRange(
        ISheet sheet,
        int startRowIndex,
        int startColumnIndex,
        int columnCount)
    {
        if (sheet.LastRowNum < startRowIndex)
        {
            return;
        }

        var endColumnIndex = startColumnIndex + columnCount - 1;
        for (var rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            for (var columnIndex = startColumnIndex; columnIndex <= endColumnIndex; columnIndex++)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null)
                {
                    row.RemoveCell(cell);
                }
            }
        }

        RemoveMergedRegionsInRange(sheet, startRowIndex, sheet.LastRowNum, startColumnIndex, endColumnIndex);
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

        var formulaTemplateRow = targetSheet.GetRow(targetHeaderRowIndex + 1);
        var dataColumnIndexes = copyColumns.Select(column => column.TargetColumnIndex).ToHashSet();
        ClearTemplateDataRows(
            targetSheet,
            targetHeaderRowIndex + 1,
            FulfillmentSummaryStartRowIndex - 1,
            targetHeaderRowIndex + 1);

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

        var formulaTemplateRow = targetSheet.GetRow(targetHeaderRowIndex + 1);
        var dataColumnIndexes = copyColumns.Select(column => column.TargetColumnIndex).ToHashSet();
        ClearTemplateDataRows(
            targetSheet,
            targetHeaderRowIndex + 1,
            PaymentStoreSummaryStartRowIndex - 1,
            targetHeaderRowIndex + 1);

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
        ApplyGeneratedSummaryStyle(sheet, headerRow, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Header, GeneratedSummaryBlockKind.FulfillmentStore);
        SetStringCell(headerRow, startColumnIndex, storeName);
        for (var i = 0; i < FulfillmentSummaryHeaders.Count; i++)
        {
            SetStringCell(headerRow, startColumnIndex + i + 1, FulfillmentSummaryHeaders[i]);
        }

        var firstPersonRowIndex = headerRowIndex + 1;
        var storePeople = ResolveStorePeople(storeName);
        for (var i = 0; i < storePeople.Count; i++)
        {
            var rowIndex = firstPersonRowIndex + i;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            CopyRowStyle(personStyleRow, row, FulfillmentSummaryStartColumnIndex, startColumnIndex, FulfillmentSummaryColumnCount);
            ApplyGeneratedSummaryStyle(sheet, row, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Person, GeneratedSummaryBlockKind.FulfillmentStore);
            SetStringCell(row, startColumnIndex, storePeople[i]);
            SetStoreSummaryFormulas(row, startColumnIndex, headerRowIndex, storeName);
        }

        var totalRowIndex = firstPersonRowIndex + storePeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, FulfillmentSummaryStartColumnIndex, startColumnIndex, FulfillmentSummaryColumnCount);
        ApplyGeneratedSummaryStyle(sheet, totalRow, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Total, GeneratedSummaryBlockKind.FulfillmentStore);
        SetStringCell(totalRow, startColumnIndex, "\u5408\u8ba1");
        SetSumFormulas(totalRow, startColumnIndex, firstPersonRowIndex, totalRowIndex - 1);

        return totalRowIndex + 1;
    }

    private static IReadOnlyList<string> ResolveStorePeople(string storeName)
    {
        return StorePeople.TryGetValue(storeName, out var people)
            ? people
            : FulfillmentSummaryPeople;
    }

    private static void UpsertStorePersonRelationSheet(IWorkbook workbook, CellStyleCache cellStyleCache)
    {
        var sheet = workbook.GetSheet(StorePersonRelationSheetName)
            ?? workbook.CreateSheet(StorePersonRelationSheetName);

        for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            for (var columnIndex = 0; columnIndex < 2; columnIndex++)
            {
                var cell = row.GetCell(columnIndex);
                if (cell is not null)
                {
                    row.RemoveCell(cell);
                }
            }
        }

        var headerRow = sheet.GetRow(0) ?? sheet.CreateRow(0);
        SetCellValue(headerRow.CreateCell(0), "\u5e97\u94fa", cellStyleCache);
        SetCellValue(headerRow.CreateCell(1), "\u4eba\u5458", cellStyleCache);

        var nextRowIndex = 1;
        foreach (var store in FulfillmentSummaryStores)
        {
            foreach (var person in ResolveStorePeople(store))
            {
                var row = sheet.GetRow(nextRowIndex) ?? sheet.CreateRow(nextRowIndex);
                SetCellValue(row.CreateCell(0), store, cellStyleCache);
                SetCellValue(row.CreateCell(1), person, cellStyleCache);
                nextRowIndex++;
            }
        }
    }

    private static int AppendStoreDailySummarySheets(
        IWorkbook workbook,
        DateOnly processingDate,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, DailySummaryMetrics>> storeDailyMetrics,
        CellStyleCache cellStyleCache)
    {
        var appendedRows = 0;

        foreach (var store in FulfillmentSummaryStores)
        {
            var storePeople = ResolveStorePeople(store);
            var targetSheet = FindSheet(workbook, store) ?? workbook.CreateSheet(store);
            EnsureStoreDailySummaryHeaders(targetSheet, cellStyleCache);

            var nextTargetRowIndex = FindNextAppendRowIndex(
                targetSheet,
                PaymentFirstDailySummaryStartColumnIndex,
                PaymentFirstDailySummaryColumnCount);
            var nextSummaryRowIndex = FindNextStoreSummaryAppendRowIndex(targetSheet);

            foreach (var person in storePeople)
            {
                var detailRow = targetSheet.GetRow(nextTargetRowIndex) ?? targetSheet.CreateRow(nextTargetRowIndex);
                ApplyStoreDailySummaryDetailStyle(
                    detailRow,
                    PaymentFirstDailySummaryStartColumnIndex,
                    PaymentFirstDailySummaryColumnCount,
                    cellStyleCache,
                    hasBorder: false);

                SetDateCell(detailRow, 0, processingDate, cellStyleCache);
                SetStringCell(detailRow, 1, person);
                var metrics = storeDailyMetrics.TryGetValue(store, out var metricsByPerson)
                    && metricsByPerson.TryGetValue(person, out var personMetrics)
                        ? personMetrics
                        : new DailySummaryMetrics();
                SetDailySummaryMetricCells(detailRow, metrics);

                var summaryRow = targetSheet.GetRow(nextSummaryRowIndex) ?? targetSheet.CreateRow(nextSummaryRowIndex);
                ApplyStoreDailySummaryDetailStyle(
                    summaryRow,
                    StoreDailySummaryStartColumnIndex,
                    StoreDailySummaryColumnCount,
                    cellStyleCache,
                    hasBorder: true);
                SetDateCell(summaryRow, StoreDailySummaryStartColumnIndex, processingDate, cellStyleCache);
                SetStringCell(summaryRow, StoreDailySummaryStartColumnIndex + 1, person);
                SetStoreDailySummaryFormulaCells(summaryRow, cellStyleCache);

                nextTargetRowIndex++;
                nextSummaryRowIndex++;
                appendedRows++;
            }

            var totalRow = targetSheet.GetRow(nextSummaryRowIndex) ?? targetSheet.CreateRow(nextSummaryRowIndex);
            SetStringCell(totalRow, StoreDailySummaryStartColumnIndex, "\u5408\u8ba1");
            SetStoreDailySummaryTotalFormulas(
                totalRow,
                nextSummaryRowIndex - storePeople.Count,
                nextSummaryRowIndex - 1,
                cellStyleCache);
            ApplyStoreDailySummaryTotalStyle(targetSheet, totalRow, cellStyleCache);
            nextSummaryRowIndex++;
            ClearBlankStoreDailySummaryTemplateRows(targetSheet, nextSummaryRowIndex);
        }

        return appendedRows;
    }

    private static void EnsureStoreDailySummaryHeaders(ISheet sheet, CellStyleCache cellStyleCache)
    {
        EnsureDateColumnWidths(sheet);

        var headerRow = sheet.GetRow(0) ?? sheet.CreateRow(0);
        for (var i = 0; i < PaymentDailySummaryHeaders.Count; i++)
        {
            var cell = headerRow.GetCell(i);
            if (cell is null || string.IsNullOrWhiteSpace(cell.ToString()))
            {
                SetCellValue(headerRow.CreateCell(i), PaymentDailySummaryHeaders[i], cellStyleCache);
            }
        }

        for (var i = 0; i < StoreDailySummaryHeaders.Count; i++)
        {
            var columnIndex = StoreDailySummaryStartColumnIndex + i;
            var cell = headerRow.GetCell(columnIndex);
            if (cell is null || string.IsNullOrWhiteSpace(cell.ToString()))
            {
                SetCellValue(headerRow.CreateCell(columnIndex), StoreDailySummaryHeaders[i], cellStyleCache);
            }
        }
    }

    private static int FindNextStoreSummaryAppendRowIndex(ISheet sheet)
    {
        for (var rowIndex = sheet.LastRowNum; rowIndex >= 1; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var dateOrTotal = row.GetCell(StoreDailySummaryStartColumnIndex)?.ToString()?.Trim() ?? string.Empty;
            var person = row.GetCell(StoreDailySummaryStartColumnIndex + 1)?.ToString()?.Trim() ?? string.Empty;
            if (person.Length > 0)
            {
                return rowIndex + 1;
            }

            if (string.Equals(dateOrTotal, "\u5408\u8ba1", StringComparison.OrdinalIgnoreCase)
                && HasStoreSummaryPersonRowBefore(sheet, rowIndex))
            {
                return rowIndex + 1;
            }

            if (dateOrTotal.Length > 0 && !string.Equals(dateOrTotal, "\u5408\u8ba1", StringComparison.OrdinalIgnoreCase))
            {
                return rowIndex + 1;
            }
        }

        return 1;
    }

    private static bool HasStoreSummaryPersonRowBefore(ISheet sheet, int rowIndex)
    {
        for (var previousRowIndex = rowIndex - 1; previousRowIndex >= 1; previousRowIndex--)
        {
            var row = sheet.GetRow(previousRowIndex);
            if (row is null)
            {
                continue;
            }

            var dateOrTotal = row.GetCell(StoreDailySummaryStartColumnIndex)?.ToString()?.Trim() ?? string.Empty;
            var person = row.GetCell(StoreDailySummaryStartColumnIndex + 1)?.ToString()?.Trim() ?? string.Empty;
            if (person.Length > 0)
            {
                return true;
            }

            if (dateOrTotal.Length > 0)
            {
                return false;
            }
        }

        return false;
    }

    private static void SetStoreDailySummaryFormulaCells(IRow row, CellStyleCache cellStyleCache)
    {
        var rowNumber = row.RowNum + 1;
        var dateRef = CellReference(StoreDailySummaryStartColumnIndex, rowNumber);
        var personRef = CellReference(StoreDailySummaryStartColumnIndex + 1, rowNumber);

        SetFormulaCell(row, 16, $"SUMIFS($C:$C,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 17, $"SUMIFS($D:$D,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 18, $"SUMIFS($E:$E,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 19, $"SUMIFS($F:$F,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 21, $"IFERROR(T{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 21, cellStyleCache);
        SetFormulaCell(row, 22, $"IFERROR(U{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 22, cellStyleCache);
        SetFormulaCell(row, 23, $"SUMIFS($I:$I,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetPercentCellStyle(row, 24, cellStyleCache);
        SetFormulaCell(row, 25, $"SUMIFS($H:$H,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 26, $"SUMIFS($J:$J,$A:$A,{dateRef},$B:$B,{personRef})");
        SetFormulaCell(row, 27, $"AA{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
        SetPercentCellStyle(row, 28, cellStyleCache);
        SetFormulaCell(row, 29, $"X{rowNumber}+Z{rowNumber}");
    }

    private static void SetStoreDailySummaryTotalFormulas(
        IRow row,
        int firstPersonRowIndex,
        int lastPersonRowIndex,
        CellStyleCache cellStyleCache)
    {
        var rowNumber = row.RowNum + 1;
        var firstRowNumber = firstPersonRowIndex + 1;
        var lastRowNumber = lastPersonRowIndex + 1;

        foreach (var columnIndex in new[] { 16, 17, 18, 19, 23, 25, 26, 27, 29 })
        {
            var columnName = ColumnIndexToName(columnIndex);
            SetFormulaCell(row, columnIndex, $"SUM({columnName}{firstRowNumber}:{columnName}{lastRowNumber})");
        }

        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 21, $"IFERROR(T{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 21, cellStyleCache);
        SetFormulaCell(row, 22, $"IFERROR(U{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 22, cellStyleCache);
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetPercentCellStyle(row, 24, cellStyleCache);
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
        SetPercentCellStyle(row, 28, cellStyleCache);
    }

    private static void ApplyStoreDailySummaryDetailStyle(
        IRow row,
        int startColumnIndex,
        int columnCount,
        CellStyleCache cellStyleCache,
        bool hasBorder)
    {
        for (var columnIndex = startColumnIndex;
             columnIndex < startColumnIndex + columnCount;
             columnIndex++)
        {
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
            cell.CellStyle = cellStyleCache.GetStoreDailyDetailStyle(hasBorder);
        }
    }

    private static void ApplyStoreDailySummaryTotalStyle(ISheet sheet, IRow totalRow, CellStyleCache cellStyleCache)
    {
        var titleRow = sheet.GetRow(0);
        var titleCell = titleRow?.GetCell(StoreDailySummaryStartColumnIndex);
        var totalStyle = cellStyleCache.GetStoreDailyTotalStyle(titleCell?.CellStyle);

        for (var columnIndex = StoreDailySummaryStartColumnIndex;
             columnIndex < StoreDailySummaryStartColumnIndex + StoreDailySummaryColumnCount;
             columnIndex++)
        {
            var cell = totalRow.GetCell(columnIndex) ?? totalRow.CreateCell(columnIndex);
            cell.CellStyle = totalStyle;
        }

        ApplyStoreDailyPercentStyles(totalRow, cellStyleCache);
        MergeStoreDailySummaryTotalLabelCells(sheet, totalRow.RowNum);
    }

    private static void ApplyStoreDailyPercentStyles(IRow row, CellStyleCache cellStyleCache)
    {
        foreach (var columnIndex in new[] { 21, 22, 24, 28 })
        {
            SetPercentCellStyle(row, columnIndex, cellStyleCache);
        }
    }

    private static void MergeStoreDailySummaryTotalLabelCells(ISheet sheet, int rowIndex)
    {
        var region = new NPOI.SS.Util.CellRangeAddress(
            rowIndex,
            rowIndex,
            StoreDailySummaryStartColumnIndex,
            StoreDailySummaryStartColumnIndex + 1);

        RemoveOverlappingMergedRegions(sheet, region);
        sheet.AddMergedRegion(region);
    }

    private static void ClearBlankStoreDailySummaryTemplateRows(ISheet sheet, int startRowIndex)
    {
        for (var rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var date = row.GetCell(StoreDailySummaryStartColumnIndex)?.ToString()?.Trim() ?? string.Empty;
            var person = row.GetCell(StoreDailySummaryStartColumnIndex + 1)?.ToString()?.Trim() ?? string.Empty;
            if (person.Length > 0
                || (date.Length > 0 && !string.Equals(date, "\u5408\u8ba1", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            for (var columnIndex = StoreDailySummaryStartColumnIndex;
                 columnIndex < StoreDailySummaryStartColumnIndex + StoreDailySummaryColumnCount;
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
        ApplyGeneratedSummaryStyle(sheet, titleRow, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Title, GeneratedSummaryBlockKind.FulfillmentAllStore);
        SetStringCell(titleRow, startColumnIndex, "\u6c47\u603b\u603b\u8ba1");
        MergeGeneratedSummaryTitleCells(sheet, titleRowIndex, startColumnIndex, FulfillmentSummaryColumnCount);

        var headerRowIndex = titleRowIndex + 1;
        var headerRow = sheet.GetRow(headerRowIndex) ?? sheet.CreateRow(headerRowIndex);
        CopyRowStyle(headerStyleRow, headerRow, 26, startColumnIndex, FulfillmentSummaryColumnCount);
        ApplyGeneratedSummaryStyle(sheet, headerRow, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Header, GeneratedSummaryBlockKind.FulfillmentAllStore);
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
            ApplyGeneratedSummaryStyle(sheet, row, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Person, GeneratedSummaryBlockKind.FulfillmentAllStore);
            SetStringCell(row, startColumnIndex, FulfillmentSummaryPeople[i]);
            SetAllStoreSummaryFormulas(row, startColumnIndex, headerRowIndex);
        }

        var totalRowIndex = firstPersonRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, 26, startColumnIndex, FulfillmentSummaryColumnCount);
        ApplyGeneratedSummaryStyle(sheet, totalRow, startColumnIndex, FulfillmentSummaryColumnCount, GeneratedSummaryRowKind.Total, GeneratedSummaryBlockKind.FulfillmentAllStore);
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
        ApplyGeneratedSummaryStyle(sheet, headerRow, startColumnIndex, PaymentStoreSummaryColumnCount, GeneratedSummaryRowKind.Header, GeneratedSummaryBlockKind.PaymentStore);
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
            ApplyGeneratedSummaryStyle(sheet, row, startColumnIndex, PaymentStoreSummaryColumnCount, GeneratedSummaryRowKind.Person, GeneratedSummaryBlockKind.PaymentStore);
            SetStringCell(row, startColumnIndex, FulfillmentSummaryPeople[i]);
            SetPaymentStoreSummaryFormulas(row, startColumnIndex, storeName);
        }

        var totalRowIndex = firstPersonRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = sheet.GetRow(totalRowIndex) ?? sheet.CreateRow(totalRowIndex);
        CopyRowStyle(totalStyleRow, totalRow, 3, startColumnIndex, PaymentStoreSummaryColumnCount);
        ApplyGeneratedSummaryStyle(sheet, totalRow, startColumnIndex, PaymentStoreSummaryColumnCount, GeneratedSummaryRowKind.Total, GeneratedSummaryBlockKind.PaymentStore);
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

    private static void ApplyGeneratedSummaryStyle(
        ISheet sheet,
        IRow row,
        int startColumnIndex,
        int columnCount,
        GeneratedSummaryRowKind rowKind,
        GeneratedSummaryBlockKind blockKind)
    {
        var style = CreateGeneratedSummaryStyle(sheet.Workbook, rowKind, blockKind);
        for (var columnIndex = startColumnIndex; columnIndex < startColumnIndex + columnCount; columnIndex++)
        {
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
            cell.CellStyle = style;
        }
    }

    private static void MergeGeneratedSummaryTitleCells(
        ISheet sheet,
        int rowIndex,
        int startColumnIndex,
        int columnCount)
    {
        var region = new NPOI.SS.Util.CellRangeAddress(
            rowIndex,
            rowIndex,
            startColumnIndex,
            startColumnIndex + columnCount - 1);

        RemoveOverlappingMergedRegions(sheet, region);
        sheet.AddMergedRegion(region);
    }

    private static void RemoveOverlappingMergedRegions(ISheet sheet, NPOI.SS.Util.CellRangeAddress region)
    {
        for (var i = sheet.NumMergedRegions - 1; i >= 0; i--)
        {
            var existing = sheet.GetMergedRegion(i);
            if (existing.FirstRow <= region.LastRow
                && existing.LastRow >= region.FirstRow
                && existing.FirstColumn <= region.LastColumn
                && existing.LastColumn >= region.FirstColumn)
            {
                sheet.RemoveMergedRegion(i);
            }
        }
    }

    private static void EnsureDateColumnWidths(ISheet sheet)
    {
        sheet.SetColumnWidth(PaymentFirstDailySummaryStartColumnIndex, DateColumnWidth);
        sheet.SetColumnWidth(StoreDailySummaryStartColumnIndex, DateColumnWidth);
    }

    private static void RemoveMergedRegionsInRange(
        ISheet sheet,
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn)
    {
        for (var i = sheet.NumMergedRegions - 1; i >= 0; i--)
        {
            var existing = sheet.GetMergedRegion(i);
            if (existing.FirstRow >= firstRow
                && existing.LastRow <= lastRow
                && existing.FirstColumn >= firstColumn
                && existing.LastColumn <= lastColumn)
            {
                sheet.RemoveMergedRegion(i);
            }
        }
    }

    private static ICellStyle CreateGeneratedSummaryStyle(
        IWorkbook workbook,
        GeneratedSummaryRowKind rowKind,
        GeneratedSummaryBlockKind blockKind)
    {
        var style = workbook.CreateCellStyle();
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        style.BorderTop = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;

        var font = workbook.CreateFont();
        font.FontHeightInPoints = 11;

        switch (rowKind)
        {
            case GeneratedSummaryRowKind.Title:
                font.IsBold = true;
                SetGeneratedSummaryFillColor(workbook, style, blockKind);
                style.FillPattern = FillPattern.SolidForeground;
                break;
            case GeneratedSummaryRowKind.Header:
                font.IsBold = true;
                SetGeneratedSummaryFillColor(workbook, style, blockKind);
                style.FillPattern = FillPattern.SolidForeground;
                break;
            case GeneratedSummaryRowKind.Total:
                font.IsBold = true;
                SetGeneratedSummaryFillColor(workbook, style, blockKind);
                style.FillPattern = FillPattern.SolidForeground;
                break;
            default:
                style.FillForegroundColor = IndexedColors.White.Index;
                style.FillPattern = FillPattern.SolidForeground;
                break;
        }

        style.SetFont(font);
        return style;
    }

    private static void SetGeneratedSummaryFillColor(
        IWorkbook workbook,
        ICellStyle style,
        GeneratedSummaryBlockKind blockKind)
    {
        var rgb = blockKind switch
        {
            GeneratedSummaryBlockKind.PaymentStore => new byte[] { 0xFF, 0xF2, 0xCC },
            GeneratedSummaryBlockKind.FulfillmentStore => new byte[] { 0xC6, 0xE0, 0xB4 },
            GeneratedSummaryBlockKind.FulfillmentAllStore => new byte[] { 0x9B, 0xC2, 0xE6 },
            _ => new byte[] { 0xFF, 0xFF, 0xFF }
        };

        if (workbook is XSSFWorkbook && style is XSSFCellStyle xssfStyle)
        {
            xssfStyle.SetFillForegroundColor(new XSSFColor(rgb));
            return;
        }

        style.FillForegroundColor = blockKind switch
        {
            GeneratedSummaryBlockKind.PaymentStore => IndexedColors.LightYellow.Index,
            GeneratedSummaryBlockKind.FulfillmentStore => IndexedColors.LightGreen.Index,
            GeneratedSummaryBlockKind.FulfillmentAllStore => IndexedColors.LightCornflowerBlue.Index,
            _ => IndexedColors.White.Index
        };
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
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var summarySheet = workbook.GetSheet(SummarySheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {SummarySheetName}");
        var dailyMetrics = BuildDailySummaryMetrics(workbook, formatter);

        var firstSummaryRowIndex = FindNextAppendRowIndex(summarySheet, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryColumnCount);
        AppendFirstDailySummaryRows(
            summarySheet,
            processingDate,
            firstSummaryRowIndex,
            dailyMetrics,
            cellStyleCache);

        var secondSummaryFirstRowIndex = FindNextSecondDailySummaryRowIndex(summarySheet);
        AppendSecondDailySummaryRows(
            summarySheet,
            processingDate,
            secondSummaryFirstRowIndex,
            formatter,
            cellStyleCache);
        ApplySummarySheetCenterAlignment(summarySheet, cellStyleCache);

        return FulfillmentSummaryPeople.Count + FulfillmentSummaryPeople.Count + 1;
    }

    private static IReadOnlyDictionary<string, DailySummaryMetrics> BuildDailySummaryMetrics(
        IWorkbook workbook,
        DataFormatter formatter)
    {
        var metricsByPerson = FulfillmentSummaryPeople.ToDictionary(
            person => person,
            _ => new DailySummaryMetrics(),
            StringComparer.OrdinalIgnoreCase);

        var skuMappings = ReadSkuMappings(workbook, formatter);
        var b2bCosts = ReadB2BCosts(workbook, formatter);

        AddFulfillmentMetrics(workbook, formatter, skuMappings, b2bCosts, metricsByPerson);
        AddAdvertisingMetrics(workbook, formatter, skuMappings, metricsByPerson);
        AddPaymentMetrics(workbook, formatter, skuMappings, b2bCosts, metricsByPerson);

        return metricsByPerson;
    }

    private static IReadOnlyDictionary<string, IReadOnlyDictionary<string, DailySummaryMetrics>> BuildStoreDailySummaryMetrics(
        IWorkbook workbook,
        DataFormatter formatter)
    {
        var metricsByStore = FulfillmentSummaryStores.ToDictionary(
            store => store,
            store => (IDictionary<string, DailySummaryMetrics>)ResolveStorePeople(store).ToDictionary(
                person => person,
                _ => new DailySummaryMetrics(),
                StringComparer.OrdinalIgnoreCase),
            StringComparer.OrdinalIgnoreCase);

        var skuMappings = ReadSkuMappings(workbook, formatter);
        var b2bCosts = ReadB2BCosts(workbook, formatter);

        AddStoreFulfillmentMetrics(workbook, formatter, skuMappings, b2bCosts, metricsByStore);
        AddStoreAdvertisingMetrics(workbook, formatter, skuMappings, metricsByStore);
        AddStorePaymentMetrics(workbook, formatter, skuMappings, b2bCosts, metricsByStore);

        return metricsByStore.ToDictionary(
            store => store.Key,
            store => (IReadOnlyDictionary<string, DailySummaryMetrics>)new Dictionary<string, DailySummaryMetrics>(store.Value, StringComparer.OrdinalIgnoreCase),
            StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, SkuMapping> ReadSkuMappings(IWorkbook workbook, DataFormatter formatter)
    {
        var sheet = workbook.GetSheet(MappingSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {MappingSheetName}");
        var ownerNamesByCode = ReadSkuOwnerNames(workbook, formatter);
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var skuColumnIndex = ResolveRequiredColumn(headerIndex, "\u5e73\u53f0SKU", MappingSheetName);
        var itemCodeColumnIndex = ResolveRequiredColumn(headerIndex, "B2B Item Code", MappingSheetName);
        var ownerColumnIndex = ResolveRequiredColumn(headerIndex, "\u8fd0\u8425", MappingSheetName);
        var mappings = new Dictionary<string, SkuMapping>(StringComparer.OrdinalIgnoreCase);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (sku.Length == 0)
            {
                continue;
            }

            var ownerCode = formatter.FormatCellValue(row.GetCell(ownerColumnIndex)).Trim();
            var ownerName = ownerNamesByCode.TryGetValue(NormalizePersonName(ownerCode), out var resolvedOwnerName)
                ? resolvedOwnerName
                : ownerCode;
            mappings[sku] = new SkuMapping(
                formatter.FormatCellValue(row.GetCell(itemCodeColumnIndex)).Trim(),
                ownerName);
        }

        return mappings;
    }

    private static Dictionary<string, string> ReadSkuOwnerNames(IWorkbook workbook, DataFormatter formatter)
    {
        var sheet = workbook.GetSheet("\u0073\u006b\u0075\u5f52\u5c5e")
            ?? throw new InvalidOperationException("Sheet not found: sku\u5f52\u5c5e");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var ownerCodeColumnIndex = ResolveRequiredColumn(headerIndex, "\u8fd0\u8425", sheet.SheetName);
        var ownerNameColumnIndex = ResolveRequiredColumn(headerIndex, "\u59d3\u540d", sheet.SheetName);
        var ownerNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var ownerCode = NormalizePersonName(formatter.FormatCellValue(row.GetCell(ownerCodeColumnIndex)));
            var ownerName = formatter.FormatCellValue(row.GetCell(ownerNameColumnIndex)).Trim();
            if (ownerCode.Length > 0 && ownerName.Length > 0)
            {
                ownerNames[ownerCode] = ownerName;
            }
        }

        return ownerNames;
    }

    private static Dictionary<string, double> ReadB2BCosts(IWorkbook workbook, DataFormatter formatter)
    {
        var sheet = FindSheet(workbook, ["B2B\uff08ol)", "B2B(ol)", "B2BOL"])
            ?? throw new InvalidOperationException("Sheet not found: B2B\uff08ol)");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var itemCodeColumnIndex = ResolveRequiredColumn(headerIndex, "Item Code", sheet.SheetName);
        var costColumnIndex = headerIndex.TryGetValue("\u603b\u4ef7\uff08\u4e00\u4ef6\u4ee3\u53d1\uff09", out var namedCostColumnIndex)
            ? namedCostColumnIndex
            : 9;
        var costs = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var itemCode = formatter.FormatCellValue(row.GetCell(itemCodeColumnIndex)).Trim();
            if (itemCode.Length > 0)
            {
                costs[itemCode] = ReadNumericCell(row.GetCell(costColumnIndex), formatter);
            }
        }

        return costs;
    }

    private static void AddFulfillmentMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IReadOnlyDictionary<string, double> b2bCosts,
        IDictionary<string, DailySummaryMetrics> metricsByPerson)
    {
        var sheet = workbook.GetSheet(FulfillmentSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentSheetName}");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var orderColumnIndex = ResolveRequiredColumn(headerIndex, AmazonOrderIdHeader, FulfillmentSheetName);
        var statusColumnIndex = ResolveRequiredColumn(headerIndex, OrderStatusHeader, FulfillmentSheetName);
        var skuColumnIndex = ResolveRequiredColumn(headerIndex, "sku", FulfillmentSheetName);
        var quantityColumnIndex = ResolveRequiredColumn(headerIndex, "quantity", FulfillmentSheetName);
        var itemPriceColumnIndex = ResolveRequiredColumn(headerIndex, "item-price", FulfillmentSheetName);
        var shippingPriceColumnIndex = ResolveRequiredColumn(headerIndex, "shipping-price", FulfillmentSheetName);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null
                || IsCellFontRed(workbook, row.GetCell(orderColumnIndex))
                || formatter.FormatCellValue(row.GetCell(statusColumnIndex)).Contains("cancelled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var quantity = ReadNumericCell(row.GetCell(quantityColumnIndex), formatter);
            var salesTotal = ReadNumericCell(row.GetCell(itemPriceColumnIndex), formatter)
                + ReadNumericCell(row.GetCell(shippingPriceColumnIndex), formatter);
            var procurement = b2bCosts.TryGetValue(mapping.ItemCode, out var unitCost) ? unitCost * quantity : 0D;
            var premium = CalculatePremium(salesTotal, procurement, salesTotal * 0.05D);

            var metrics = metricsByPerson[owner];
            metricsByPerson[owner] = metrics with
            {
                Orders = metrics.Orders + quantity,
                SalesTotal = metrics.SalesTotal + salesTotal,
                Premium = metrics.Premium + premium
            };
        }
    }

    private static void AddAdvertisingMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IDictionary<string, DailySummaryMetrics> metricsByPerson)
    {
        var sheet = workbook.GetSheet(AdvertisingSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {AdvertisingSheetName}");
        const int skuColumnIndex = 6; // Excel column G.
        const int costColumnIndex = 12; // Excel column M.
        var headerRowIndex = FindHeaderRow(sheet, formatter);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var metrics = metricsByPerson[owner];
            metricsByPerson[owner] = metrics with
            {
                AdvertisingCost = metrics.AdvertisingCost + ReadNumericCell(row.GetCell(costColumnIndex), formatter)
            };
        }
    }

    private static void AddPaymentMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IReadOnlyDictionary<string, double> b2bCosts,
        IDictionary<string, DailySummaryMetrics> metricsByPerson)
    {
        var sheet = workbook.GetSheet(PaymentsSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentsSheetName}");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var typeColumnIndex = ResolveRequiredColumn(headerIndex, "type", PaymentsSheetName);
        var skuColumnIndex = ResolveRequiredColumn(headerIndex, "sku", PaymentsSheetName);
        var quantityColumnIndex = ResolveRequiredColumn(headerIndex, "quantity", PaymentsSheetName);
        var productSalesColumnIndex = ResolveRequiredColumn(headerIndex, "product sales", PaymentsSheetName);
        var shippingCreditsColumnIndex = ResolveRequiredColumn(headerIndex, "shipping credits", PaymentsSheetName);
        var totalColumnIndex = ResolveRequiredColumn(headerIndex, "total", PaymentsSheetName);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var type = formatter.FormatCellValue(row.GetCell(typeColumnIndex)).Trim();
            var metrics = metricsByPerson[owner];
            if (string.Equals(type, "order", StringComparison.OrdinalIgnoreCase))
            {
                var quantity = ReadNumericCell(row.GetCell(quantityColumnIndex), formatter);
                var salesTotal = ReadNumericCell(row.GetCell(productSalesColumnIndex), formatter)
                    + ReadNumericCell(row.GetCell(shippingCreditsColumnIndex), formatter);
                var procurement = b2bCosts.TryGetValue(mapping.ItemCode, out var unitCost) ? unitCost * quantity : 0D;
                metricsByPerson[owner] = metrics with
                {
                    Payments = metrics.Payments + salesTotal,
                    PaymentPremium = metrics.PaymentPremium + CalculatePremium(salesTotal, procurement, 0D)
                };
            }
            else if (string.Equals(type, "refund", StringComparison.OrdinalIgnoreCase))
            {
                metricsByPerson[owner] = metrics with
                {
                    Refund = metrics.Refund + ReadNumericCell(row.GetCell(totalColumnIndex), formatter)
                };
            }
        }
    }

    private static void AddStoreFulfillmentMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IReadOnlyDictionary<string, double> b2bCosts,
        IReadOnlyDictionary<string, IDictionary<string, DailySummaryMetrics>> metricsByStore)
    {
        var sheet = workbook.GetSheet(FulfillmentSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentSheetName}");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var orderColumnIndex = ResolveRequiredColumn(headerIndex, AmazonOrderIdHeader, FulfillmentSheetName);
        var statusColumnIndex = ResolveRequiredColumn(headerIndex, OrderStatusHeader, FulfillmentSheetName);
        var skuColumnIndex = ResolveRequiredColumn(headerIndex, "sku", FulfillmentSheetName);
        var quantityColumnIndex = ResolveRequiredColumn(headerIndex, "quantity", FulfillmentSheetName);
        var itemPriceColumnIndex = ResolveRequiredColumn(headerIndex, "item-price", FulfillmentSheetName);
        var shippingPriceColumnIndex = ResolveRequiredColumn(headerIndex, "shipping-price", FulfillmentSheetName);
        var storeColumnIndex = ResolveStoreColumnIndex(headerIndex, 14);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null
                || IsCellFontRed(workbook, row.GetCell(orderColumnIndex))
                || formatter.FormatCellValue(row.GetCell(statusColumnIndex)).Contains("cancelled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!TryResolveStore(formatter.FormatCellValue(row.GetCell(storeColumnIndex)), out var store)
                || !metricsByStore.TryGetValue(store, out var metricsByPerson))
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var quantity = ReadNumericCell(row.GetCell(quantityColumnIndex), formatter);
            var salesTotal = ReadNumericCell(row.GetCell(itemPriceColumnIndex), formatter)
                + ReadNumericCell(row.GetCell(shippingPriceColumnIndex), formatter);
            var procurement = b2bCosts.TryGetValue(mapping.ItemCode, out var unitCost) ? unitCost * quantity : 0D;
            var premium = CalculatePremium(salesTotal, procurement, salesTotal * 0.05D);

            var metrics = metricsByPerson[owner];
            metricsByPerson[owner] = metrics with
            {
                Orders = metrics.Orders + quantity,
                SalesTotal = metrics.SalesTotal + salesTotal,
                Premium = metrics.Premium + premium
            };
        }
    }

    private static void AddStoreAdvertisingMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IReadOnlyDictionary<string, IDictionary<string, DailySummaryMetrics>> metricsByStore)
    {
        var sheet = workbook.GetSheet(AdvertisingSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {AdvertisingSheetName}");
        const int skuColumnIndex = 6; // Excel column G.
        const int costColumnIndex = 12; // Excel column M.
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var storeColumnIndex = ResolveStoreColumnIndex(headerIndex, 15);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            if (!TryResolveStore(formatter.FormatCellValue(row.GetCell(storeColumnIndex)), out var store)
                || !metricsByStore.TryGetValue(store, out var metricsByPerson))
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var metrics = metricsByPerson[owner];
            metricsByPerson[owner] = metrics with
            {
                AdvertisingCost = metrics.AdvertisingCost + ReadNumericCell(row.GetCell(costColumnIndex), formatter)
            };
        }
    }

    private static void AddStorePaymentMetrics(
        IWorkbook workbook,
        DataFormatter formatter,
        IReadOnlyDictionary<string, SkuMapping> skuMappings,
        IReadOnlyDictionary<string, double> b2bCosts,
        IReadOnlyDictionary<string, IDictionary<string, DailySummaryMetrics>> metricsByStore)
    {
        var sheet = workbook.GetSheet(PaymentsSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentsSheetName}");
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        var headerIndex = ReadHeaderIndex(sheet, headerRowIndex, formatter);
        var typeColumnIndex = ResolveRequiredColumn(headerIndex, "type", PaymentsSheetName);
        var skuColumnIndex = ResolveRequiredColumn(headerIndex, "sku", PaymentsSheetName);
        var quantityColumnIndex = ResolveRequiredColumn(headerIndex, "quantity", PaymentsSheetName);
        var productSalesColumnIndex = ResolveRequiredColumn(headerIndex, "product sales", PaymentsSheetName);
        var shippingCreditsColumnIndex = ResolveRequiredColumn(headerIndex, "shipping credits", PaymentsSheetName);
        var totalColumnIndex = ResolveRequiredColumn(headerIndex, "total", PaymentsSheetName);
        var storeColumnIndex = ResolveStoreColumnIndex(headerIndex, 32);

        for (var rowIndex = headerRowIndex + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            if (!TryResolveStore(formatter.FormatCellValue(row.GetCell(storeColumnIndex)), out var store)
                || !metricsByStore.TryGetValue(store, out var metricsByPerson))
            {
                continue;
            }

            var sku = formatter.FormatCellValue(row.GetCell(skuColumnIndex)).Trim();
            if (!skuMappings.TryGetValue(sku, out var mapping))
            {
                continue;
            }

            var owner = ResolveSummaryPerson(mapping.Owner, metricsByPerson.Keys);
            if (owner is null)
            {
                continue;
            }

            var type = formatter.FormatCellValue(row.GetCell(typeColumnIndex)).Trim();
            var metrics = metricsByPerson[owner];
            if (string.Equals(type, "order", StringComparison.OrdinalIgnoreCase))
            {
                var quantity = ReadNumericCell(row.GetCell(quantityColumnIndex), formatter);
                var salesTotal = ReadNumericCell(row.GetCell(productSalesColumnIndex), formatter)
                    + ReadNumericCell(row.GetCell(shippingCreditsColumnIndex), formatter);
                var procurement = b2bCosts.TryGetValue(mapping.ItemCode, out var unitCost) ? unitCost * quantity : 0D;
                metricsByPerson[owner] = metrics with
                {
                    Payments = metrics.Payments + salesTotal,
                    PaymentPremium = metrics.PaymentPremium + CalculatePremium(salesTotal, procurement, 0D)
                };
            }
            else if (string.Equals(type, "refund", StringComparison.OrdinalIgnoreCase))
            {
                metricsByPerson[owner] = metrics with
                {
                    Refund = metrics.Refund + ReadNumericCell(row.GetCell(totalColumnIndex), formatter)
                };
            }
        }
    }

    private static int ResolveStoreColumnIndex(IReadOnlyDictionary<string, int> headerIndex, int fallbackColumnIndex)
    {
        return headerIndex.TryGetValue("\u8d26\u53f7", out var accountColumnIndex)
            ? accountColumnIndex
            : fallbackColumnIndex;
    }

    private static bool TryResolveStore(string rawStoreName, out string storeName)
    {
        var normalized = NormalizeSheetName(rawStoreName);
        foreach (var store in FulfillmentSummaryStores)
        {
            if (string.Equals(NormalizeSheetName(store), normalized, StringComparison.OrdinalIgnoreCase))
            {
                storeName = store;
                return true;
            }
        }

        storeName = string.Empty;
        return false;
    }

    private static string? ResolveSummaryPerson(
        string rawPersonName,
        IEnumerable<string> personNames)
    {
        var normalized = NormalizePersonName(rawPersonName);
        if (normalized.Length == 0)
        {
            return null;
        }

        foreach (var personName in personNames)
        {
            if (string.Equals(NormalizePersonName(personName), normalized, StringComparison.OrdinalIgnoreCase))
            {
                return personName;
            }
        }

        return null;
    }

    private static string NormalizePersonName(string value)
    {
        return new string((value ?? string.Empty)
            .Trim()
            .Where(character => !char.IsWhiteSpace(character))
            .ToArray());
    }

    private static double CalculatePremium(double salesTotal, double procurement, double fixedFee)
    {
        if (procurement == 0)
        {
            return 0;
        }
        else
        {
            var orderIncome = salesTotal > 200D
            ? (salesTotal - 200D) * 0.9D + 200D * 0.85D
            : salesTotal * 0.85D;

            return orderIncome - procurement - fixedFee;
        }
    }

    private static double ReadNumericCell(ICell? cell, DataFormatter formatter)
    {
        if (cell is null)
        {
            return 0D;
        }

        if (cell.CellType == CellType.Numeric)
        {
            return cell.NumericCellValue;
        }

        return TryParseNumericCell(formatter.FormatCellValue(cell), out var parsed)
            ? parsed.Value
            : 0D;
    }

    private static void AppendFirstDailySummaryRows(
        ISheet summarySheet,
        DateOnly processingDate,
        int startRowIndex,
        IReadOnlyDictionary<string, DailySummaryMetrics> dailyMetrics,
        CellStyleCache cellStyleCache)
    {
        var styleRow = summarySheet.GetRow(Math.Max(1, startRowIndex - 1)) ?? summarySheet.GetRow(1);
        var targetHeaderIndex = ReadHeaderIndexAtColumns(
            summarySheet.GetRow(0),
            PaymentFirstDailySummaryStartColumnIndex,
            PaymentFirstDailySummaryColumnCount);

        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var personName = FulfillmentSummaryPeople[i];
            var targetRowIndex = startRowIndex + i;
            var targetRow = summarySheet.GetRow(targetRowIndex) ?? summarySheet.CreateRow(targetRowIndex);
            CopyRowStyle(styleRow, targetRow, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryStartColumnIndex, PaymentFirstDailySummaryColumnCount);

            SetDateCell(targetRow, ResolveRequiredColumn(targetHeaderIndex, "\u65e5\u671f", SummarySheetName), processingDate, cellStyleCache);
            SetStringCell(targetRow, ResolveRequiredColumn(targetHeaderIndex, "\u59d3\u540d", SummarySheetName), personName);

            dailyMetrics.TryGetValue(personName, out var metrics);
            SetDailySummaryMetricCells(targetRow, targetHeaderIndex, metrics, SummarySheetName);
        }
    }

    private static void SetDailySummaryMetricCells(IRow row, DailySummaryMetrics metrics)
    {
        SetNumericCell(row, 2, metrics.Orders);
        SetNumericCell(row, 3, metrics.SalesTotal);
        SetNumericCell(row, 4, metrics.Premium);
        SetNumericCell(row, 5, metrics.AdvertisingCost);
        SetNumericCell(row, 6, metrics.Premium - metrics.AdvertisingCost);
        SetNumericCell(row, 7, metrics.Payments);
        SetNumericCell(row, 8, metrics.Refund);
        SetNumericCell(row, 9, metrics.PaymentPremium);
    }

    private static void SetDailySummaryMetricCells(
        IRow row,
        IReadOnlyDictionary<string, int> targetHeaderIndex,
        DailySummaryMetrics metrics,
        string sheetName)
    {
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "\u8ba2\u5355", sheetName), metrics.Orders);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "\u9500\u552e\u603b\u989d", sheetName), metrics.SalesTotal);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "\u6ea2\u4ef7", sheetName), metrics.Premium);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "\u5e7f\u544a\u82b1\u8d39", sheetName), metrics.AdvertisingCost);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "\u6700\u7ec8\u6ea2\u4ef7", sheetName), metrics.Premium - metrics.AdvertisingCost);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "payments", sheetName), metrics.Payments);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "Refund", sheetName), metrics.Refund);
        SetNumericCell(row, ResolveRequiredColumn(targetHeaderIndex, "payment\u6ea2\u4ef7", sheetName), metrics.PaymentPremium);
    }

    private static void AppendSecondDailySummaryRows(
        ISheet summarySheet,
        DateOnly processingDate,
        int startRowIndex,
        DataFormatter formatter,
        CellStyleCache cellStyleCache)
    {
        var templatePersonRow = FindSecondDailySummaryTemplatePersonRow(summarySheet, formatter)
            ?? summarySheet.GetRow(Math.Max(1, startRowIndex - 1))
            ?? summarySheet.GetRow(1)
            ?? summarySheet.GetRow(0);
        var lastTotalRowIndex = TryFindLastTotalRowIndex(summarySheet, PaymentSecondDailySummaryStartColumnIndex);

        for (var i = 0; i < FulfillmentSummaryPeople.Count; i++)
        {
            var personName = FulfillmentSummaryPeople[i];
            var targetRowIndex = startRowIndex + i;
            var targetRow = summarySheet.GetRow(targetRowIndex) ?? summarySheet.CreateRow(targetRowIndex);
            CopyRowStyle(templatePersonRow, targetRow, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryColumnCount);
            SetDateCell(targetRow, PaymentSecondDailySummaryStartColumnIndex, processingDate, cellStyleCache);
            SetStringCell(targetRow, PaymentSecondDailySummaryStartColumnIndex + 1, personName);
            SetSecondDailySummaryPersonFormulas(targetRow, personName, cellStyleCache);
            ApplySecondDailySummaryDetailStyle(targetRow, cellStyleCache);
        }

        var totalRowIndex = startRowIndex + FulfillmentSummaryPeople.Count;
        var totalRow = summarySheet.GetRow(totalRowIndex) ?? summarySheet.CreateRow(totalRowIndex);
        var templateTotalRow = lastTotalRowIndex >= 0
            ? summarySheet.GetRow(lastTotalRowIndex) ?? templatePersonRow
            : templatePersonRow;
        CopyRowStyle(templateTotalRow, totalRow, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex, PaymentSecondDailySummaryColumnCount);
        SetStringCell(totalRow, PaymentSecondDailySummaryStartColumnIndex, "\u5408\u8ba1");
        SetSecondDailySummaryTotalFormulas(totalRow, startRowIndex, totalRowIndex - 1, cellStyleCache);
        ApplySecondDailySummaryTotalStyle(summarySheet, totalRow, cellStyleCache);
    }

    private static void SetSecondDailySummaryPersonFormulas(IRow row, string personName, CellStyleCache cellStyleCache)
    {
        var rowNumber = row.RowNum + 1;
        var personRef = CellReference(PaymentSecondDailySummaryStartColumnIndex + 1, rowNumber);

        SetFormulaCell(row, 16, $"SUMIFS($C:$C,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 17, $"SUMIFS($D:$D,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 18, $"SUMIFS($E:$E,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 19, $"SUMIFS($F:$F,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 21, $"IFERROR(T{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 21, cellStyleCache);
        SetFormulaCell(row, 22, $"IFERROR(U{rowNumber}/R{rowNumber},0)");
        SetPercentCellStyle(row, 22, cellStyleCache);
        SetFormulaCell(row, 23, $"SUMIFS($I:$I,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetPercentCellStyle(row, 24, cellStyleCache);
        SetFormulaCell(row, 25, $"SUMIFS($H:$H,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 26, $"SUMIFS($J:$J,$A:$A,O{rowNumber},$B:$B,{personRef})");
        SetFormulaCell(row, 27, $"AA{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
        SetPercentCellStyle(row, 28, cellStyleCache);
        SetFormulaCell(row, 29, $"AF{rowNumber}/30");
        SetFormulaCell(row, 30, $"X{rowNumber}+Z{rowNumber}+SUMIFS($AE:$AE,$O:$O,O{rowNumber}-1,$P:$P,{personRef})");
        SetNumericCell(row, 31, PaymentMonthlyBudgetByPerson.TryGetValue(personName, out var budget) ? budget : 0D);
        SetFormulaCell(row, 32, $"R{rowNumber}-AD{rowNumber}");
    }

    private static void SetSecondDailySummaryTotalFormulas(
        IRow row,
        int firstRowIndex,
        int lastRowIndex,
        CellStyleCache cellStyleCache)
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
        SetPercentCellStyle(row, 21, cellStyleCache);
        SetFormulaCell(row, 22, $"U{rowNumber}/R{rowNumber}");
        SetPercentCellStyle(row, 22, cellStyleCache);
        SetFormulaCell(row, 20, $"S{rowNumber}-T{rowNumber}");
        SetFormulaCell(row, 24, $"IFERROR(X{rowNumber}/(Z{rowNumber}+X{rowNumber}),0)");
        SetPercentCellStyle(row, 24, cellStyleCache);
        SetFormulaCell(row, 28, $"IFERROR((AA{rowNumber}-T{rowNumber})/Z{rowNumber},0)");
        SetPercentCellStyle(row, 28, cellStyleCache);
    }

    private static void ApplySummarySheetCenterAlignment(ISheet summarySheet, CellStyleCache cellStyleCache)
    {
        for (var rowIndex = 0; rowIndex <= summarySheet.LastRowNum; rowIndex++)
        {
            var row = summarySheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            foreach (var cell in row.Cells)
            {
                cell.CellStyle = cellStyleCache.GetCenteredStyle(cell.CellStyle);
            }
        }
    }

    private static void ApplySecondDailySummaryTotalStyle(
        ISheet summarySheet,
        IRow totalRow,
        CellStyleCache cellStyleCache)
    {
        var headerCell = summarySheet.GetRow(0)?.GetCell(PaymentSecondDailySummaryStartColumnIndex);
        var totalStyle = cellStyleCache.GetStoreDailyTotalStyle(headerCell?.CellStyle);
        for (var columnIndex = PaymentSecondDailySummaryStartColumnIndex;
             columnIndex < PaymentSecondDailySummaryStartColumnIndex + PaymentSecondDailySummaryColumnCount;
             columnIndex++)
        {
            var cell = totalRow.GetCell(columnIndex) ?? totalRow.CreateCell(columnIndex);
            cell.CellStyle = totalStyle;
        }

        ApplyStoreDailyPercentStyles(totalRow, cellStyleCache);
        MergeSecondDailySummaryTotalLabelCells(summarySheet, totalRow.RowNum);
    }

    private static void ApplySecondDailySummaryDetailStyle(IRow row, CellStyleCache cellStyleCache)
    {
        for (var columnIndex = PaymentSecondDailySummaryStartColumnIndex; columnIndex <= 31; columnIndex++)
        {
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
            cell.CellStyle = cellStyleCache.GetSummaryDetailBorderStyle(cell.CellStyle);
        }
    }

    private static void MergeSecondDailySummaryTotalLabelCells(ISheet sheet, int rowIndex)
    {
        var region = new NPOI.SS.Util.CellRangeAddress(
            rowIndex,
            rowIndex,
            PaymentSecondDailySummaryStartColumnIndex,
            PaymentSecondDailySummaryStartColumnIndex + 1);

        RemoveOverlappingMergedRegions(sheet, region);
        sheet.AddMergedRegion(region);
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
        var lastTotalRowIndex = TryFindLastTotalRowIndex(sheet, PaymentSecondDailySummaryStartColumnIndex);
        return lastTotalRowIndex >= 0 ? lastTotalRowIndex + 1 : 1;
    }

    private static int FindLastTotalRowIndex(ISheet sheet, int columnIndex)
    {
        var rowIndex = TryFindLastTotalRowIndex(sheet, columnIndex);
        if (rowIndex >= 0)
        {
            return rowIndex;
        }

        throw new InvalidOperationException($"No total row found in sheet: {sheet.SheetName}");
    }

    private static int TryFindLastTotalRowIndex(ISheet sheet, int columnIndex)
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

        return -1;
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

    private static void SetDateCell(IRow row, int columnIndex, DateOnly date, CellStyleCache cellStyleCache)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.SetCellValue(date.ToDateTime(TimeOnly.MinValue));
        cell.CellStyle = cellStyleCache.GetDateStyle(cell.CellStyle);
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

    private static void SetPercentCellStyle(IRow row, int columnIndex, CellStyleCache cellStyleCache)
    {
        var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        cell.CellStyle = cellStyleCache.GetPercentStyle(cell.CellStyle);
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

    private static string FindB2BOlSourceFile(string sourceDirectory)
    {
        var sourceFile = Directory
            .EnumerateFiles(sourceDirectory, "*", SearchOption.TopDirectoryOnly)
            .Where(path =>
                Path.GetFileName(path).Contains(B2BOlSourceFileNameKeyword, StringComparison.OrdinalIgnoreCase)
                && IsExcelFile(path)
                && !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();

        return sourceFile
            ?? throw new FileNotFoundException(
                $"B2B\uff08ol) source file not found. Expected an Excel file whose name contains '{B2BOlSourceFileNameKeyword}' under {sourceDirectory}.");
    }

    private static bool IsExcelFile(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
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

    private static TableData ReadExcelTableWithTwoRowHeaders(string path, DataFormatter formatter)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = WorkbookFactory.Create(stream);
        var table = ReadExcelTableWithTwoRowHeaders(workbook, workbook.GetSheetAt(0), formatter);
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

    private static TableData ReadExcelTableWithTwoRowHeaders(IWorkbook workbook, ISheet sheet, DataFormatter formatter)
    {
        var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
        var headerRowIndex = FindHeaderRow(sheet, formatter);
        if (headerRowIndex < 0)
        {
            return new TableData([], []);
        }

        var firstHeaderRow = sheet.GetRow(headerRowIndex);
        var secondHeaderRow = sheet.GetRow(headerRowIndex + 1);
        if (secondHeaderRow is null || !HasMergedHeaderRegion(sheet, headerRowIndex))
        {
            return ReadExcelTable(workbook, sheet, formatter);
        }

        var headerColumnCount = Math.Max(
            firstHeaderRow.LastCellNum > 0 ? firstHeaderRow.LastCellNum : 0,
            secondHeaderRow.LastCellNum > 0 ? secondHeaderRow.LastCellNum : 0);
        var firstHeaders = ReadRow(firstHeaderRow, formatter, headerColumnCount, evaluator);
        var secondHeaders = ReadRow(secondHeaderRow, formatter, headerColumnCount, evaluator);
        var headers = firstHeaders
            .Select((header, columnIndex) =>
            {
                var detailHeader = DelimitedTableReader.CleanHeader(secondHeaders[columnIndex]);
                return detailHeader.Length > 0
                    ? detailHeader
                    : DelimitedTableReader.CleanHeader(header);
            })
            .ToList();

        var rows = new List<IReadOnlyList<string>>();
        for (var rowIndex = headerRowIndex + 2; rowIndex <= sheet.LastRowNum; rowIndex++)
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

    private static bool HasMergedHeaderRegion(ISheet sheet, int headerRowIndex)
    {
        for (var i = 0; i < sheet.NumMergedRegions; i++)
        {
            var region = sheet.GetMergedRegion(i);
            if (region.FirstRow <= headerRowIndex
                && region.LastRow >= headerRowIndex + 1)
            {
                return true;
            }
        }

        return false;
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

        ClearDataRows(targetSheet, targetHeaderRowIndex + 1, targetSheet.LastRowNum, preserveFormulas: true);

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

        ClearDataRows(targetSheet, targetHeaderRowIndex + 1, targetSheet.LastRowNum, preserveFormulas: true);

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

    private static int CopyMappingSpreadsheetSheetsToTarget(
        IReadOnlyList<FeishuSpreadsheetSheetData> sourceSheets,
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

        ClearDataRows(targetSheet, targetHeaderRowIndex + 1, targetSheet.LastRowNum, preserveFormulas: true);

        var nextTargetRowIndex = targetHeaderRowIndex + 1;
        foreach (var sourceSheet in MappingSourceSheets)
        {
            var matchedSheet = FindMappingSpreadsheetSheet(sourceSheets, sourceSheet)
                ?? throw new InvalidOperationException(
                    $"Sheet not found in Feishu mapping spreadsheet. Expected one of: {string.Join(", ", sourceSheet.SheetNameCandidates)}. Actual sheets: {string.Join(", ", sourceSheets.Select(sheet => sheet.Title))}");

            var sourceHeaderIndex = BuildHeaderIndex(matchedSheet.Table.Headers);
            var resolvedColumns = copyColumns
                .Select(column => column with
                {
                    SourceColumnIndex = ResolveRequiredColumn(sourceHeaderIndex, column.Header, matchedSheet.Title)
                })
                .ToList();

            foreach (var sourceRow in matchedSheet.Table.Rows)
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

    private static FeishuSpreadsheetSheetData? FindMappingSpreadsheetSheet(
        IReadOnlyList<FeishuSpreadsheetSheetData> sourceSheets,
        MappingSourceSheet expectedSheet)
    {
        return sourceSheets.FirstOrDefault(sheet =>
            expectedSheet.SheetNameCandidates.Any(candidate =>
                string.Equals(NormalizeSheetName(sheet.Title), NormalizeSheetName(candidate), StringComparison.OrdinalIgnoreCase)));
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

    private static void CopyResolvedCellValue(
        ICell? sourceCell,
        ICell targetCell,
        DataFormatter formatter,
        IFormulaEvaluator evaluator)
    {
        if (sourceCell is null)
        {
            targetCell.SetCellValue(string.Empty);
            return;
        }

        if (sourceCell.CellType != CellType.Formula)
        {
            CopyCellValue(sourceCell, targetCell);
            return;
        }

        var evaluated = evaluator.Evaluate(sourceCell);
        if (evaluated is null)
        {
            targetCell.SetCellValue(formatter.FormatCellValue(sourceCell, evaluator));
            return;
        }

        switch (evaluated.CellType)
        {
            case CellType.Boolean:
                targetCell.SetCellValue(evaluated.BooleanValue);
                break;
            case CellType.Numeric:
                targetCell.SetCellValue(evaluated.NumberValue);
                break;
            case CellType.String:
                targetCell.SetCellValue(evaluated.StringValue);
                break;
            case CellType.Blank:
                targetCell.SetCellValue(string.Empty);
                break;
            case CellType.Error:
                targetCell.SetCellErrorValue(unchecked((byte)evaluated.ErrorValue));
                break;
            default:
                targetCell.SetCellValue(formatter.FormatCellValue(sourceCell, evaluator));
                break;
        }
    }

    private static void CopyCellValue(ICell sourceCell, ICell targetCell)
    {
        switch (sourceCell.CellType)
        {
            case CellType.Boolean:
                targetCell.SetCellValue(sourceCell.BooleanCellValue);
                break;
            case CellType.Numeric:
                targetCell.SetCellValue(sourceCell.NumericCellValue);
                break;
            case CellType.String:
                targetCell.SetCellValue(sourceCell.StringCellValue);
                break;
            case CellType.Blank:
                targetCell.SetCellValue(string.Empty);
                break;
            case CellType.Error:
                targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                break;
            default:
                targetCell.SetCellValue(sourceCell.ToString());
                break;
        }
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

    private static string NormalizeImportedCellValue(string sheetName, string header, string value)
    {
        if (string.Equals(sheetName, PaymentsSheetName, StringComparison.OrdinalIgnoreCase)
            && string.Equals(header, "type", StringComparison.OrdinalIgnoreCase)
            && string.Equals(value.Trim(), "Chargeback Refund", StringComparison.OrdinalIgnoreCase))
        {
            return "Refund";
        }

        return value;
    }

    private static void ClearDataRows(ISheet sheet, int startRowIndex, int endRowIndex, bool preserveFormulas)
    {
        if (endRowIndex < startRowIndex)
        {
            return;
        }

        for (var rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            foreach (var cell in row.Cells.ToList())
            {
                if (preserveFormulas && cell.CellType == CellType.Formula)
                {
                    continue;
                }

                row.RemoveCell(cell);
            }
        }
    }

    private static void ClearTemplateDataRows(
        ISheet sheet,
        int startRowIndex,
        int endRowIndex,
        int formulaTemplateRowIndex)
    {
        if (endRowIndex < startRowIndex)
        {
            return;
        }

        for (var rowIndex = startRowIndex; rowIndex <= Math.Min(endRowIndex, sheet.LastRowNum); rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            foreach (var cell in row.Cells.ToList())
            {
                if (rowIndex == formulaTemplateRowIndex && cell.CellType == CellType.Formula)
                {
                    continue;
                }

                row.RemoveCell(cell);
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

    private readonly record struct SkuMapping(string ItemCode, string Owner);

    private readonly record struct DailySummaryMetrics(
        double Orders,
        double SalesTotal,
        double Premium,
        double AdvertisingCost,
        double Payments,
        double Refund,
        double PaymentPremium);

    private enum GeneratedSummaryRowKind
    {
        Title,
        Header,
        Person,
        Total
    }

    private enum GeneratedSummaryBlockKind
    {
        PaymentStore,
        FulfillmentStore,
        FulfillmentAllStore
    }

    private sealed class CellStyleCache
    {
        private readonly IWorkbook _workbook;
        private readonly IDataFormat _dataFormat;
        private readonly Dictionary<string, ICellStyle> _styles = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<short, ICellStyle> _redFontStyles = new();
        private readonly Dictionary<short, ICellStyle> _dateStyles = new();
        private readonly Dictionary<short, ICellStyle> _percentStyles = new();
        private readonly Dictionary<short, ICellStyle> _centeredStyles = new();
        private readonly Dictionary<short, ICellStyle> _plainStyles = new();
        private readonly Dictionary<short, ICellStyle> _summaryDetailBorderStyles = new();
        private readonly Dictionary<short, ICellStyle> _storeDailyTotalStyles = new();

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

        public ICellStyle GetDateStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_dateStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
            }

            style.DataFormat = _dataFormat.GetFormat(ResolveFormat("date"));
            _dateStyles.Add(baseStyleIndex, style);
            return style;
        }

        public ICellStyle GetPercentStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_percentStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
            }

            style.DataFormat = _dataFormat.GetFormat(ResolveFormat("percent"));
            _percentStyles.Add(baseStyleIndex, style);
            return style;
        }

        public ICellStyle GetCenteredStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_centeredStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
            }

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            _centeredStyles.Add(baseStyleIndex, style);
            return style;
        }

        public ICellStyle GetStoreDailyDetailStyle(bool hasBorder)
        {
            var key = hasBorder ? "store-daily-detail-bordered" : "store-daily-detail";
            if (_styles.TryGetValue(key, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            ApplyStoreDailyAlignment(style);
            if (hasBorder)
            {
                ApplyThinBorder(style);
            }

            style.FillPattern = FillPattern.NoFill;
            _styles.Add(key, style);
            return style;
        }

        public ICellStyle GetSummaryDetailBorderStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_summaryDetailBorderStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
            }

            ApplyThinBorder(style);
            _summaryDetailBorderStyles.Add(baseStyleIndex, style);
            return style;
        }

        public ICellStyle GetStoreDailyTotalStyle(ICellStyle? headerStyle)
        {
            var headerStyleIndex = headerStyle?.Index ?? 0;
            if (_storeDailyTotalStyles.TryGetValue(headerStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            ApplyStoreDailyAlignment(style);
            ApplyThinBorder(style);
            if (headerStyle is not null)
            {
                style.FillPattern = headerStyle.FillPattern;
                style.FillForegroundColor = headerStyle.FillForegroundColor;
                style.FillBackgroundColor = headerStyle.FillBackgroundColor;
                CopyXssfFillColors(headerStyle, style);
            }

            _storeDailyTotalStyles.Add(headerStyleIndex, style);
            return style;
        }

        public ICellStyle GetPlainStyle(ICellStyle? baseStyle)
        {
            var baseStyleIndex = baseStyle?.Index ?? 0;
            if (_plainStyles.TryGetValue(baseStyleIndex, out var style))
            {
                return style;
            }

            style = _workbook.CreateCellStyle();
            if (baseStyle is not null)
            {
                style.CloneStyleFrom(baseStyle);
                style.SetFont(CreateNonBoldFont(baseStyle));
            }

            style.FillPattern = FillPattern.NoFill;
            _plainStyles.Add(baseStyleIndex, style);
            return style;
        }

        private IFont CreateNonBoldFont(ICellStyle baseStyle)
        {
            var baseFont = _workbook.GetFontAt(baseStyle.FontIndex);
            var font = _workbook.CreateFont();
            font.FontName = baseFont.FontName;
            font.FontHeightInPoints = baseFont.FontHeightInPoints;
            font.Color = baseFont.Color;
            font.IsItalic = baseFont.IsItalic;
            font.Underline = baseFont.Underline;
            font.TypeOffset = baseFont.TypeOffset;
            font.IsStrikeout = baseFont.IsStrikeout;
            font.IsBold = false;
            return font;
        }

        private static void ApplyStoreDailyAlignment(ICellStyle style)
        {
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
        }

        private static void ApplyThinBorder(ICellStyle style)
        {
            style.BorderTop = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
        }

        private static void CopyXssfFillColors(ICellStyle sourceStyle, ICellStyle targetStyle)
        {
            var sourceType = sourceStyle.GetType();
            var targetType = targetStyle.GetType();
            if (!sourceType.Name.Contains("XSSFCellStyle", StringComparison.Ordinal)
                || !targetType.Name.Contains("XSSFCellStyle", StringComparison.Ordinal))
            {
                return;
            }

            CopyXssfColor(
                sourceType.GetProperty("FillForegroundXSSFColor")?.GetValue(sourceStyle),
                targetStyle,
                targetType,
                "SetFillForegroundColor");
            CopyXssfColor(
                sourceType.GetProperty("FillBackgroundXSSFColor")?.GetValue(sourceStyle),
                targetStyle,
                targetType,
                "SetFillBackgroundColor");
        }

        private static void CopyXssfColor(object? color, ICellStyle targetStyle, Type targetType, string methodName)
        {
            if (color is null)
            {
                return;
            }

            var method = targetType
                .GetMethods()
                .FirstOrDefault(methodInfo =>
                    string.Equals(methodInfo.Name, methodName, StringComparison.Ordinal)
                    && methodInfo.GetParameters().Length == 1
                    && methodInfo.GetParameters()[0].ParameterType.IsAssignableFrom(color.GetType()));
            method?.Invoke(targetStyle, [color]);
        }

        private static string ResolveFormat(string key)
        {
            if (string.Equals(key, "percent", StringComparison.OrdinalIgnoreCase))
            {
                return "0.00%";
            }

            if (string.Equals(key, "date", StringComparison.OrdinalIgnoreCase))
            {
                return "m\"月\"d\"日\"";
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
