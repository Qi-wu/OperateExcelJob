using Microsoft.Extensions.Configuration;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OperateExcel.Job;

var profile = new DailyReportProfileOptions
{
    People = [new() { Name = "Tester" }],
    Stores = [new() { Name = "TestStore", People = ["Tester"], MappingSheetNameCandidates = ["Test mapping"] }],
    SkuOwners = [new() { OwnerCode = "T", Name = "Tester" }]
};
using var http = new HttpClient();
var feishuOptions = Options.Create(new FeishuOptions { Enabled = false, UploadGeneratedAttachmentsEnabled = false });
ExcelImportJob NewJob(DailyReportProfileOptions options) => new(
    Options.Create(new ExcelImportOptions { RootDirectory = Path.Combine(Path.GetTempPath(), "daily-report-regression-no-sources") }),
    feishuOptions, new FixedOptionsMonitor<DailyReportProfileOptions>(options),
    new FeishuApiClient(http, feishuOptions, NullLogger<FeishuApiClient>.Instance), NullLogger<ExcelImportJob>.Instance);
var job = NewJob(profile);
var formatter = new DataFormatter();
var checks = 0;
void Check(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
    checks++;
}
object? Call(string name, object? target, params object?[] arguments)
{
    var method = typeof(ExcelImportJob).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)
        ?? throw new MissingMethodException(name);
    try { return method.Invoke(target, arguments); }
    catch (TargetInvocationException exception) when (exception.InnerException is not null)
    {
        ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
        throw;
    }
}
object Styles(IWorkbook workbook) => Activator.CreateInstance(
    typeof(ExcelImportJob).GetNestedType("CellStyleCache", BindingFlags.NonPublic)!, workbook)!;
void Reject(Action action, string message)
{
    try { action(); }
    catch (InvalidOperationException) { checks++; return; }
    throw new InvalidOperationException(message);
}
string[] currentHeaders = ["归属", "SKU", "Itemcode", "销量", "单价", "运费", "订单收入", "采购", "固定费用",
    "单个毛利比", "毛利", "店铺", "销售总额", "库存", "账号", "", "", "归属", "sku", "quantity", "item-price",
    "item-tax", "shipping-price", "gift wrap credits", "total", "销售总额"];
XSSFWorkbook Fixture(bool legacy)
{
    var workbook = new XSSFWorkbook();
    var payment = workbook.CreateSheet("模板P");
    var header = payment.CreateRow(1);
    var headers = legacy ? currentHeaders.Where((_, i) => i != 8).ToArray() : currentHeaders;
    for (var i = 0; i < headers.Length; i++) header.CreateCell(i).SetCellValue(headers[i]);
    payment.CreateRow(2).CreateCell(0).SetCellFormula("Q3");
    var fulfillment = workbook.CreateSheet("模版F");
    var fheader = fulfillment.CreateRow(1);
    fheader.CreateCell(3).SetCellValue("销量");
    fheader.CreateCell(7).SetCellValue("采购");
    fheader.CreateCell(8).SetCellValue("固定费用");
    fulfillment.CreateRow(2).CreateCell(8).SetCellFormula("M3*0.05");
    return workbook;
}
Check(Math.Abs((double)Call("CalculatePremium", null, 100D, 60D)! - 22D) < 1e-9, "Below-threshold premium");
Check(Math.Abs((double)Call("CalculatePremium", null, 300D, 100D)! - 155D) < 1e-9, "Above-threshold premium");
Check((double)Call("CalculatePremium", null, 100D, 0D)! == 0D, "Missing-cost rule changed");
foreach (var legacy in new[] { true, false })
{
    using var workbook = Fixture(legacy);
    Call("PrepareDailyTemplateLayout", null, workbook, formatter, new List<string>());
    var payment = workbook.GetSheet("模板P");
    for (var col = 0; col < currentHeaders.Length; col++)
        Check(payment.GetRow(1).GetCell(col).StringCellValue == currentHeaders[col], $"Header {col}");
    Check(payment.GetRow(2).GetCell(10).CellFormula == "IFERROR((G3-H3-I3),0)", "Payment profit formula");
    var firstFormula = payment.GetRow(2).GetCell(8).CellFormula;
    Call("PrepareDailyTemplateLayout", null, workbook, formatter, new List<string>());
    Check(payment.GetRow(1).LastCellNum == 26 && payment.GetRow(2).GetCell(8).CellFormula == firstFormula, "Upgrade is not idempotent");
    var summary = payment.CreateRow(1501);
    Call("SetPaymentStoreSummaryFormulas", null, summary, 2, "TestStore");
    Check(summary.GetCell(3).CellFormula.Contains("SUMIFS(M:M,A:A,C1502,O:O,"), "Payment sales/account ranges");
    Check(summary.GetCell(5).CellFormula.Contains("SUMIFS(K:K,A:A,C1502,O:O,"), "Payment profit range");
    var fsummary = workbook.GetSheet("模版F").CreateRow(1335);
    Call("SetStoreSummaryFormulas", null, fsummary, 2, 1334, "TestStore");
    Check(fsummary.GetCell(8).CellFormula.Contains("!M:M") && fsummary.GetCell(8).CellFormula.Contains("!O:O"), "F -> P sales/account ranges");
    Call("SetAllStoreSummaryFormulas", null, fsummary, 2, 1334);
    Check(fsummary.GetCell(10).CellFormula.Contains("!K:K"), "All-store profit range");
    var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
    var row = payment.GetRow(2);
    foreach (var (quantity, procurement, expectedFee) in new[] { (2D, 120D, 6D), (0D, 0D, 0D) })
    {
        row.GetCell(3).SetCellFormula(null); row.GetCell(3).SetCellValue(quantity);
        row.GetCell(7).SetCellFormula(null); row.GetCell(7).SetCellValue(procurement);
        evaluator.ClearAllCachedResultValues();
        Check(Math.Abs(evaluator.Evaluate(row.GetCell(8)).NumberValue - expectedFee) < 1e-9, "Quantity/zero-quantity fixed fee");
    }
}
using (var invalid = Fixture(false))
{
    invalid.GetSheet("模板P").GetRow(1).GetCell(14).SetCellValue("unexpected");
    Reject(() => Call("PrepareDailyTemplateLayout", null, invalid, formatter, new List<string>()), "Unknown layout accepted");
}
// Generated blocks must not inherit old title merges at rows now occupied by people.
foreach (var (method, staleRow, lastColumn) in new[]
{
    ("ClearGeneratedFulfillmentSummaryArea", 1364, 10),
    ("ClearGeneratedPaymentSummaryArea", 1505, 6)
})
{
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("merge-regression");
    sheet.CreateRow(staleRow).CreateCell(2).SetCellValue("stale title");
    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(staleRow, staleRow, 2, lastColumn));
    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 2, lastColumn));
    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(staleRow, staleRow, 12, 14));
    Call(method, null, sheet);
    Check(sheet.NumMergedRegions == 2, $"{method} retained a stale person-row merge");
    Check(Enumerable.Range(0, sheet.NumMergedRegions).All(i =>
        sheet.GetMergedRegion(i).FirstRow == 0 || sheet.GetMergedRegion(i).FirstColumn == 12), "Unrelated merged region changed");
    Call(method, null, sheet);
    Check(sheet.NumMergedRegions == 2, "Merged-region cleanup is not idempotent");
}
using (var raw = new XSSFWorkbook())
{
    var sheet = raw.CreateSheet("raw");
    sheet.CreateRow(0).CreateCell(0).SetCellValue("sku");
    sheet.CreateRow(1).CreateCell(1).SetCellFormula("A2");
    sheet.CreateRow(2).CreateCell(0).SetCellValue("current-day-sku");
    sheet.CreateRow(8).CreateCell(1).SetCellFormula("A9");
    Call("TrimImportedSheetTail", null, sheet, 0, 3);
    Check(sheet.LastRowNum == 2 && sheet.GetRow(8) is null, "Stale imported formulas retained");
    Check(sheet.GetRow(2).GetCell(0).StringCellValue == "current-day-sku", "Current-day import was trimmed");
    Call("TrimImportedSheetTail", null, sheet, 0, 1);
    Check(sheet.LastRowNum == 1 && sheet.GetRow(1).GetCell(1).CellType == CellType.Formula, "Empty-day formula seed lost");
}
using (var history = new XSSFWorkbook())
{
    var sheet = history.CreateSheet("汇总");
    var row = sheet.CreateRow(1);
    row.CreateCell(0).SetCellValue(new DateTime(2026, 9, 21));
    row.CreateCell(4).SetCellValue(123.45);
    Call("EnsureAppendOnlySummaryDate", job, history, new DateOnly(2026, 9, 22), formatter);
    Reject(() => Call("EnsureAppendOnlySummaryDate", job, history, new DateOnly(2026, 9, 21), formatter), "Duplicate date accepted");
    Reject(() => Call("EnsureAppendOnlySummaryDate", job, history, new DateOnly(2026, 9, 20), formatter), "Older date accepted");
    Call("RestoreSummaryDateFormats", job, history, Styles(history));
    Check(row.GetCell(4).NumericCellValue == 123.45 && row.GetCell(0).NumericCellValue == new DateTime(2026, 9, 21).ToOADate(), "History values changed");
    Check(DateUtil.IsCellDateFormatted(row.GetCell(0)), "Date style not restored");
}

void CheckGeneratedMerges(IWorkbook workbook)
{
    foreach (var (name, firstRow, lastColumn) in new[] { ("模版F", 1334, 10), ("模板P", 1500, 6) })
    {
        var sheet = workbook.GetSheet(name);
        var count = 0;
        for (var i = 0; i < sheet.NumMergedRegions; i++)
        {
            var region = sheet.GetMergedRegion(i);
            if (region.LastRow < firstRow || region.LastColumn < 2 || region.FirstColumn > lastColumn) continue;
            count++;
            Check(name == "模版F" && region.FirstRow == region.LastRow
                && region.FirstColumn == 2 && region.LastColumn == 10
                && sheet.GetRow(region.FirstRow)?.GetCell(2)?.ToString() == "汇总总计",
                $"Unexpected generated merge: {name}!{region.FormatAsString()}");
        }
        Check(count == (name == "模版F" ? 1 : 0), $"Unexpected title merge count in {name}");
    }
}

// Optional read-only real-workbook regression. All generated files stay in an explicitly supplied output directory.
if (args.Length > 0)
{
    if (args.Length < 3) throw new ArgumentException("Usage: <profile.json> <output-directory> <old.xlsx> [new.xlsx]");
    var configuration = new Microsoft.Extensions.Configuration.ConfigurationBuilder()
        .AddJsonFile(Path.GetFullPath(args[0])).Build();
    var realProfile = configuration.GetSection("DailyReportProfile").Get<DailyReportProfileOptions>()!;
    var realJob = NewJob(realProfile);
    var names = new[] { "汇总" }.Concat(realProfile.Stores.Select(s => s.Name!)).ToArray();
    Directory.CreateDirectory(args[1]);
    for (var i = 2; i < args.Length; i++)
    {
        using var input = File.OpenRead(args[i]);
        using var workbook = WorkbookFactory.Create(input);
        var before = Snapshot(workbook, names);
        Call("ApplyPostImportRules", realJob, workbook, new DateOnly(2026, 9, 23), formatter, Styles(workbook), new List<string>());
        CheckGeneratedMerges(workbook);
        var after = Snapshot(workbook, names);
        foreach (var (key, value) in before) Check(after.TryGetValue(key, out var actual) && actual == value, $"Historical cell changed: {key}");
        Reject(() => Call("ApplyPostImportRules", realJob, workbook, new DateOnly(2026, 9, 23), formatter, Styles(workbook), new List<string>()), "Rerun appended duplicate history");
        var output = Path.Combine(args[1], $"regression-{i - 1}.xlsx");
        using (var stream = File.Create(output)) workbook.Write(stream, leaveOpen: true);
        using var rereadStream = File.OpenRead(output);
        using var reread = WorkbookFactory.Create(rereadStream);
        CheckGeneratedMerges(reread);
        var persisted = Snapshot(reread, names);
        foreach (var (key, value) in before) Check(persisted.TryGetValue(key, out var actual) && actual == value, $"Serialized history changed: {key}");
        Console.WriteLine($"PASS real workbook {Path.GetFileName(args[i])}: {before.Count} historical cells preserved; output {output}");
    }
}
Console.WriteLine($"PASS: {checks} checks; no Feishu/network calls made by regression tests.");

static Dictionary<string, string> Snapshot(IWorkbook workbook, IEnumerable<string> names)
{
    var snapshot = new Dictionary<string, string>();
    foreach (var name in names)
    {
        var sheet = workbook.GetSheet(name);
        if (sheet is null) continue;
        foreach (IRow row in sheet)
        foreach (var cell in row.Cells)
        {
            var value = cell.CellType switch
            {
                CellType.Formula => "F:" + cell.CellFormula,
                CellType.Numeric => "N:" + cell.NumericCellValue.ToString("R", CultureInfo.InvariantCulture),
                CellType.String => "S:" + cell.StringCellValue,
                CellType.Boolean => "B:" + cell.BooleanCellValue,
                CellType.Error => "E:" + cell.ErrorCellValue,
                _ => ""
            };
            if (value is not "" and not "S:") snapshot[$"{name}!{row.RowNum}:{cell.ColumnIndex}"] = value;
        }
    }
    return snapshot;
}

sealed class FixedOptionsMonitor<T>(T value) : IOptionsMonitor<T>
{
    public T CurrentValue => value;
    public T Get(string? name) => value;
    public IDisposable? OnChange(Action<T, string?> listener) => null;
}
