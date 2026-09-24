using System.Globalization;
using NPOI.SS.UserModel;

namespace OperateExcel.Job;

public sealed partial class ExcelImportJob
{
    private const double ProcurementFixedFeeRate = 0.05D;

    private static readonly string[] CurrentPaymentDetailHeaders =
    [
        "\u5f52\u5c5e", "SKU", "Itemcode", "\u9500\u91cf", "\u5355\u4ef7", "\u8fd0\u8d39",
        "\u8ba2\u5355\u6536\u5165", "\u91c7\u8d2d", "\u56fa\u5b9a\u8d39\u7528", "\u5355\u4e2a\u6bdb\u5229\u6bd4",
        "\u6bdb\u5229", "\u5e97\u94fa", "\u9500\u552e\u603b\u989d", "\u5e93\u5b58", "\u8d26\u53f7", "", "",
        "\u5f52\u5c5e", "sku", "quantity", "item-price", "item-tax", "shipping-price", "gift wrap credits", "total",
        "\u9500\u552e\u603b\u989d"
    ];

    private static void PrepareDailyTemplateLayout(IWorkbook workbook, DataFormatter formatter, ICollection<string> messages)
    {
        var paymentSheet = workbook.GetSheet(PaymentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentTemplateSheetName}");
        var headerRowIndex = FindHeaderRow(paymentSheet, formatter);
        var headers = ReadHeaders(paymentSheet, headerRowIndex, formatter);
        var legacyHeaders = CurrentPaymentDetailHeaders.Where((_, index) => index != 8).ToArray();
        var isLegacy = MatchesDetailHeaders(headers, legacyHeaders);
        if (!isLegacy && !MatchesDetailHeaders(headers, CurrentPaymentDetailHeaders))
        {
            throw new InvalidOperationException($"Unsupported detail layout in {PaymentTemplateSheetName}. Expected the legacy layout or the layout with fixed fees in column I.");
        }

        var fulfillmentSheet = workbook.GetSheet(FulfillmentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {FulfillmentTemplateSheetName}");
        var fulfillmentHeaderRow = FindHeaderRow(fulfillmentSheet, formatter);
        var fulfillmentHeaders = ReadHeaderIndex(fulfillmentSheet, fulfillmentHeaderRow, formatter);
        foreach (var (name, column) in new[] { ("\u9500\u91cf", 3), ("\u91c7\u8d2d", 7), ("\u56fa\u5b9a\u8d39\u7528", 8) })
        {
            if (ResolveRequiredColumn(fulfillmentHeaders, name, FulfillmentTemplateSheetName) != column)
            {
                throw new InvalidOperationException($"Unsupported {name} column in {FulfillmentTemplateSheetName}.");
            }
        }

        var headerRow = paymentSheet.GetRow(headerRowIndex);
        var seedRow = paymentSheet.GetRow(headerRowIndex + 1) ?? paymentSheet.CreateRow(headerRowIndex + 1);
        var headerStyles = headerRow.Cells.ToDictionary(cell => cell.ColumnIndex, cell => cell.CellStyle);
        var seedStyles = seedRow.Cells.ToDictionary(cell => cell.ColumnIndex, cell => cell.CellStyle);
        if (isLegacy)
        {
            for (var column = legacyHeaders.Length - 1; column >= 8; column--)
            {
                paymentSheet.SetColumnWidth(column + 1, paymentSheet.GetColumnWidth(column));
            }
            paymentSheet.SetColumnWidth(8, paymentSheet.GetColumnWidth(7));
        }

        // Rebuild only the header and formula seed; today's import replaces the detail rows immediately afterwards.
        // Do not shift workbook-wide references or touch accumulated daily summary values.
        foreach (var cell in seedRow.Cells.ToList())
        {
            seedRow.RemoveCell(cell);
        }
        for (var column = 0; column < CurrentPaymentDetailHeaders.Length; column++)
        {
            var sourceColumn = isLegacy && column >= 8 ? column - 1 : column;
            var headerCell = headerRow.GetCell(column) ?? headerRow.CreateCell(column);
            headerCell.SetCellValue(CurrentPaymentDetailHeaders[column]);
            if (headerStyles.TryGetValue(sourceColumn, out var headerStyle))
            {
                headerCell.CellStyle = headerStyle;
            }
            var seedCell = seedRow.CreateCell(column);
            if (seedStyles.TryGetValue(sourceColumn, out var seedStyle))
            {
                seedCell.CellStyle = seedStyle;
            }
        }

        var row = seedRow.RowNum + 1;
        var formulas = new Dictionary<int, string>
        {
            [0] = $"R{row}",
            [1] = $"S{row}",
            [2] = $"VLOOKUP(B{row},\u6620\u5c04\u8868!A:C,3,0)",
            [3] = $"T{row}",
            [4] = $"IFERROR(U{row}/D{row},0)*D{row}",
            [5] = $"IFERROR(W{row}/D{row},0)*D{row}",
            [6] = $"IFERROR(IF((E{row}+F{row})>200,(E{row}+F{row}-200)*0.9+200*0.85,(E{row}+F{row})*0.85),0)",
            [7] = $"VLOOKUP(C{row},'B2B\uff08ol)'!A:J,10,0)*D{row}",
            [8] = ProcurementFixedFeeFormula(row),
            [9] = $"IFERROR((G{row}-H{row}-I{row})/(E{row}+F{row}),0)",
            [10] = $"IFERROR((G{row}-H{row}-I{row}),0)",
            [11] = $"VLOOKUP(C{row},'B2B\uff08ol)'!A:N,3,0)",
            [12] = $"(E{row}+F{row})",
            [13] = $"VLOOKUP(C{row},'B2B\uff08ol)'!A:K,11,0)",
            [14] = $"VLOOKUP(B{row},payments!E:AG,29,0)",
            [17] = $"VLOOKUP(S{row},\u6620\u5c04\u8868!A:G,7,0)",
            [25] = $"M{row}"
        };
        foreach (var (column, formula) in formulas)
        {
            SetFormulaCell(seedRow, column, formula);
        }

        var fulfillmentSeed = fulfillmentSheet.GetRow(fulfillmentHeaderRow + 1)
            ?? throw new InvalidOperationException($"Formula seed row missing in {FulfillmentTemplateSheetName}.");
        SetFormulaCell(fulfillmentSeed, 8, ProcurementFixedFeeFormula(fulfillmentSeed.RowNum + 1));
        messages.Add($"Prepared daily detail template ({(isLegacy ? "upgraded legacy payment layout" : "current payment layout")}); fixed fee = procurement x 5%. Existing daily summary values were not recalculated.");
    }

    private static void TrimImportedSheetTail(ISheet sheet, int headerRowIndex, int nextDataRowIndex)
    {
        // Keep one formula seed on an empty day, but never carry yesterday's trailing formula-only rows forward.
        var firstUnusedRow = Math.Max(headerRowIndex + 2, nextDataRowIndex);
        for (var rowIndex = sheet.LastRowNum; rowIndex >= firstUnusedRow; rowIndex--)
        {
            var row = sheet.GetRow(rowIndex);
            if (row is not null)
            {
                sheet.RemoveRow(row);
            }
        }
    }

    private static bool MatchesDetailHeaders(IReadOnlyList<string> actual, IReadOnlyList<string> expected)
    {
        return expected.Select((header, index) => string.Equals(
                index < actual.Count ? actual[index] : string.Empty, header, StringComparison.OrdinalIgnoreCase)).All(matches => matches)
            && actual.Skip(expected.Count).All(string.IsNullOrWhiteSpace);
    }

    private static string ProcurementFixedFeeFormula(int row)
    {
        var rate = ProcurementFixedFeeRate.ToString(CultureInfo.InvariantCulture);
        return $"IF(D{row}=0,0,(H{row}*{rate}/D{row})*D{row})";
    }

    private static PaymentSummaryColumns ReadPaymentSummaryColumns(IWorkbook workbook)
    {
        var sheet = workbook.GetSheet(PaymentTemplateSheetName)
            ?? throw new InvalidOperationException($"Sheet not found: {PaymentTemplateSheetName}");
        var formatter = new DataFormatter();
        // Prefer the left-hand calculated columns when headers occur again in the imported detail block.
        var headers = ReadHeaderIndex(sheet, FindHeaderRow(sheet, formatter), formatter);
        string Range(string header)
        {
            var column = ColumnIndexToName(ResolveRequiredColumn(headers, header, sheet.SheetName));
            return $"{column}:{column}";
        }
        return new PaymentSummaryColumns(Range("\u9500\u552e\u603b\u989d"), Range("\u6bdb\u5229"),
            Range("\u8d26\u53f7"), Range("\u5f52\u5c5e"), Range("\u9500\u91cf"));
    }

    private IEnumerable<ISheet> ExistingDailySummarySheets(IWorkbook workbook)
    {
        return new[] { SummarySheetName }.Concat(FulfillmentSummaryStores)
            .Select(name => FindSheet(workbook, name)).OfType<ISheet>();
    }

    private void EnsureAppendOnlySummaryDate(IWorkbook workbook, DateOnly processingDate, DataFormatter formatter)
    {
        foreach (var sheet in ExistingDailySummarySheets(workbook))
        {
            for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                foreach (var columnIndex in new[] { PaymentFirstDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex })
                {
                    var cell = sheet.GetRow(rowIndex)?.GetCell(columnIndex);
                    DateOnly? existingDate = null;
                    if (cell?.CellType == CellType.Numeric && cell.NumericCellValue is > 0 and < 2958466)
                    {
                        existingDate = DateOnly.FromDateTime(DateTime.FromOADate(cell.NumericCellValue));
                    }
                    else if (cell?.CellType == CellType.String && TryParseDateCell(formatter.FormatCellValue(cell), out var date))
                    {
                        existingDate = DateOnly.FromDateTime(date);
                    }
                    if (existingDate >= processingDate)
                    {
                        throw new InvalidOperationException($"Cannot append report for {processingDate:yyyy-MM-dd}: sheet '{sheet.SheetName}' already contains {existingDate:yyyy-MM-dd}. Historical summaries are append-only; use a base report earlier than the processing date.");
                    }
                }
            }
        }
    }

    private void RestoreSummaryDateFormats(IWorkbook workbook, CellStyleCache cellStyleCache)
    {
        foreach (var sheet in ExistingDailySummarySheets(workbook))
        {
            for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                foreach (var columnIndex in new[] { PaymentFirstDailySummaryStartColumnIndex, PaymentSecondDailySummaryStartColumnIndex })
                {
                    var cell = sheet.GetRow(rowIndex)?.GetCell(columnIndex);
                    if (cell?.CellType == CellType.Numeric && cell.NumericCellValue is > 0 and < 2958466)
                    {
                        cell.CellStyle = cellStyleCache.GetDateStyle(cell.CellStyle);
                    }
                }
            }
        }
    }

    private readonly record struct PaymentSummaryColumns(string Sales, string Profit, string Account, string Owner, string Quantity);
}
