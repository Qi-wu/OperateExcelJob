namespace OperateExcel.Job;

public sealed record ImportResult(
    DateOnly ProcessingDate,
    string SourceDirectory,
    string TemplateFilePath,
    int FulfillmentRows,
    int PaymentRows,
    int AdvertisingRows,
    IReadOnlyList<string> Messages);
