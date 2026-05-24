namespace OperateExcel.Job;

public sealed class ExcelImportOptions
{
    public string RootDirectory { get; set; } = @"D:\code\OperateExcelTemp";
    public string TemplateFilePath { get; set; } = @"D:\code\OperateExcelTemp\Temp.xlsx";
    public int DateOffsetDays { get; set; } = -1;
    public string? ProcessingDateOverride { get; set; }
    public string DailyCron { get; set; } = "0 2 * * *";
    public bool ClearExistingData { get; set; } = true;
}
