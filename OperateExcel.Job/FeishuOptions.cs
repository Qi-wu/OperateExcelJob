namespace OperateExcel.Job;

public sealed class FeishuOptions
{
    public bool Enabled { get; set; } = true;
    public bool UploadGeneratedAttachments { get; set; } = true;
    public string AppId { get; set; } = string.Empty;
    public string AppSecret { get; set; } = string.Empty;
    public string WikiUrl { get; set; } = "https://bcnt3e3uyrxk.feishu.cn/base/T45fbKkq1azywosncEOcF8fynsY?table=tbldlsREWMxiL8sC&view=vewEfMbIRP";
    public string? BitableAppToken { get; set; } = "T45fbKkq1azywosncEOcF8fynsY";
    public string? TableId { get; set; } = "tbldlsREWMxiL8sC";
    public string? ViewId { get; set; } = "vewEfMbIRP";
    public string DateFieldName { get; set; } = "\u65e5\u671f";
    public string AttachmentFieldName { get; set; } = "\u635f\u76ca\u8868";
    public string MappingAttachmentFieldName { get; set; } = "\u6620\u5c04\u8868";
    public string RmaAttachmentFieldName { get; set; } = "RMA\u7533\u8bf7\u8868";
    public string DailyReportAttachmentFieldName { get; set; } = "\u65e5\u62a5";
    public string CompletionFieldName { get; set; } = "\u662f\u5426\u5b8c\u6210";
    public string CompletionValue { get; set; } = "\u662f";
    public string SourceSheetName { get; set; } = "B2BOL";
    public string TargetSheetName { get; set; } = "B2B\uff08ol)";
    public string MappingTargetSheetName { get; set; } = "\u6620\u5c04\u8868";
    public string MappingSpreadsheetUrl { get; set; } = "https://bcnt3e3uyrxk.feishu.cn/wiki/OVeOwY8nFiEaLokfPrlceAjinbe";
    public string[] MappingSpreadsheetSheetUrls { get; set; } =
    [
        "https://bcnt3e3uyrxk.feishu.cn/wiki/OVeOwY8nFiEaLokfPrlceAjinbe?sheet=0146aa",
        "https://bcnt3e3uyrxk.feishu.cn/wiki/OVeOwY8nFiEaLokfPrlceAjinbe?sheet=F4lcCs",
        "https://bcnt3e3uyrxk.feishu.cn/wiki/OVeOwY8nFiEaLokfPrlceAjinbe?sheet=nlYgkl"
    ];
}
