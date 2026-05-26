namespace OperateExcel.Job;

public sealed class FeishuOptions
{
    public bool Enabled { get; set; } = true;
    public string AppId { get; set; } = string.Empty;
    public string AppSecret { get; set; } = string.Empty;
    public string WikiUrl { get; set; } = "https://bcnt3e3uyrxk.feishu.cn/wiki/TPsJw75WriCKtTkMNjscSwU0nAe?table=tbl7Bz3tw6FvEJHW&view=vewbgLs0Mv";
    public string? BitableAppToken { get; set; }
    public string? TableId { get; set; }
    public string? ViewId { get; set; }
    public string DateFieldName { get; set; } = "\u65e5\u671f";
    public string AttachmentFieldName { get; set; } = "\u635f\u76ca\u8868";
    public string MappingAttachmentFieldName { get; set; } = "\u6620\u5c04\u8868";
    public string RmaAttachmentFieldName { get; set; } = "RMA\u7533\u8bf7\u8868";
    public string SourceSheetName { get; set; } = "B2BOL";
    public string TargetSheetName { get; set; } = "B2B\uff08ol)";
    public string MappingTargetSheetName { get; set; } = "\u6620\u5c04\u8868";
}
