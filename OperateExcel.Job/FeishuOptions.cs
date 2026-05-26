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
    public string DateFieldName { get; set; } = "日期";
    public string AttachmentFieldName { get; set; } = "损益表";
    public string SourceSheetName { get; set; } = "B2BOL";
    public string TargetSheetName { get; set; } = "B2B（ol)";
}
