# 日报模板升级与历史保护

## 适用范围

本次升级以 2026-09-22 上传的日报格式为准。每次新生成的日报采用新费用口径；不回填、不重算基础工作簿中已有的「汇总」和各店铺历史数据。重新生成 2026-09-22 时，仍从飞书 2026-09-21 日报承接历史，再生成当天数据。

- 「模板P」自动识别旧版或 I 列为“固定费用”的新版结构。旧版的每日明细表头升级至 A:Z，生成当前日期的明细时重建公式种子。
- 不对整个工作簿执行插列，避免改动历史引用；不需要把上传的样例文件写死为生产模板路径。
- 「模版F」「模板P」固定费用均使用采购金额的 5%；销量为 0 时费用为 0。
- 「模板P」毛利为订单收入减采购金额减固定费用，毛利率同步扣除固定费用。
- 人员/店铺层面的 fulfillment 和 payment 毛利都使用同一费用计算函数。缺少采购成本时，继续沿用原来毛利为 0 的规则。
- 生成汇总时从「模板P」表头解析销售总额、毛利、账号、归属和销量列，不再使用旧的 L/J/N 硬编码。
- 已有汇总日期的数值和公式保持不变，仅修复日期列的显示格式。重复日期或早于基础文件历史日期的追加会被拒绝，防止 SUMIFS 重复计入。
- 原有每月 1 日清空新月份工作簿中月度汇总区的行为保持不变；模板升级本身不清空历史。
- AN 店铺人员归属保持配置原样，本次没有删除人员。
- 原始数据比上一日少时，清理尾部残留公式行；「模版F」跳过空订单号行。明细超过预留区域时直接报错，避免覆盖汇总区。

配置中的真实路径和凭据、人员配置、用户原有 README 修改均未改动。

## 构建与回归

```powershell
dotnet build OperateExcel.Job/OperateExcel.Job.csproj
dotnet run --project OperateExcel.Job.RegressionTests/OperateExcel.Job.RegressionTests.csproj
```

可选真实模板回归：

```powershell
dotnet run --project OperateExcel.Job.RegressionTests/OperateExcel.Job.RegressionTests.csproj -- `
  "E:/OperateExcelJob-main/OperateExcel.Job/daily-report-profile.json" `
  "E:/OperateExcelJob-main/verify-output/choose-a-new-regression-directory" `
  "<旧版日报绝对路径>" "<新版日报绝对路径>"
```

回归程序不调用飞书，仅在指定输出目录写入测试副本。可选真实模板测试使用 2026-09-22 原始明细模拟追加 2026-09-23，目的是验证新旧布局、历史保护和序列化，**生成的 regression 文件不是真实 9 月 23 日日报**。

## 仅生成本地 9 月 22 日日报

```powershell
dotnet run --project OperateExcel.Job/OperateExcel.Job.csproj -- --run-once `
  --date=2026-09-22 `
  "--ExcelImport:RootDirectory=E:/财务每日报表下载(美国)" `
  "--output-dir=E:/OperateExcelJob-main/verify-output/choose-a-new-output-directory" `
  --Feishu:UploadGeneratedAttachmentsEnabled=false `
  "--FileLog:Directory=E:/OperateExcelJob-main/verify-output/choose-a-new-output-directory/logs" `
  --Logging:LogLevel:System.Net.Http.HttpClient=Warning
```

保持飞书读取开启，以读取前一天日报及已配置的映射表；`UploadGeneratedAttachmentsEnabled=false` 同时跳过日报上传、完成标记写入和 RMA 更新。这个命令行覆盖只影响本次运行，不修改生产配置。指定不同的隔离输出目录，避免覆盖已有生成文件。

## 2026-09-24 验收结果

最终文件：`E:/OperateExcelJob-main/verify-output/2026-09-22-us-merge-fixed/2026.9月日报22日.xlsx`。

- 真实来源：`E:/财务每日报表下载(美国)/2026-09-22`，基础日报为飞书 2026-09-21 日报。
- 导入 fulfillment 118 行、payments 117 行、广告 823 行。
- 过滤等待/取消订单后，「模版F」103 行明细；「模板P」105 行订单明细。
- 与飞书 9 月 21 日基础文件比对，5 张汇总/店铺页共 14,835 个原有非空单元格的数值、公式全部保持一致，原有公式缓存值也无变化。
- 独立复算新生成的 9 月 22 日人员及店铺指标，共核对 184 个数值，差异为 0。
- 日志明确记录跳过 RMA 更新和日报上传。
- 工作簿开启自动计算及打开时完整重算；新生成的公式会在 Excel 打开时计算。

## 合并单元格修复

发现基础日报残留 `模版F!C1365:K1365` 合并区域。原生成区域清理只删除单元格内容，未删除合并定义，导致第 1365 行的人员姓名横跨指标列显示。

已在 fulfillment/payment 自动生成汇总区清理时移除相交的旧合并，再由生成逻辑为当前标题建立所需合并。范围外的合并保持不变。

新增回归覆盖：旧合并清理、区域外合并保留、重复清理、真实文件生成及保存后合并结构。修复前测试可复现失败，修复后全部通过。

重新生成的 9 月 22 日日报中，自动生成的 F 汇总区只保留 `C1362:K1362` 标题合并；`C1365` 为丁芳芳，`D1365:K1365` 均为独立公式单元格。与前次生成文件相比，五张汇总/店铺页的数值和公式未改变；历史缓存及 184 项当天指标验证再次通过。本次同样未上传飞书。
