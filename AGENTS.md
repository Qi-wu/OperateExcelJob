# 仓库指南

## 项目概览

- 本仓库包含一个 .NET 9 控制台/后台服务项目，位于 `OperateExcel.Job`。
- 该任务使用 NPOI 将每日店铺导出文件导入 Excel 工作簿，刷新日报汇总工作表，并可按配置读取/写入飞书多维表格或知识库附件。
- 运行时配置分为 `OperateExcel.Job/appsettings.json` 和 `OperateExcel.Job/daily-report-profile.json`。
- 根目录下的 `.xlsx` 文件是样例或手工数据文件；除非任务明确要求处理工作簿数据，否则不要修改或重新生成它们。

## 构建与运行

- 还原并构建：`dotnet build OperateExcel.Job\OperateExcel.Job.csproj`
- 单次运行：`dotnet run --project OperateExcel.Job\OperateExcel.Job.csproj -- --run-once`
- 回填指定日期：`dotnet run --project OperateExcel.Job\OperateExcel.Job.csproj -- --run-once --date=yyyy-MM-dd`
- 验证时覆盖本地模板或输出目录：添加 `--template=...` 和 `--output-dir=...`。
- 省略 `--run-once` 时默认进入服务模式；运行计划由 `ExcelImport:DailyCron` 控制。

## 测试与验证

- 当前没有独立测试项目。代码变更后以 `dotnet build` 作为基础验证。
- 验证工作簿行为时，优先使用复制出的模板和输出目录执行单次冒烟测试，避免生成文件覆盖生产路径。
- 联网飞书流程需要有效凭据和外部网络访问；测试时根据appsettings.json的配置，填写指定日期测试一遍。

## 代码风格

- 遵循现有 C# 风格：文件范围命名空间、启用 nullable、隐式 using、适合时使用 sealed 类、用 record 表达不可变结果/数据结构，并通过清晰异常信息尽早校验。
- 将业务相关的工作簿常量、工作表名和表头名集中放在相关导入逻辑附近。
- 保留现有 Unicode 转义序列和非英文业务字符串。不要仅为改变显示编码而重写它们。
- 注释保持简短，只在工作簿或飞书行为不直观时解释原因。

## 配置与数据保护

- `appsettings.json` 可能包含环境特定路径和飞书凭据。除非用户要求，不要把真实值替换成占位符，也不要新增密钥。
- `daily-report-profile.json` 驱动人员、店铺、源目录别名、映射表名称、预算和 SKU 归属映射。生成汇总会使用其中顺序，因此需要保留顺序。
- 谨慎处理 `verify-output` 和任何已配置输出目录下的生成文件；只有任务需要时才删除或覆盖生成的工作簿。
- 修改工作簿导入逻辑前，先检查 `ExcelImportJob` 中目标表头的匹配方式，以及 `DelimitedTableReader` 中分隔文件的解析方式。
