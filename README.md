# OperateExcel.Job

.NET 9 Web job for importing daily store files into `Temp.xlsx`.

## What it imports

- `B2B（ol)` sheet is filled first from the previous day's Feishu bitable attachment `损益表`.

- `fulfillment` sheet reads store `.txt` files.
- `payments` sheet reads store `.csv` files.
- `广告` sheet reads store `.xlsx` files.

The importer matches columns by the target sheet headers in `Temp.xlsx`. Source files may contain extra columns. If a target column is not found in the source file, that column is left empty.

## Date folder

By default the job uses the previous local date:

```text
D:\code\OperateExcelTemp\yyyy-MM-dd\yyyy-MM-dd
```

For example, when the current date is `2026-05-24`, the default source folder is:

```text
D:\code\OperateExcelTemp\2026-05-23\2026-05-23
```

Use `--date=yyyy-MM-dd` only for manual backfill.

## Feishu attachment import

Before importing `fulfillment`, `payments`, and `广告`, the job queries the configured Feishu bitable record whose `日期` field equals the processing date, validates that the `损益表` attachment exists and contains exactly one `.xlsx` file, downloads it, then copies the `B2BOL` sheet into `B2B（ol)` by matching column headers.

Fill these values in `OperateExcel.Job/appsettings.json` before running:

```json
"Feishu": {
  "Enabled": true,
  "AppId": "your-app-id",
  "AppSecret": "your-app-secret"
}
```

If the Feishu attachment is missing, not `.xlsx`, or has more than one file, the job logs `附件读取失败` and stops.

## Run once

```powershell
dotnet run --project OperateExcel.Job\OperateExcel.Job.csproj -- --run-once
```

Backfill the supplied sample date:

```powershell
dotnet run --project OperateExcel.Job\OperateExcel.Job.csproj -- --run-once --date=2026-05-22
```

## Hangfire

Set `ConnectionStrings:Hangfire` in `OperateExcel.Job/appsettings.json` to a SQL Server connection string. When configured, the app registers a recurring Hangfire job with `ExcelImport:DailyCron`, defaulting to `0 2 * * *`.

Hangfire dashboard is available at:

```text
http://localhost:5000/hangfire
```
