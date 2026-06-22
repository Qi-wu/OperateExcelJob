# OperateExcel.Job

.NET 9 Web job for importing daily store files into `Temp.xlsx`.

## What it imports

- `B2B（ol)` sheet is filled first from the processing date folder's Excel file whose name contains `产品信息下载_基础`.

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

Before importing `fulfillment`, `payments`, and `广告`, the job reads the `产品信息下载_基础` Excel file from the processing date folder and copies its first sheet into `B2B（ol)` by matching column headers.

The `映射表` sheet is filled from the configured Feishu spreadsheet sheets instead of the bitable `映射表` attachment. The importer reads the four sheet URLs in `MappingSpreadsheetSheetUrls`, matches each sheet by its title using the original mapping-sheet name rules, copies `平台SKU`, `Asin`, `B2B Item Code`, and `运营`, then fills the target `账号` column with the original account values.

## SKU owner mappings

The `sku归属` sheet is refreshed from `OperateExcel.Job/sku-owner-mappings.json` on every run. It is no longer inherited from the previous day's daily report attachment.

Maintain the mapping file like this:

```json
{
  "mappings": [
    {
      "运营": "付江爽",
      "姓名": "付江爽"
    }
  ]
}
```

`运营` must match the `运营` value imported into the `映射表` sheet. `姓名` must match the people used by the daily summary, such as `付江爽`, `龙杨`, `DD`, or `李慧`.

The mapping file path is controlled by `ExcelImport:SkuOwnerMappingFilePath` in `OperateExcel.Job/appsettings.json`. You can also override it for a manual run:

```powershell
dotnet run --project OperateExcel.Job\OperateExcel.Job.csproj -- --run-once --sku-owner-mapping=D:\path\sku-owner-mappings.json
```

Fill these values in `OperateExcel.Job/appsettings.json` before running:

```json
"Feishu": {
  "Enabled": true,
  "AppId": "your-app-id",
  "AppSecret": "your-app-secret",
  "MappingSpreadsheetUrl": "https://bcnt3e3uyrxk.feishu.cn/wiki/NYS8wnqv1i2oKwkRXADc0IQ7nrd",
  "MappingSpreadsheetSheetUrls": [
    "https://bcnt3e3uyrxk.feishu.cn/wiki/NYS8wnqv1i2oKwkRXADc0IQ7nrd?sheet=79edea",
    "https://bcnt3e3uyrxk.feishu.cn/wiki/NYS8wnqv1i2oKwkRXADc0IQ7nrd?sheet=JLQ3Ie",
    "https://bcnt3e3uyrxk.feishu.cn/wiki/NYS8wnqv1i2oKwkRXADc0IQ7nrd?sheet=PKwiQn",
    "https://bcnt3e3uyrxk.feishu.cn/wiki/NYS8wnqv1i2oKwkRXADc0IQ7nrd?sheet=G8R0HR"
  ]
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

## Windows Service

The job runs as a .NET background service and can be registered as a Windows service. The schedule is controlled by `ExcelImport:DailyCron`, defaulting to `0 2 * * *` (02:00 local time every day). Only daily cron expressions in the form `minute hour * * *` are supported.

Publish the service:

```powershell
dotnet publish OperateExcel.Job\OperateExcel.Job.csproj -c Release -o C:\OperateExcelJob
```

Register it from an elevated PowerShell prompt:

```powershell
New-Service -Name OperateExcelJob -BinaryPathName "C:\OperateExcelJob\OperateExcel.Job.exe" -DisplayName "OperateExcelJob" -StartupType Automatic
Start-Service OperateExcelJob
```

Stop and remove it:

```powershell
Stop-Service OperateExcelJob
sc.exe delete OperateExcelJob
```

## Local logs

Runtime logs are written to daily files under the configured `FileLog:Directory`. With the default config, logs are stored beside the published executable:

```text
C:\OperateExcelJob\logs\operate-excel-yyyyMMdd.log
```
