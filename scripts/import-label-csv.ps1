<#
    import-label-csv.ps1 - 用 Excel 批量标注：读取 label-report.csv，把已填的 CourseCode 导回 uid-map.psd1

    流程：
      1) 打开 out\label-report.csv（Excel / 记事本均可）
      2) 在 CourseCode 列填入课程代号（如 CS101）；留空 = 暂不标
         可用代号见 config\courses.psd1；想新增课程先加进那里
      3) 保存（保持 CSV 格式、UTF-8）
      4) 运行本脚本 —— 它会校验代号、写入 uid-map.psd1、并重新生成日历与报表

    用法：
        powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\import-label-csv.ps1
        可选: -CsvPath <路径>  -DryRun
#>
[CmdletBinding()]
param(
    [string]$CsvPath = '',
    [switch]$DryRun
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Root = Split-Path -Parent $SelfDir

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

if ([string]::IsNullOrEmpty($CsvPath)) { $CsvPath = Join-Path $Root 'out\label-report.csv' }
$coursesPath = Join-Path $Root 'config\courses.psd1'
$uidMapPath  = Join-Path $Root 'config\uid-map.psd1'

if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Host "找不到 $CsvPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $coursesPath)) { Write-Host "找不到 $coursesPath" -ForegroundColor Red; exit 1 }

$cc = Import-PowerShellDataFile -LiteralPath $coursesPath
$validCodes = @()
if ($cc.ContainsKey('Courses') -and $cc['Courses']) { $validCodes = @($cc['Courses'].Keys) }

$uidMap = @{}
if (Test-Path -LiteralPath $uidMapPath) {
    $um = Import-PowerShellDataFile -LiteralPath $uidMapPath
    foreach ($k in $um.Keys) { $uidMap[[string]$k] = [string]$um[$k] }
}

$rows = Import-Csv -LiteralPath $CsvPath -Encoding UTF8
Write-Host ('读取 ' + $rows.Count + ' 行: ' + $CsvPath) -ForegroundColor Cyan

$added = 0; $changed = 0; $invalid = @(); $blank = 0
foreach ($r in $rows) {
    $uid = [string]$r.Uid
    if ([string]::IsNullOrWhiteSpace($uid)) { continue }
    $code = ([string]$r.CourseCode).Trim()
    if ([string]::IsNullOrWhiteSpace($code)) { $blank++; continue }
    if ($validCodes -notcontains $code) { $invalid += ($code + '  (行: ' + $r.Title + ')'); continue }
    if ($uidMap.ContainsKey($uid)) {
        if ($uidMap[$uid] -ne $code) { $uidMap[$uid] = $code; $changed++ }
    } else {
        $uidMap[$uid] = $code; $added++
    }
}

if ($invalid.Count -gt 0) {
    Write-Host ''
    Write-Host ('以下 CourseCode 不在 courses.psd1 的 Courses 里，已跳过（请先添加课程）:') -ForegroundColor Yellow
    $invalid | Select-Object -Unique | ForEach-Object { Write-Host ('  - ' + $_) -ForegroundColor Yellow }
    Write-Host ('  可用代号: ' + ($validCodes -join ', ')) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host ('新增 ' + $added + ' 条，更新 ' + $changed + ' 条，留空跳过 ' + $blank + ' 条，无效 ' + $invalid.Count + ' 条') -ForegroundColor Cyan

if ($DryRun) { Write-Host '[DryRun] 未写入。' -ForegroundColor Yellow; exit 0 }
if ($added -eq 0 -and $changed -eq 0) { Write-Host '没有需要写入的变更。' -ForegroundColor Yellow; exit 0 }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('# UID -> course code map, maintained by scripts\label-events.ps1')
[void]$sb.AppendLine('# Key   : Blackboard event UID (stable across feed refreshes)')
[void]$sb.AppendLine('# Value : course code from config\courses.psd1')
[void]$sb.AppendLine('# Hand-editing is fine too. ASCII only.')
[void]$sb.AppendLine('@{')
foreach ($k in ($uidMap.Keys | Sort-Object)) {
    [void]$sb.AppendLine("    '" + $k + "' = '" + $uidMap[$k] + "'")
}
[void]$sb.AppendLine('}')
Set-Content -LiteralPath $uidMapPath -Value $sb.ToString() -Encoding ASCII
Write-Host ('已写入 ' + $uidMapPath + '  (共 ' + $uidMap.Count + ' 条映射)') -ForegroundColor Green

Write-Host ''
Write-Host '重新生成日历与报表…' -ForegroundColor Cyan
& (Join-Path $Root 'lib\build-merged.ps1')
