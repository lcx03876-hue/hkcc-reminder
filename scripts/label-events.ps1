<#
    label-events.ps1 - 给 Blackboard 事件标注科目（只需做一次，之后每天自动套用）

    原理：Blackboard 的事件 UID 是稳定的（同一项作业永远同一个 UID），
          把 UID -> 课程代号 记进 config\uid-map.ps1，以后抓取时自动带上科目标签。
          63 条事件去重后通常只有 ~40 项需要标注。

    用法（建议另开一个 PowerShell 窗口，不要在 DSH 服务控制台里跑）：
        powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\label-events.ps1
        可选参数:
          -All            连已标注的也重新过一遍
          -OnlyUnlabeled  只处理未标注的（默认行为）
          -DryRun         只显示将要标注的内容，不写入
#>
[CmdletBinding()]
param(
    [switch]$All,
    [switch]$OnlyUnlabeled = $true,
    [switch]$DryRun
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Root = Split-Path -Parent $SelfDir
$HkOffset = 8

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

. (Join-Path $Root 'lib\ics.ps1')

$feedPath   = Join-Path $Root 'data\blackboard.ics'
$coursesPath = Join-Path $Root 'config\courses.psd1'
$uidMapPath = Join-Path $Root 'config\uid-map.psd1'

if (-not (Test-Path -LiteralPath $feedPath)) { Write-Host "找不到 $feedPath，请先运行 sync-feed.ps1 抓取。" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $coursesPath)) { Write-Host "找不到 $coursesPath" -ForegroundColor Red; exit 1 }

$cc = Import-PowerShellDataFile -LiteralPath $coursesPath
$courses = @{}
if ($cc.ContainsKey('Courses') -and $cc['Courses']) {
    foreach ($k in $cc['Courses'].Keys) { $courses[[string]$k] = [string]$cc['Courses'][$k] }
}
if ($courses.Count -eq 0) { Write-Host 'courses.psd1 里 Courses 为空，请先填入你的课程。' -ForegroundColor Red; exit 1 }

$uidMap = @{}
if (Test-Path -LiteralPath $uidMapPath) {
    $um = Import-PowerShellDataFile -LiteralPath $uidMapPath
    foreach ($k in $um.Keys) { $uidMap[[string]$k] = [string]$um[$k] }
}

$events = Read-IcsEvents -Lines (Get-Content -LiteralPath $feedPath -Encoding UTF8)
# 去重：同一 UID 只问一次
$unique = @{}
foreach ($e in $events) {
    if (-not $e.Uid) { continue }
    if (-not $unique.ContainsKey($e.Uid)) { $unique[$e.Uid] = $e }
}
$list = @($unique.Values | Sort-Object DtStartUtc)

$todo = @()
foreach ($e in $list) {
    $has = $uidMap.ContainsKey($e.Uid)
    if ($All -or -not $has) { $todo += $e }
}

Write-Host ''
Write-Host ('=== 科目标注 ===') -ForegroundColor Cyan
Write-Host ('源事件 ' + $events.Count + ' 条，去重后 ' + $list.Count + ' 项；本次待处理 ' + $todo.Count + ' 项') -ForegroundColor Cyan
if ($todo.Count -eq 0) { Write-Host '全部已标注，无需处理。' -ForegroundColor Green; exit 0 }
Write-Host ''

# 菜单
$menu = @()
$i = 0
foreach ($k in $courses.Keys) { $menu += $k; Write-Host ("  [" + $i + "] " + $k + '  ' + $courses[$k]) -ForegroundColor DarkGray; $i++ }
Write-Host ("  [s] 跳过此项") -ForegroundColor DarkGray
Write-Host ("  [q] 保存并退出") -ForegroundColor DarkGray
Write-Host ''

$changed = 0
$skipped = 0

function Save-UidMap {
    param($Map, $Path)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('# UID -> course code map, maintained by scripts\label-events.ps1')
    [void]$sb.AppendLine('# Key   : Blackboard event UID (stable across feed refreshes)')
    [void]$sb.AppendLine('# Value : course code from config\courses.psd1')
    [void]$sb.AppendLine('# Hand-editing is fine too. ASCII only.')
    [void]$sb.AppendLine('@{')
    foreach ($k in ($Map.Keys | Sort-Object)) {
        [void]$sb.AppendLine("    '" + $k + "' = '" + $Map[$k] + "'")
    }
    [void]$sb.AppendLine('}')
    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding ASCII
}

foreach ($e in $todo) {
    $when = $e.DtStartUtc.AddHours($HkOffset)
    Write-Host '--------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ('时间: ' + $when.ToString('yyyy-MM-dd HH:mm')) -ForegroundColor White
    Write-Host ('标题: ' + $e.Summary) -ForegroundColor Yellow
    if ($uidMap.ContainsKey($e.Uid)) { Write-Host ('当前标注: ' + $uidMap[$e.Uid]) -ForegroundColor DarkGray }
    $ans = Read-Host '选择课程编号 (s=跳过, q=退出并保存)'

    if ($ans -eq 'q') { break }
    if ($ans -eq 's' -or [string]::IsNullOrWhiteSpace($ans)) { $skipped++; continue }

    $idx = -1
    if (-not [int]::TryParse($ans, [ref]$idx) -or $idx -lt 0 -or $idx -ge $menu.Count) {
        Write-Host '输入无效，已跳过。' -ForegroundColor Red
        $skipped++
        continue
    }
    $uidMap[$e.Uid] = $menu[$idx]
    $changed++
    # 每标一条立即存盘，中途关窗口也不会丢
    if (-not $DryRun) { Save-UidMap -Map $uidMap -Path $uidMapPath }
    Write-Host ('  -> 已记为 ' + $menu[$idx] + '  (已存盘)') -ForegroundColor Green
}

if ($DryRun) {
    Write-Host ''
    Write-Host ('[DryRun] 本应写入 ' + $changed + ' 条，未保存。') -ForegroundColor Yellow
    exit 0
}

Save-UidMap -Map $uidMap -Path $uidMapPath

Write-Host ''
Write-Host ('已保存 ' + $changed + ' 条标注（跳过 ' + $skipped + ' 项，共 ' + $uidMap.Count + ' 条映射）') -ForegroundColor Green
Write-Host ('文件: ' + $uidMapPath) -ForegroundColor Green
Write-Host ''
Write-Host '下一步：重新生成日历与邮件内容' -ForegroundColor Cyan
Write-Host '  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\daily-reminder.ps1 -NoMail'
