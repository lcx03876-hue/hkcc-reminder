<#
    mail-dump.ps1 - 手动读取收件箱并导出/预览（用于验证登录是否真的能读邮件）

    用法：
      powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\mail-dump.ps1
      powershell ... -File .\scripts\mail-dump.ps1 -Hours 72 -Max 50 -Export

    -Export 会把结果写成 out\mail-YYYY-MM-DD.json，便于我再帮你分析。
#>
[CmdletBinding()]
param(
    [int]$Hours = 24,
    [int]$Max = 40,
    [switch]$Export,
    [switch]$Raw
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Root = Split-Path -Parent $SelfDir

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

. (Join-Path $Root 'lib\graph-auth.ps1')
. (Join-Path $Root 'lib\mail-core.ps1')

$configPath = Join-Path $Root 'config\config.psd1'
$credPath   = Join-Path $Root 'config\graph.cred.xml'

try {
    $token = Get-GraphAccessToken -ConfigPath $configPath -CredPath $credPath
} catch {
    Write-Host ('✗ ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
if ([string]::IsNullOrEmpty($token)) {
    Write-Host '✗ 未登录或凭据失效，请先运行 scripts\mail-login.ps1' -ForegroundColor Red
    exit 2
}

Write-Host ("读取最近 $Hours 小时收件箱（最多 $Max 封）…") -ForegroundColor Cyan
try {
    $msgs = Get-MailMessages -Hours $Hours -Max $Max -Token $token
} catch {
    Write-Host ('✗ 读取失败: ' + $_.Exception.Message) -ForegroundColor Red
    try { Write-Host ('  ' + $_.ErrorDetails.Message) -ForegroundColor DarkRed } catch {}
    exit 1
}

Write-Host ("✓ 共 " + @($msgs).Count + " 封") -ForegroundColor Green

if ($Export) {
    $outDir = Join-Path $Root 'out'
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $outFile = Join-Path $outDir ('mail-' + (Get-Date -Format 'yyyy-MM-dd') + '.json')
    $msgs | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $outFile -Encoding UTF8
    Write-Host ('已导出: ' + $outFile) -ForegroundColor Green
}

if ($Raw) {
    $msgs | Format-List Subject, From, FromAddress, Received, Preview | Out-String -Width 200 | Write-Output
    exit 0
}

$summary = Get-MailSummary -Messages $msgs
Write-Host ''
Write-Host ('分类: 需行动 ' + @($summary.Action).Count + ' | 疑似调课 ' + @($summary.Change).Count + ' | 通知 ' + @($summary.Notice).Count + ' | 其他 ' + @($summary.Other).Count) -ForegroundColor Cyan
Write-Host ''
Write-Host '--- 需行动 ---' -ForegroundColor Yellow
foreach ($m in $summary.Action) { Write-Host ('  ' + $m.Received.ToString('MM-dd HH:mm') + '  [' + $m.From + ']  ' + $m.Subject) }
Write-Host ''
Write-Host '--- 疑似调课/改期（待人工确认，不会自动改课表）---' -ForegroundColor Yellow
foreach ($c in $summary.Change) { Write-Host ('  ' + $c.Msg.Received.ToString('MM-dd HH:mm') + '  [' + $c.Msg.From + ']  ' + $c.Msg.Subject + '   (' + $c.Reason + ')') }
Write-Host ''
Write-Host '--- 通知 ---' -ForegroundColor Yellow
foreach ($m in $summary.Notice) { Write-Host ('  ' + $m.Received.ToString('MM-dd HH:mm') + '  [' + $m.From + ']  ' + $m.Subject) }
Write-Host ''
Write-Host '--- 其他 ---' -ForegroundColor Yellow
foreach ($m in $summary.Other) { Write-Host ('  ' + $m.Received.ToString('MM-dd HH:mm') + '  [' + $m.From + ']  ' + $m.Subject) }
