<#
    sync-feed.ps1 - 抓取 Blackboard 日历源并保存为本地 .ics（保持更新）
    每天由计划任务自动运行；也可手动运行。
#>
[CmdletBinding()]
param(
    [string]$Url  = '',
    [string]$OutFile = ''
)

# $PSScriptRoot 在 param 默认值求值时可能为空（-File 调用），故在此推导
$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $SelfDir
if ([string]::IsNullOrEmpty($OutFile)) { $OutFile = Join-Path $ProjectRoot 'data\blackboard.ics' }

# 日历源地址属于个人凭据，放在 config\config.psd1 里，不写死在代码中。
if ([string]::IsNullOrEmpty($Url)) {
    $cfgPath = Join-Path $ProjectRoot 'config\config.psd1'
    if (Test-Path -LiteralPath $cfgPath) {
        try {
            $cfg = Import-PowerShellDataFile -LiteralPath $cfgPath
            if ($cfg.ContainsKey('FeedUrl')) { $Url = [string]$cfg['FeedUrl'] }
        } catch { }
    }
}
if ([string]::IsNullOrEmpty($Url)) {
    Write-Output '[sync] FAIL 未配置日历源：请在 config\config.psd1 里填写 FeedUrl'
    exit 2
}

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$dir = Split-Path -Parent $OutFile
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

try {
    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 60
    $text = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    if ($text -notmatch 'BEGIN:VCALENDAR') { throw "返回内容不是 iCalendar（可能链接已失效或需要登录）" }
    Set-Content -LiteralPath $OutFile -Value $text -Encoding UTF8
    $n = ([regex]::Matches($text, 'BEGIN:VEVENT')).Count
    Write-Output ("[sync] OK  " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "  事件数=" + $n + "  -> " + (Resolve-Path $OutFile).Path)
    exit 0
} catch {
    Write-Output ("[sync] FAIL " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "  " + $_.Exception.Message)
    Write-Output "[sync] 保留上一次成功抓取的本地副本，提醒仍可继续。"
    exit 1
}
