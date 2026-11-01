<#
    test-mail.ps1 - 只做一件事：读配置 + 读凭据，发一封测试邮件，并给出明确的成败诊断。
    不占用控制台输入（避免在 DSH 服务控制台里跑时卡住）。
    退出码：0 发送成功；2 缺凭据；1 发送失败
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = '',
    [string]$CredPath   = ''
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $SelfDir
if ([string]::IsNullOrEmpty($ConfigPath)) { $ConfigPath = Join-Path $ProjectRoot 'config\config.psd1' }
if ([string]::IsNullOrEmpty($CredPath))   { $CredPath   = Join-Path $ProjectRoot 'config\smtp.cred.xml' }

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$cfg = Import-PowerShellDataFile -LiteralPath $ConfigPath
Write-Host ('配置: ' + $cfg['From'] + ' -> ' + ($cfg['To'] -join ', ')) -ForegroundColor Cyan
Write-Host ('SMTP: ' + $cfg['SmtpServer'] + ':' + $cfg['SmtpPort'] + '  SSL=' + $cfg['UseSsl']) -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $CredPath)) {
    Write-Host ('缺凭据文件: ' + $CredPath) -ForegroundColor Red
    Write-Host '请先运行: powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\init-credential.ps1' -ForegroundColor Yellow
    exit 2
}
$cred = Import-Clixml -LiteralPath $CredPath
Write-Host ('凭据账号: ' + $cred.UserName + '  (密码已由 DPAPI 解密，长度 ' + $cred.GetNetworkCredential().Password.Length + ')') -ForegroundColor DarkGray

if ($cred.UserName -ne $cfg['From']) {
    Write-Host ('警告: 凭据账号与配置 From 不一致 —— 很可能发送失败。') -ForegroundColor Yellow
}

$now = [datetime]::UtcNow.AddHours(8)
$body = @"
这是 HKCC 每日提醒的测试邮件。

如果你收到这封邮件，说明整条链路已打通：
  - 每天在计划任务设定的时间自动运行（默认 08:00）
  - 内容为未来 7 天的课程与作业截止
  - 同时自动刷新 out\HKCC_merged.ics

配置时间: $($now.ToString('yyyy-MM-dd HH:mm')) (UTC+8)
发件账号: $($cfg['From'])
SMTP: $($cfg['SmtpServer']):$($cfg['SmtpPort'])

没收到后续每日邮件时，先看 logs\reminder-YYYY-MM.log。
"@

Write-Host '正在发送…' -ForegroundColor Cyan
try {
    # 注意：-UseSsl 是 switch 型参数，只能写成 -UseSsl:$true，不能传位置值
    $mailArgs = @{
        SmtpServer  = $cfg['SmtpServer']
        Port        = [int]$cfg['SmtpPort']
        Credential  = $cred
        From        = $cfg['From']
        To          = $cfg['To']
        Subject     = '[HKCC] 测试邮件 - 提醒链路验证'
        Body        = $body
        Encoding    = [System.Text.Encoding]::UTF8
        ErrorAction = 'Stop'
    }
    if ([bool]$cfg['UseSsl']) { $mailArgs['UseSsl'] = $true }
    Send-MailMessage @mailArgs
    Write-Host '✅ 测试邮件已发送成功，请查收（也看看垃圾邮件箱）。' -ForegroundColor Green
    exit 0
} catch {
    Write-Host '❌ 发送失败' -ForegroundColor Red
    Write-Host ('   错误: ' + $_.Exception.Message) -ForegroundColor Red
    $inner = $_.Exception.InnerException
    while ($inner) { Write-Host ('   内层: ' + $inner.Message) -ForegroundColor DarkRed; $inner = $inner.InnerException }
    Write-Host ''
    Write-Host '常见原因与对策:' -ForegroundColor Yellow
    Write-Host '  1) 密码用错：必须用「应用专用密码/授权码」，不是网页登录密码。'
    Write-Host '  2) 机构封禁基本认证（Office365 Security Defaults / SMTP AUTH 被关）：'
    Write-Host '     此时本方案无法发信，需改用 Gmail/QQ 账号发件，或改用 Microsoft Graph OAuth。'
    Write-Host '  3) 账号或密码输错：重跑 init-credential.ps1。'
    exit 1
}
