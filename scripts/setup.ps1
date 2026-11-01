<#
    setup.ps1 - 一次性配置向导
      1) 问你发件邮箱，自动推荐 SMTP 服务器（可覆盖）
      2) 让你自己输入应用专用密码/授权码（Get-Credential，我不接触）
      3) DPAPI 加密保存到 config\smtp.cred.xml
      4) 写 config\config.psd1（纯 ASCII，避免 PS5.1 编码坑）
      5) 立即发一封测试邮件，确认整条链路可用

    运行：  powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup.ps1
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = '',
    [string]$CredPath   = '',
    [switch]$SkipTestMail
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $SelfDir
if ([string]::IsNullOrEmpty($ConfigPath)) { $ConfigPath = Join-Path $ProjectRoot 'config\config.psd1' }
if ([string]::IsNullOrEmpty($CredPath))   { $CredPath   = Join-Path $ProjectRoot 'config\smtp.cred.xml' }
$HkOffset = 8

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

function Read-Default {
    param([string]$Prompt, [string]$Default)
    $shown = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
    $ans = Read-Host $shown
    if ([string]::IsNullOrWhiteSpace($ans)) { return $Default }
    return $ans.Trim()
}

Write-Host ''
Write-Host '=== HKCC 每日提醒 - 配置向导 ===' -ForegroundColor Cyan
Write-Host ''

# ---------- 1. 邮箱与 SMTP ----------
$from = Read-Default -Prompt '发件邮箱地址' -Default ''
if ([string]::IsNullOrWhiteSpace($from)) { Write-Host '必须填写发件邮箱。' -ForegroundColor Red; exit 1 }

$domain = ($from -split '@')[-1].ToLower()
$guessServer = 'smtp.gmail.com'; $guessPort = 587
switch -Regex ($domain) {
    'gmail\.com|googlemail\.com' { $guessServer = 'smtp.gmail.com';     $guessPort = 587 }
    'outlook\.|hotmail\.|live\.|office365' { $guessServer = 'smtp.office365.com'; $guessPort = 587 }
    'qq\.com'   { $guessServer = 'smtp.qq.com';  $guessPort = 587 }
    '163\.com'  { $guessServer = 'smtp.163.com'; $guessPort = 465 }
    '126\.com'  { $guessServer = 'smtp.126.com'; $guessPort = 465 }
    'polyu\.edu\.hk|cpce-polyu\.edu\.hk' { $guessServer = 'smtp.office365.com'; $guessPort = 587 }
    default     { $guessServer = 'smtp.' + $domain; $guessPort = 587 }
}
Write-Host ("  推测 SMTP: " + $guessServer + " : " + $guessPort + "  (可直接回车接受)") -ForegroundColor DarkGray
$smtpServer = Read-Default -Prompt 'SMTP 服务器' -Default $guessServer
$smtpPort   = [int](Read-Default -Prompt 'SMTP 端口' -Default ([string]$guessPort))

$toRaw = Read-Default -Prompt '收件邮箱（多个用逗号分隔）' -Default $from
$toList = @($toRaw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# Blackboard 日历源（可留空，之后也能改 config\config.psd1）
$feedUrl = Read-Default -Prompt 'Blackboard 日历源 URL（可留空）' -Default ''

# ---------- 2. 凭据 ----------
Write-Host ''
Write-Host '接下来弹出 Windows 凭据窗口：' -ForegroundColor Yellow
Write-Host '  用户名 = 上面的发件邮箱'
Write-Host '  密码   = 应用专用密码 / 授权码（不是网页登录密码）'
Write-Host '           Gmail: 需先开两步验证 -> 应用专用密码'
Write-Host '           QQ/163: 邮箱设置 -> 生成授权码'
Write-Host ''
$cred = Get-Credential -UserName $from -Message 'SMTP 账号与应用专用密码'
if ($null -eq $cred) { Write-Host '已取消。' -ForegroundColor Yellow; exit 1 }

$credDir = Split-Path -Parent $CredPath
if (-not (Test-Path -LiteralPath $credDir)) { New-Item -ItemType Directory -Path $credDir -Force | Out-Null }
$cred | Export-Clixml -LiteralPath $CredPath
Write-Host ("凭据已加密保存: " + $CredPath) -ForegroundColor Green

# ---------- 3. 写配置（纯 ASCII） ----------
$cfgDir = Split-Path -Parent $ConfigPath
if (-not (Test-Path -LiteralPath $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null }
$toLiteral = ($toList | ForEach-Object { "'" + $_ + "'" }) -join ', '
$cfgText = @"
# HKCC daily reminder - configuration (ASCII only)
@{
    SmtpServer    = '$smtpServer'
    SmtpPort      = $smtpPort
    UseSsl        = `$true
    From          = '$from'
    To            = @($toLiteral)
    FeedUrl       = '$feedUrl'
    SubjectPrefix = '[HKCC]'
    Days          = 7
}
"@
Set-Content -LiteralPath $ConfigPath -Value $cfgText -Encoding ASCII
Write-Host ("配置已写入: " + $ConfigPath) -ForegroundColor Green

# 验证可解析
try {
    $chk = Import-PowerShellDataFile -LiteralPath $ConfigPath
    Write-Host ("  配置校验通过: From=" + $chk['From'] + " Server=" + $chk['SmtpServer'] + ":" + $chk['SmtpPort']) -ForegroundColor Green
} catch {
    Write-Host ("  配置校验失败: " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}

if ($SkipTestMail) { Write-Host '已跳过测试邮件。' -ForegroundColor Yellow; exit 0 }

# ---------- 4. 发测试邮件 ----------
Write-Host ''
Write-Host '发送测试邮件…' -ForegroundColor Cyan
try {
    $nowLocal = [datetime]::UtcNow.AddHours($HkOffset)
    Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl `
        -Credential $cred -From $from -To $toList `
        -Subject '[HKCC] 测试邮件 - 提醒已就绪' `
        -Body ("配置成功。`n`n你的每日提醒将在计划任务设定的时间发送（默认 08:00），内容为未来 7 天的课程与作业截止。`n配置时间: " + $nowLocal.ToString('yyyy-MM-dd HH:mm') + " (UTC+8)`n`n如果没有收到，请检查垃圾邮件文件夹。") `
        -Encoding ([System.Text.Encoding]::UTF8)
    Write-Host '测试邮件已发送，请查收（含垃圾邮件箱）。' -ForegroundColor Green
} catch {
    Write-Host ('测试邮件发送失败: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host '常见原因：端口/服务器不对、用了登录密码而非授权码、邮箱未开启 SMTP 服务。' -ForegroundColor Yellow
    exit 1
}

Write-Host ''
Write-Host '配置完成。计划任务 "HKCC Daily Reminder" 将在你设定的时间自动运行（默认 08:00）。' -ForegroundColor Green
Write-Host '立即试跑: Start-ScheduledTask -TaskName "HKCC Daily Reminder"' -ForegroundColor Cyan
