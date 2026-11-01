[CmdletBinding()]
param(
    [string]$Time = '',
    [string]$TaskName = 'HKCC Daily Reminder',
    [string]$ExpectedUser = '',
    [switch]$Reconfigure,
    [switch]$SkipPreview,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
Set-Location -LiteralPath $Root

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $Host.UI.RawUI.WindowTitle = 'HKCC 提醒 - 安装'
} catch { }

function Write-Step {
    param([string]$Text)
    Write-Host ''
    Write-Host ('== ' + $Text + ' ==') -ForegroundColor Cyan
}

function Stop-Install {
    param([string]$Message)
    Write-Host ''
    Write-Host ('[失败] ' + $Message) -ForegroundColor Red
    if (-not $NoPause) { Read-Host '按回车退出' }
    exit 1
}

Write-Host ''
Write-Host '  HKCC / 校园日程提醒 · 一键安装' -ForegroundColor Green
Write-Host '  ----------------------------------------' -ForegroundColor DarkGray
Write-Host ('  安装目录: ' + $Root) -ForegroundColor DarkGray

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Stop-Install '需要 Windows PowerShell 5.1 或更高版本。'
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Stop-Install '需要管理员权限来注册计划任务。请双击 install.bat，它会自动申请提升。'
}
if (-not [string]::IsNullOrEmpty($ExpectedUser) -and $ExpectedUser -ne $env:USERNAME) {
    Stop-Install ('当前是另一个管理员账号（' + $env:USERNAME + '），不是 ' + $ExpectedUser +
        '。DPAPI 凭据会绑定到错误的账号，请用你平时登录的 Windows 账号运行。')
}

# ---------- 1. 检查文件 ----------
Write-Step '检查安装文件'
$requiredFiles = @(
    'scripts\setup.ps1',
    'scripts\install-task.ps1',
    'scripts\daily-reminder.ps1',
    'lib\build-merged.ps1',
    'config\config.example.psd1'
)
foreach ($rel in $requiredFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $Root $rel))) {
        Stop-Install ('安装包不完整，缺少: ' + $rel)
    }
}

$configDir = Join-Path $Root 'config'
$dataDir = Join-Path $Root 'data'
if (-not (Test-Path -LiteralPath $configDir)) { New-Item -ItemType Directory -Path $configDir -Force | Out-Null }
if (-not (Test-Path -LiteralPath $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

$exampleMap = @(
    @{ Src = 'courses.example.psd1';          Dst = 'courses.psd1';          Label = '课程目录' },
    @{ Src = 'extra-deadlines.example.psd1';  Dst = 'extra-deadlines.psd1';  Label = '额外截止' },
    @{ Src = 'overrides.example.psd1';        Dst = 'overrides.psd1';        Label = '调课覆盖' },
    @{ Src = 'uid-map.example.psd1';          Dst = 'uid-map.psd1';          Label = 'UID 映射' }
)
foreach ($m in $exampleMap) {
    $src = Join-Path $configDir $m.Src
    $dst = Join-Path $configDir $m.Dst
    if (-not (Test-Path -LiteralPath $dst)) {
        if (-not (Test-Path -LiteralPath $src)) { Stop-Install ('缺少示例配置: ' + $m.Src) }
        Copy-Item -LiteralPath $src -Destination $dst
        Write-Host ('  已创建 config\' + $m.Dst + '（' + $m.Label + '）') -ForegroundColor DarkGray
    }
}

$timetable = Join-Path $dataDir 'timetable.ics'
$sampleTimetable = Join-Path $dataDir 'timetable.sample.ics'
if (-not (Test-Path -LiteralPath $timetable)) {
    if (-not (Test-Path -LiteralPath $sampleTimetable)) { Stop-Install '缺少 data\timetable.sample.ics' }
    Copy-Item -LiteralPath $sampleTimetable -Destination $timetable
    Write-Host '  已创建 data\timetable.ics（示例课表，之后请替换）' -ForegroundColor DarkGray
}

# ---------- 2. 配置邮箱 / 日历源 ----------
$configPath = Join-Path $configDir 'config.psd1'
$credPath = Join-Path $configDir 'smtp.cred.xml'
$needSetup = $true
if (-not $Reconfigure -and (Test-Path -LiteralPath $configPath) -and (Test-Path -LiteralPath $credPath)) {
    try {
        $cfg = Import-PowerShellDataFile -LiteralPath $configPath
        $from = [string]$cfg['From']
        if (-not [string]::IsNullOrWhiteSpace($from) -and $from -notmatch '@example\.com$') {
            $needSetup = $false
        }
    } catch {
        $needSetup = $true
    }
}

if (-not $needSetup) {
    Write-Host ''
    Write-Host '检测到已有配置（config.psd1 + smtp.cred.xml）。' -ForegroundColor DarkGray
    $redo = Read-Host '要重新配置邮箱/日历源吗？(y/N)'
    if ($redo -match '^[yY]') { $needSetup = $true }
}

if ($needSetup) {
    Write-Step '配置邮箱与日历源'
    Write-Host '接下来会询问发件邮箱、SMTP、收件人、Blackboard 日历源，' -ForegroundColor Gray
    Write-Host '然后弹出密码窗口，请输入应用专用密码 / 授权码（不是网页登录密码）。' -ForegroundColor Gray
    try {
        $global:LASTEXITCODE = 0
        & (Join-Path $Root 'scripts\setup.ps1')
        if ($LASTEXITCODE -ne 0) { throw 'setup.ps1 返回失败' }
    } catch {
        Stop-Install ('邮箱/日历源配置未完成: ' + $_.Exception.Message)
    }
} else {
    Write-Host ''
    Write-Host '使用现有配置，跳过邮箱向导。' -ForegroundColor DarkGray
}

if (-not (Test-Path -LiteralPath $configPath)) {
    Stop-Install '没有生成 config\config.psd1，无法继续。'
}
try {
    $cfg = Import-PowerShellDataFile -LiteralPath $configPath
} catch {
    Stop-Install ('config\config.psd1 解析失败: ' + $_.Exception.Message)
}
if ([string]::IsNullOrWhiteSpace([string]$cfg['From'])) {
    Stop-Install 'config.psd1 里的 From（发件邮箱）为空。'
}
if ([string]::IsNullOrWhiteSpace([string]$cfg['FeedUrl'])) {
    Write-Host ''
    Write-Host '提示: FeedUrl 为空，不会自动抓取 Blackboard 日历源；' -ForegroundColor Yellow
    Write-Host '      daily-reminder.ps1 会使用本地 data\blackboard.ics（如果存在）。' -ForegroundColor Yellow
}

# ---------- 3. 提醒时间 ----------
if ([string]::IsNullOrWhiteSpace($Time)) {
    $defaultTime = '08:00'
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $existingTask -and @($existingTask.Triggers).Count -gt 0) {
        try {
            $defaultTime = ([datetime]$existingTask.Triggers[0].StartBoundary).ToString('HH:mm')
        } catch { }
    }
    $answer = Read-Host ('每天几点发送提醒？(HH:mm，直接回车 = ' + $defaultTime + ')')
    if ([string]::IsNullOrWhiteSpace($answer)) { $Time = $defaultTime } else { $Time = $answer.Trim() }
}
if ($Time -notmatch '^([01]\d|2[0-3]):[0-5]\d$') {
    Stop-Install ('提醒时间格式应为 HH:mm，例如 08:00 或 19:30；当前是: ' + $Time)
}

# ---------- 4. 注册计划任务 ----------
Write-Step '注册计划任务'
Write-Host ('  任务名: ' + $TaskName) -ForegroundColor Gray
Write-Host ('  每天运行: ' + $Time) -ForegroundColor Gray
try {
    & (Join-Path $Root 'scripts\install-task.ps1') -Time $Time -TaskName $TaskName
} catch {
    Stop-Install ('计划任务注册失败: ' + $_.Exception.Message)
}
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) {
    Stop-Install '注册后没有找到计划任务，请查看上面的错误信息。'
}
Write-Host '  计划任务已注册。' -ForegroundColor Green

# ---------- 5. 可选试跑 ----------
if (-not $SkipPreview) {
    Write-Step '试跑一次（可选）'
    $runNow = Read-Host '现在预览一次「今天起 7 天」的摘要吗？(Y/n)'
    if ($runNow -notmatch '^[nN]') {
        try {
            & (Join-Path $Root 'scripts\daily-reminder.ps1') -Preview
        } catch {
            Write-Host ('预览失败（不影响安装）: ' + $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# ---------- 6. 完成 ----------
Write-Host ''
Write-Host '安装完成！' -ForegroundColor Green
Write-Host ('计划任务 "' + $TaskName + '" 每天 ' + $Time + ' 运行。') -ForegroundColor Gray
Write-Host ''
Write-Host '⚠ 到点时电脑必须开机、联网，并且当前 Windows 用户已登录；睡眠不会自动唤醒。' -ForegroundColor Yellow
Write-Host ''
Write-Host '下一步：' -ForegroundColor Cyan
Write-Host '  1) 编辑 data\timetable.ics，换成你的课表（参考 data\timetable.sample.ics）'
Write-Host '  2) 编辑 config\courses.psd1，加入你的课程代号'
Write-Host '  3) 详细说明见 README.md / INSTALL.txt'
Write-Host ''
if (-not $NoPause) { Read-Host '按回车退出' }
