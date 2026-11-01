<#
    install-task.ps1 - 注册 Windows 计划任务：每天在指定时间（默认 08:00）运行 daily-reminder.ps1
    以当前用户身份运行（这样 DPAPI 凭据才可解密），无需存储 Windows 密码。
    需要以管理员身份运行 PowerShell 才能注册（注册本身需要权限，之后运行不需要）。
#>
[CmdletBinding()]
param(
    [string]$Time = '08:00',
    [string]$TaskName = 'HKCC Daily Reminder',
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot 'daily-reminder.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) { throw "找不到 $scriptPath" }

if ($Uninstall) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "已删除计划任务: $TaskName" -ForegroundColor Green
    } else {
        Write-Host "计划任务不存在: $TaskName" -ForegroundColor Yellow
    }
    exit 0
}

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$argLine = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $scriptPath + '"'

$action = New-ScheduledTaskAction -Execute $psExe -Argument $argLine -WorkingDirectory (Split-Path -Parent $PSScriptRoot)
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId ("$env:USERDOMAIN\$env:USERNAME") -LogonType Interactive -RunLevel Limited

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "已存在同名任务，先移除后重建。" -ForegroundColor Yellow
}

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'HKCC 每日定时汇总未来 7 天课程与作业截止，并发送邮件提醒。' | Out-Null

Write-Host ''
Write-Host "已注册计划任务: $TaskName" -ForegroundColor Green
Write-Host ("  运行时间: 每天 " + $Time)
Write-Host ("  执行命令: " + $psExe + " " + $argLine)
Write-Host '  注意: 到点时电脑需开机且当前用户已登录；默认不会唤醒睡眠中的电脑。' -ForegroundColor Yellow
Write-Host ''
Write-Host '如需立即测试（不等到设定时间）:' -ForegroundColor Cyan
Write-Host ('  Start-ScheduledTask -TaskName "' + $TaskName + '"')
Write-Host '查看运行结果:' -ForegroundColor Cyan
Write-Host ('  Get-ScheduledTaskInfo -TaskName "' + $TaskName + '" | Format-List')
Write-Host '删除任务:' -ForegroundColor Cyan
Write-Host ('  .\install-task.ps1 -Uninstall')
