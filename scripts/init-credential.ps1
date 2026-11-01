<#
    init-credential.ps1 - 【由你本人运行一次】
    交互式输入 SMTP 账号 / 密码（或应用专用密码 / 授权码），
    用 Windows DPAPI 加密后保存到 config\smtp.cred.xml。
    该文件只有本机、本 Windows 账户能解密；不会上传到任何地方。
#>
[CmdletBinding()]
param(
    [string]$CredPath = ''
)

# $PSScriptRoot 在 param 默认值求值时可能为空（-File 调用），故在此推导
$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrEmpty($CredPath)) { $CredPath = Join-Path (Split-Path -Parent $SelfDir) 'config\smtp.cred.xml' }

Write-Host ''
Write-Host '=== HKCC 提醒 - SMTP 凭据初始化 ===' -ForegroundColor Cyan
Write-Host '提示：多数邮箱需要"应用专用密码/授权码"，不是网页登录密码。' -ForegroundColor Yellow
Write-Host '      Gmail 开启两步验证后在账号设置里生成；QQ/163 在邮箱设置里生成授权码。'
Write-Host ''
Write-Host '接下来会弹出 Windows 凭据窗口，请输入：' -ForegroundColor Cyan
Write-Host '  用户名 = 你的邮箱地址（必须与 config.psd1 的 From 一致）'
Write-Host '  密码   = 应用专用密码 / 授权码'
Write-Host ''

$cred = Get-Credential -Message 'SMTP 账号（邮箱地址）与应用专用密码'
if ($null -eq $cred) { Write-Host '已取消。' -ForegroundColor Yellow; exit 1 }

$dir = Split-Path -Parent $CredPath
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

$cred | Export-Clixml -LiteralPath $CredPath
Write-Host ''
Write-Host ('已加密保存: ' + (Resolve-Path $CredPath).Path) -ForegroundColor Green
Write-Host '确认 config\config.psd1 中的 From 与上面输入的用户名一致。' -ForegroundColor Yellow
Write-Host '随后可运行 scripts\daily-reminder.ps1 测试发送。' -ForegroundColor Green
