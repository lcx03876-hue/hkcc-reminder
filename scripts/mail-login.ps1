<#
    mail-login.ps1 - 一次性登录：用设备码方式授权读邮件，并把 refresh token 缓存到本机。

    运行后：
      1) 屏幕会显示一个代码（如 ABCD-EFGH）
      2) 浏览器打开 aka.ms/devicelogin，输入该代码，用你的 CPCE 账号登录
      3) 授权成功后 token 加密存到 config\graph.cred.xml
      4) 之后每日任务可静默刷新，无需再次登录

    注意：需要先在 config\config.psd1 里填好 ClientId（Azure 应用注册的"应用程序(客户端) ID"）
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = '',
    [string]$CredPath   = '',
    [switch]$Status     # 只查看当前登录状态，不重新登录
)

$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Root = Split-Path -Parent $SelfDir
if ([string]::IsNullOrEmpty($ConfigPath)) { $ConfigPath = Join-Path $Root 'config\config.psd1' }
if ([string]::IsNullOrEmpty($CredPath))   { $CredPath   = Join-Path $Root 'config\graph.cred.xml' }

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

. (Join-Path $Root 'lib\graph-auth.ps1')

try { $cfg = Get-GraphConfig -ConfigPath $ConfigPath } catch {
    Write-Host ('✗ ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ''
    Write-Host '请先完成 Azure 应用注册，并把 ClientId 填入 config.psd1：' -ForegroundColor Yellow
    Write-Host '  ClientId = ''<应用程序(客户端) ID>'''
    Write-Host '  Tenant   = ''common''   # 或你的租户 ID / 域名'
    exit 1
}
Write-Host ('ClientId: ' + $cfg['ClientId']) -ForegroundColor DarkGray
Write-Host ('Tenant  : ' + $cfg['Tenant']) -ForegroundColor DarkGray

# ---------- 状态查询 ----------
if ($Status) {
    $tok = Get-GraphTokensSilent -Config $cfg -CredPath $CredPath
    if ($null -eq $tok) {
        Write-Host '状态: 未登录或 refresh token 已失效' -ForegroundColor Yellow
        exit 2
    }
    Write-Host '状态: 已登录，可静默刷新 ✓' -ForegroundColor Green
    try {
        $me = Get-GraphJson -Uri 'https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName' -Token $tok.access_token
        Write-Host ('账号: ' + $me.displayName + '  <' + $(if ($me.mail) { $me.mail } else { $me.userPrincipalName }) + '>')
    } catch {
        Write-Host ('读取账号信息失败: ' + $_.Exception.Message) -ForegroundColor Yellow
    }
    exit 0
}

# ---------- 登录 ----------
$tok = Get-GraphTokensSilent -Config $cfg -CredPath $CredPath
if ($null -ne $tok) {
    Write-Host ''
    Write-Host '检测到已有有效凭据（静默刷新成功），无需重新登录。' -ForegroundColor Green
    Write-Host '如果换账号，请删除 config\graph.cred.xml 后重跑本脚本。' -ForegroundColor DarkGray
} else {
    $tok = Invoke-GraphDeviceCodeLogin -Config $cfg -CredPath $CredPath
}

# ---------- 验证真的能读邮件 ----------
if ($null -ne $tok -and $tok.access_token) {
    Write-Host ''
    Write-Host '验证能否读取邮箱…' -ForegroundColor Cyan
    try {
        $me = Get-GraphJson -Uri 'https://graph.microsoft.com/v1.0/me?$select=displayName,mail' -Token $tok.access_token
        Write-Host ('✓ 账号: ' + $me.displayName + '  <' + $me.mail + '>') -ForegroundColor Green
        $top = Get-GraphJson -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=3&$select=subject,receivedDateTime&$orderby=receivedDateTime desc' -Token $tok.access_token
        $n = @($top.value).Count
        Write-Host ("✓ 收件箱可读，取到 $n 封最近邮件:") -ForegroundColor Green
        foreach ($m in $top.value) {
            Write-Host ('    ' + ([datetime]$m.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') + '  ' + $m.subject)
        }
        Write-Host ''
        Write-Host '凭据已加密保存: ' + $CredPath -ForegroundColor Green
        exit 0
    } catch {
        Write-Host ('✗ 读取邮箱失败: ' + $_.Exception.Message) -ForegroundColor Red
        $d = ''
        try { $d = $_.ErrorDetails.Message } catch {}
        if ($d) { Write-Host ('  ' + $d) -ForegroundColor DarkRed }
        Write-Host '常见原因：缺少 Mail.Read 权限、管理员未同意、或条件访问策略阻止。' -ForegroundColor Yellow
        exit 1
    }
}
Write-Host '未取得 token。' -ForegroundColor Red
exit 1
