<#
    graph-auth.ps1 - 零依赖 Microsoft Graph OAuth2 认证（设备码流程 + 静默刷新）

    为什么手写而不用 Microsoft.Graph 模块：
      - 本机只有 Windows PowerShell 5.1，Graph SDK 又大又慢
      - 只用到 Invoke-RestMethod + DPAPI，逻辑完全透明可控

    凭据存储：config\graph.cred.xml，用 DPAPI 加密（只有本机本账户能解密）
    注意：refresh token 本身也是敏感凭据，等同长期密码，切勿外传。
#>

$script:GraphTokens = @{}   # 进程内缓存，避免同一次运行重复换 token

function Get-GraphConfig {
    param([string]$ConfigPath)
    if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "找不到配置: $ConfigPath" }
    $cfg = Import-PowerShellDataFile -LiteralPath $ConfigPath
    foreach ($k in @('ClientId', 'Tenant')) {
        if (-not $cfg.ContainsKey($k) -or [string]::IsNullOrWhiteSpace([string]$cfg[$k])) {
            throw "config.psd1 缺少 $k（Graph 读邮件必需）"
        }
    }
    return $cfg
}

function Get-GraphScopeList {
    param($Config)
    $scopes = @('offline_access')
    if ($Config.ContainsKey('GraphScopes') -and $Config['GraphScopes']) { $scopes += @($Config['GraphScopes']) }
    else { $scopes += 'https://graph.microsoft.com/Mail.Read' }    return ($scopes -join ' ')
}

function Get-GraphAuthority {
    <#
        OAuth2 授权端点。单租户应用（"仅此组织目录中的帐户"）用 'common'
        有时会报租户不匹配，'organizations' 更稳；两者都能正常刷新 token。
    #>
    param($Config)
    $tenant = [string]$Config['Tenant']
    if ([string]::IsNullOrWhiteSpace($tenant)) { $tenant = 'organizations' }
    switch -Regex ($tenant.ToLower()) {
        '^(common|organizations|consumers)$' { return $tenant.ToLower() }
        default { return $tenant }   # 具体租户 ID 或域名
    }
}

function Invoke-GraphDeviceCodeLogin {
    <#
        交互式设备码登录（首次运行用）。
        会打印一个代码，让你在浏览器里输入；登录成功后 refresh token 被缓存。
    #>
    param($Config, [string]$CredPath, [switch]$NoSave)

    $tenant = Get-GraphAuthority -Config $Config
    $clientId = [string]$Config['ClientId']
    $scope = Get-GraphScopeList -Config $Config
    $dcUrl = "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode"

    Write-Host ''
    Write-Host '=== 正在申请设备码 ===' -ForegroundColor Cyan
    try {
        $dc = Invoke-RestMethod -Method Post -Uri $dcUrl -Body @{ client_id = $clientId; scope = $scope } -ErrorAction Stop
    } catch {
        $detail = ''
        try { $detail = $_.ErrorDetails.Message } catch {}
        if (-not $detail) { $detail = $_.Exception.Message }
        throw ('申请设备码失败（端点 ' + $dcUrl + '）: ' + $detail)
    }
    if (-not $dc -or -not $dc.device_code) { throw '设备码响应异常：未返回 device_code' }

    Write-Host ''
    Write-Host '--------------------------------------------------------' -ForegroundColor Yellow
    Write-Host ('  请在浏览器打开: ' + $dc.verification_uri) -ForegroundColor White
    Write-Host ('  并输入代码:     ' + $dc.user_code) -ForegroundColor Green
    Write-Host ($dc.message) -ForegroundColor DarkGray
    Write-Host '--------------------------------------------------------' -ForegroundColor Yellow
    Write-Host ''

    $interval = 5
    if ($dc.PSObject.Properties['interval']) { $interval = [int]$dc.interval }
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $clientId
                device_code = $dc.device_code
            } -ErrorAction Stop
        } catch {
            # 400 = authorization_pending（正常，继续轮询）；其他错误直接抛出
            $body = ''
            try { $body = $_.ErrorDetails.Message } catch {}
            if ($body -match 'authorization_pending') { continue }
            if ($body -match 'slow_down') { $interval += 5; continue }
            if ($body -match 'expired_token') { throw '设备码已过期，请重新运行。' }
            if ($body -match 'authorization_declined') { throw '你在浏览器里拒绝了授权。' }
            throw ("设备码轮询失败: " + $body)
        }
        Write-Host '✓ 登录成功' -ForegroundColor Green
        if (-not $NoSave) { Save-GraphTokens -Tokens $tok -Config $Config -CredPath $CredPath }
        return $tok
    }
    throw '设备码超时（未在有效期内完成登录）。'
}

function Save-GraphTokens {
    param($Tokens, $Config, [string]$CredPath)
    $obj = [pscustomobject]@{
        ClientId     = [string]$Config['ClientId']
        Tenant       = [string]$Config['Tenant']
        RefreshToken = $Tokens.refresh_token
        Scope        = $Tokens.scope
        SavedAt      = (Get-Date).ToString('o')
    }
    $dir = Split-Path -Parent $CredPath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $obj | Export-Clixml -LiteralPath $CredPath
}

function Get-GraphTokensSilent {
    <#
        用缓存的 refresh token 静默换取 access token。
        返回 $null 表示需要重新交互登录。
    #>
    param($Config, [string]$CredPath)

    if (-not (Test-Path -LiteralPath $CredPath)) { return $null }
    try { $saved = Import-Clixml -LiteralPath $CredPath } catch { return $null }
    if (-not $saved -or -not $saved.RefreshToken) { return $null }
    # client_id 变了则缓存的 refresh token 无效
    if ($saved.ClientId -ne [string]$Config['ClientId']) { return $null }

    $tenant = Get-GraphAuthority -Config $Config
    $scope  = Get-GraphScopeList -Config $Config
    try {
        $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body @{
            grant_type    = 'refresh_token'
            client_id     = [string]$Config['ClientId']
            refresh_token = $saved.RefreshToken
            scope         = $scope
        } -ErrorAction Stop
    } catch {
        $body = ''
        try { $body = $_.ErrorDetails.Message } catch {}
        if ($body -match 'invalid_grant') { return $null }   # refresh token 失效，需重新登录
        throw ("刷新 token 失败: " + $body)
    }
    # 轮换后的 refresh token 要落盘，否则下次可能失效
    if ($tok.refresh_token) { Save-GraphTokens -Tokens $tok -Config $Config -CredPath $CredPath }
    return $tok
}

function Get-GraphAccessToken {
    <#
        统一入口：进程内缓存 -> 静默刷新 -> （可选）交互式设备码登录
    #>
    param(
        [string]$ConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'config\config.psd1'),
        [string]$CredPath   = (Join-Path (Split-Path -Parent $PSScriptRoot) 'config\graph.cred.xml'),
        [switch]$Interactive
    )
    $cfg = Get-GraphConfig -ConfigPath $ConfigPath
    $key = [string]$cfg['ClientId']
    if ($script:GraphTokens.ContainsKey($key) -and $script:GraphTokens[$key].expires_at -gt (Get-Date)) {
        return $script:GraphTokens[$key].access_token
    }
    $tok = Get-GraphTokensSilent -Config $cfg -CredPath $CredPath
    if ($null -eq $tok) {
        if (-not $Interactive) { return $null }   # 调用方决定是否交互
        $tok = Invoke-GraphDeviceCodeLogin -Config $cfg -CredPath $CredPath
    }
    if ($null -eq $tok -or -not $tok.access_token) { return $null }
    $expiresAt = (Get-Date).AddSeconds([int]$tok.expires_in - 60)
    $script:GraphTokens[$key] = @{ access_token = $tok.access_token; expires_at = $expiresAt }
    return $tok.access_token
}

function Get-GraphJson {
    <# 调 Graph REST，自动带 token；分页由调用方用 NextLink 处理 #>
    param([string]$Uri, [string]$Token = '')
    if ([string]::IsNullOrEmpty($Token)) { $Token = Get-GraphAccessToken }
    if ([string]::IsNullOrEmpty($Token)) { throw '没有得到 access token（需先运行 scripts\mail-login.ps1）' }
    $headers = @{ Authorization = "Bearer $Token"; Accept = 'application/json' }
    return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -ErrorAction Stop
}
