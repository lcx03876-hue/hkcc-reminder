<#
    mail-from-thunderbird.ps1 - 从 Thunderbird 的本地缓存读邮件（不需要任何 OAuth / 管理员同意）

    背景：Office 365 的 Graph Mail.Read 权限大多要管理员同意，学生拿不到。
          但 Thunderbird 走 IMAP + OAuth2 是"用户同意"就够了，邮件会缓存到
          %APPDATA%\Thunderbird\Profiles\<profile>\ImapMail\<server>\<folder>。
          本脚本直接解析这些 mbox 文件，只读，不改任何东西。

    前置：Thunderbird 里配好账号；想读的文件夹要先右键账户 -> 订阅；
          想让我看到内容，文件夹还要设为"本地保留"（属性 -> 同步）。

    用法：
      powershell -File .\scripts\mail-from-thunderbird.ps1
      powershell -File .\scripts\mail-from-thunderbird.ps1 -Folder SEHS2250 -Max 20
      powershell -File .\scripts\mail-from-thunderbird.ps1 -Hours 72 -List
      powershell -File .\scripts\mail-from-thunderbird.ps1 -Folder SEHS2250 -Max 5 -Body

    参数：
      -Folder   只看某个文件夹（默认全部）
      -Max      每个文件夹最多显示几封（默认 10）
      -Hours    只看最近 N 小时的邮件（0 = 不限制）
      -List     只列标题，不打印发件人/正文
      -Body     额外打印正文前 400 字
#>
[CmdletBinding()]
param(
    [string]$Folder = '',
    [int]$Max = 10,
    [int]$Hours = 0,
    [switch]$List,
    [switch]$Body,
    [switch]$Sync,            # Thunderbird 没开就先拉起来收一次信
    [switch]$KeepThunderbird  # 用 -Sync 时，读完后不要关掉它
)

$ErrorActionPreference = 'Stop'
# ---------- 按需同步：TB 没开就拉起来，读完再关 ----------
$script:TbStartedByUs = $false
function Start-MailSync {
    $exe = $null
    foreach ($c in @('C:\Program Files\Mozilla Thunderbird\thunderbird.exe',
                     'C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe')) {
        if (Test-Path -LiteralPath $c) { $exe = $c; break }
    }
    if (-not $exe) { Write-Warning '没找到 Thunderbird，跳过同步'; return }
    if (Get-Process thunderbird -ErrorAction SilentlyContinue) { return }
    Write-Host 'Thunderbird 没开，先拉起来收信…' -ForegroundColor DarkGray
    Start-Process -FilePath $exe -ArgumentList '-nosplash'
    $script:TbStartedByUs = $true
    Start-Sleep -Seconds 8
}
function Stop-MailSync {
    if (-not $script:TbStartedByUs -or $KeepThunderbird) { return }
    # 先好声好气地关（让它把索引写回去），关不掉再强杀
    Get-Process thunderbird -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.CloseMainWindow() | Out-Null } catch { }
    }
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 700
        if (-not (Get-Process thunderbird -ErrorAction SilentlyContinue)) {
            Write-Host 'Thunderbird 已关闭。' -ForegroundColor DarkGray
            return
        }
    }
    Get-Process thunderbird -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Host 'Thunderbird 已强制关闭。' -ForegroundColor DarkGray
}
if ($Sync) { Start-MailSync }

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

# ---------- RFC 2047 编码标题 ----------
function ConvertFrom-MimeWord {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    # 相邻的两个编码词之间的空白按 RFC 2047 应当丢弃
    $Text = [regex]::Replace($Text, '\?=\s+=\?', '?==?')
    $eval = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $charset = $m.Groups[1].Value
        $enc = $m.Groups[2].Value.ToUpper()
        $payload = $m.Groups[3].Value
        try {
            if ($enc -eq 'B') {
                $bytes = [Convert]::FromBase64String($payload)
            } else {
                $bytes = [System.Text.Encoding]::ASCII.GetBytes($payload)
            }
            $cs = [System.Text.Encoding]::GetEncoding($charset)
            return $cs.GetString($bytes)
        } catch { return $m.Value }
    }
    return [regex]::Replace($Text, '=\?([^?]+)\?([BbQq])\?([^?]*)\?=', $eval)
}

# ---------- 找 Thunderbird 配置 ----------
function Get-ThunderbirdImapRoots {
    $roots = @()
    $base = Join-Path $env:APPDATA 'Thunderbird\Profiles'
    if (-not (Test-Path $base)) { return $roots }
    foreach ($prof in Get-ChildItem $base -Directory) {
        $imap = Join-Path $prof.FullName 'ImapMail'
        if (Test-Path $imap) { $roots += $imap }
    }
    return $roots
}

# ---------- 解析一个 mbox 的头部 ----------
function Read-MboxHeaders {
    param([string]$Path)

    $out = New-Object System.Collections.Generic.List[object]
    $inMsg = $false
    $inHeaders = $false
    $cur = $null
    $lastKey = $null

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line.StartsWith('From ')) {
            if ($cur) { $out.Add([pscustomobject]$cur) }
            $cur = @{ Subject = ''; From = ''; Date = '' }
            $inMsg = $true
            $inHeaders = $true
            $lastKey = $null
            continue
        }
        if (-not $inMsg) { continue }

        if ($inHeaders) {
            if ($line -eq '') { $inHeaders = $false; continue }
            if ($line -match '^\s' -and $lastKey) {
                # 折行续写
                $cur[$lastKey] = ($cur[$lastKey] + ' ' + $line.Trim())
                continue
            }
            if ($line -match '^([A-Za-z\-]+):\s*(.*)$') {
                $k = $Matches[1]
                $v = $Matches[2]
                switch -Regex ($k) {
                    '^(Subject)$' { $cur.Subject = $v; $lastKey = 'Subject' }
                    '^(From)$'    { $cur.From = $v;    $lastKey = 'From' }
                    '^(Date)$'    { $cur.Date = $v;    $lastKey = 'Date' }
                    default       { $lastKey = $null }
                }
            }
        }
    }
    if ($cur) { $out.Add([pscustomobject]$cur) }
    return $out
}

function ConvertTo-DateTime {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    try { return [datetime]::Parse($Raw) } catch { return $null }
}

# ---------- 主流程 ----------
$roots = Get-ThunderbirdImapRoots
if ($roots.Count -eq 0) {
    Write-Host '✗ 没找到 Thunderbird 配置（%APPDATA%\Thunderbird\Profiles）。先装好并且登录一次。' -ForegroundColor Red
    exit 1
}

$cutoff = if ($Hours -gt 0) { (Get-Date).AddHours(-$Hours) } else { $null }
$total = 0

foreach ($imapRoot in $roots) {
    foreach ($server in Get-ChildItem $imapRoot -Directory) {
        $folders = Get-ChildItem $server.FullName -File | Where-Object { $_.Name -notlike '*.msf' -and $_.Name -ne 'msgFilterRules.dat' }
        foreach ($f in $folders) {
            if ($Folder -and $f.Name -ne $Folder) { continue }
            $msgs = Read-MboxHeaders -Path $f.FullName
            if ($msgs.Count -eq 0) { continue }

            $rows = foreach ($m in $msgs) {
                $dt = ConvertTo-DateTime -Raw $m.Date
                if ($null -eq $dt) { continue }
                if ($cutoff -and $dt -lt $cutoff) { continue }
                [pscustomobject]@{
                    When    = $dt
                    Subject = ConvertFrom-MimeWord $m.Subject
                    From    = ConvertFrom-MimeWord $m.From
                }
            }
            $rows = @($rows | Sort-Object When -Descending)
            if ($rows.Count -eq 0) { continue }

            $total += $rows.Count
            Write-Host ''
            Write-Host ('=== ' + $server.Name + ' / ' + $f.Name + '  (' + $rows.Count + ' 封)') -ForegroundColor Cyan
            $rows | Select-Object -First $Max | ForEach-Object {
                Write-Host ('  ' + $_.When.ToString('MM-dd HH:mm') + '  ' + $_.Subject)
                if (-not $List) {
                    Write-Host ('           ' + $_.From) -ForegroundColor DarkGray
                }
            }
        }
    }
}

if ($total -eq 0) {
    Write-Host '没有匹配的邮件。检查：Thunderbird 是否登录、文件夹是否已订阅且本地保留。' -ForegroundColor Yellow
}

Stop-MailSync
