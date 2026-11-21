<#
    daily-reminder.ps1 - 每天运行：同步日历源 -> 生成未来 7 天摘要 -> 发送邮件
    退出码：0 成功（含无事件）；1 出错；2 邮件发送失败
#>
[CmdletBinding()]
param(
    [int]$Days = 7,
    [string]$ConfigPath = '',
    [string]$CredPath   = '',
    [switch]$NoMail,
    [switch]$Preview
)

# 注意：$PSScriptRoot 在 param 默认值求值时尚未赋值（用 -File 调用时为空），
# 因此路径必须在这里推导，不能写在 param 默认值里。
$SelfDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ProjectRoot = Split-Path -Parent $SelfDir
if ([string]::IsNullOrEmpty($ConfigPath)) { $ConfigPath = Join-Path $ProjectRoot 'config\config.psd1' }
if ([string]::IsNullOrEmpty($CredPath))   { $CredPath   = Join-Path $ProjectRoot 'config\smtp.cred.xml' }

$ErrorActionPreference = 'Stop'
$HkOffset = 8   # Asia/Hong_Kong = UTC+8（香港 1979 年后无夏令时）

# 关键：Windows PowerShell 5.1 默认按 OEM 代码页输出，中文会乱码（也会污染邮件正文字节）。
# 显式切到 UTF-8，保证控制台 / 日志 / 邮件 / 文件四处一致。
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch {}
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ('System.Web.HttpUtility' -as [type]) { return [System.Web.HttpUtility]::HtmlEncode($Text) }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$stamp] $Message"
    Write-Output $line
    $logDir = Join-Path $ProjectRoot 'logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir ('reminder-' + (Get-Date -Format 'yyyy-MM') + '.log')
    Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
}

try {
    $libPath = Join-Path $ProjectRoot 'lib\ics.ps1'
    if (-not (Test-Path -LiteralPath $libPath)) { throw "找不到解析库: $libPath" }
    . $libPath
    $i18nPath = Join-Path $ProjectRoot 'lib\i18n.ps1'
    if (Test-Path -LiteralPath $i18nPath) { . $i18nPath }

    # ---------- 配置 ----------
    $config = @{
        SmtpServer = 'smtp.gmail.com'
        SmtpPort   = 587
        UseSsl     = $true
        From       = ''
        To         = @()
        SubjectPrefix = '[HKCC]'
        Days       = $Days
        HideBeforeDate = ''
    }
    if (Test-Path -LiteralPath $ConfigPath) {
        $userCfg = Import-PowerShellDataFile -LiteralPath $ConfigPath
        foreach ($k in $userCfg.Keys) { $config[$k] = $userCfg[$k] }
    }
    if ($config.ContainsKey('Days') -and $config['Days']) { $Days = [int]$config['Days'] }
    # 旧学期过滤：早于该日期的事件不进入摘要（与 build-merged.ps1 保持一致）
    $sinceLocalDate = $null
    if ($config.ContainsKey('HideBeforeDate') -and -not [string]::IsNullOrWhiteSpace([string]$config['HideBeforeDate'])) {
        try { $sinceLocalDate = [datetime]::ParseExact([string]$config['HideBeforeDate'], 'yyyy-MM-dd', $null) } catch {
            Write-Log ('  [warn] HideBeforeDate 格式应为 yyyy-MM-dd，已忽略: ' + $config['HideBeforeDate'])
        }
    }
    # Panel overrides (panel\settings.json): the panel writes this when it cannot
# write config.psd1 directly. Values here win over config.psd1.
    $panelSettings = Join-Path (Split-Path -Parent $ProjectRoot) 'panel\settings.json'
    if (Test-Path -LiteralPath $panelSettings) {
        try {
            $ov = Get-Content -LiteralPath $panelSettings -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($ov.to) { $config['To'] = @($ov.to) }
            if ($ov.days) { $config['Days'] = [int]$ov.days; $Days = [int]$ov.days }
            if ($ov.subject) { $config['SubjectPrefix'] = [string]$ov.subject }
            if ($ov.language) { $config['Language'] = [string]$ov.language }
            Write-Log ('  闈㈡澘璁剧疆宸插悎骞? panel\settings.json')
        } catch {
            Write-Log ('  [warn] 闈㈡澘璁剧疆璇诲彇澶辫触: ' + $_.Exception.Message)
        }
    }
$T = $null
    if (Get-Command Get-MailStrings -ErrorAction SilentlyContinue) {
        $T = Get-MailStrings ([string]$config['Language'])
    }
    if (-not $T) { Write-Log '  [warn] lib\i18n.ps1 missing or failed, falling back to zh-CN wording' }
    if ([string]::IsNullOrWhiteSpace($config['From'])) { throw "config.psd1 里 From（发件邮箱）不能为空" }
    if (-not $config['To'] -or $config['To'].Count -eq 0) { throw "config.psd1 里 To（收件邮箱）不能为空" }

    # ---------- 1. 同步数据源 ----------
    Write-Log '开始同步 Blackboard 日历源…'
    # sync-feed.ps1 在 lib\（不是 scripts\）。历史上这里写的是 $SelfDir，
    # 导致 Test-Path 落空 → 整段被静默跳过 → 日历源一个多月没更新（2026-11-05 修复）。
    # 两个位置都试，避免以后再挪目录时又静默失效。
    $syncScript = $null
    foreach ($cand in @((Join-Path $ProjectRoot 'lib\sync-feed.ps1'),
                        (Join-Path $SelfDir    'sync-feed.ps1'))) {
        if (Test-Path -LiteralPath $cand) { $syncScript = $cand; break }
    }
    if ($syncScript) {
        $syncOut = & $syncScript 2>&1
        foreach ($l in $syncOut) { Write-Log ('  ' + $l) }
    } else {
        Write-Log '  [ERROR] 找不到 sync-feed.ps1（lib\ 与 scripts\ 都没有），本次未同步源'
    }

    # ---------- 1a. 监测日历源变动（调课/补课/停课/换教室）----------
    # 不需要读邮件：Blackboard 日历源只要发生变化就能发现。
    $feedChanges = @()
    $monitorScript = Join-Path $ProjectRoot 'lib\feed-monitor.ps1'
    if (Test-Path -LiteralPath $monitorScript) {
        . $monitorScript
        $bbIcs   = Join-Path $ProjectRoot 'data\blackboard.ics'
        $statePath = Join-Path $ProjectRoot 'data\feed-state.json'
        try {
            $feedChanges = @(Update-FeedState -IcsPath $bbIcs -StatePath $statePath)
            Write-Log ('  日历源变动监测: ' + @($feedChanges).Count + ' 项')
            foreach ($ch in $feedChanges) {
                Write-Log ('    [' + $ch.Kind + '] ' + $ch.Summary)
            }
            # 只保留今天及以后的变动。往年（2025/26 上学期残留）和已经过去的调课/移除
            # 是历史噪声，用户 2026-11-10 明确要求以后不要再提示。
            if ($feedChanges.Count -gt 0 -and (Get-Command Select-ActualFeedChanges -ErrorAction SilentlyContinue)) {
                $sel = Select-ActualFeedChanges -Changes $feedChanges -Today ([datetime]::UtcNow.AddHours($HkOffset))
                if (@($sel.Skip).Count -gt 0) {
                    Write-Log ('  忽略已过期/往年变动: ' + @($sel.Skip).Count + ' 项')
                    foreach ($ch in @($sel.Skip)) { Write-Log ('    [skip] ' + $ch.Summary) }
                }
                $feedChanges = @($sel.Keep)
            }
        } catch {
            Write-Log ('  [warn] 变动监测失败: ' + $_.Exception.Message)
        }
    }

    # ---------- 1b. 重新合并导出可导入 Outlook 的 .ics ----------
    $mergeScript = Join-Path $ProjectRoot 'lib\build-merged.ps1'
    if (Test-Path -LiteralPath $mergeScript) {
        $mergeOut = & $mergeScript 2>&1
        foreach ($l in $mergeOut) { Write-Log ($l) }
    }

    # ---------- 2. 汇总未来 N 天 ----------
    $nowUtc   = [datetime]::UtcNow
    $nowLocal = $nowUtc.AddHours($HkOffset)
    # 数据源：读"过滤 + 标注"后的产物（由 build-merged.ps1 生成），
    # 保证邮件 / 日历 / CSV 报表三者完全一致，不会出现两套过滤或标注逻辑。
    $outDirLocal = Join-Path $ProjectRoot 'out'
    $sources = @(
        @{ Path = (Join-Path $outDirLocal 'HKCC_deadlines.ics'); Label = 'Blackboard' },
        @{ Path = (Join-Path $outDirLocal 'HKCC_classes.ics');   Label = 'Timetable'  }
    )
    $rows = @()
    $fullTimeline = @()   # 不做窗口限制，供"后续截止"预览使用（同一份解析结果）
    foreach ($src in $sources) {
        if (Test-Path -LiteralPath $src.Path) {
            $part = Get-IcsDigest -Path $src.Path -Days $Days -NowUtc $nowUtc -OffsetHours $HkOffset -Label $src.Label -SinceLocalDate $sinceLocalDate
            if ($part) { $rows += $part }
            $tl = Get-IcsTimeline -Path $src.Path -NowUtc $nowUtc -OffsetHours $HkOffset -Label $src.Label -SinceLocalDate $sinceLocalDate
            if ($tl) { $fullTimeline += $tl }
            Write-Log ("  读取 " + $src.Label + ": " + (@($part).Count) + " 项（窗口内）")
        } else {
            Write-Log ("  跳过 " + $src.Label + "（文件不存在: " + $src.Path + "）")
        }
    }
    $rows = @($rows | Sort-Object Start)
    # 7 天窗口之外、仍在未来的最近 8 条"截止 + 一次性日程"（不与窗口内已列出的项重复）
    $windowEndLocal = $nowLocal.Date.AddDays($Days)
    $deadlinePreview = Get-UpcomingDeadlinePreview -Timeline $fullTimeline -WindowEndLocal $windowEndLocal -Max 8
    if (@($deadlinePreview).Count -gt 0) {
        Write-Log ('  后续日程预览: ' + @($deadlinePreview).Count + ' 项')
    }

    # ---------- 3. 生成正文 ----------
    # 标题：窗口是"今天起 7 个自然日（含今天）"，且末尾另有窗口外的"后续截止"，
    # 所以不写"未来 7 天"，而是直接印出真实日期区间。
    # 注意：中文措辞放在本文件（带 BOM），config.psd1 必须保持纯 ASCII。
    $prefix    = [string]$config['SubjectPrefix']
    $winStart  = $nowLocal.Date
    $winEnd    = $nowLocal.Date.AddDays($Days - 1)
    $rangeText = $winStart.ToString('MM-dd') + ' ~ ' + $winEnd.ToString('MM-dd')
    # 仅当区间跨年时给结束日补年份，避免 "12-28 ~ 01-03" 的歧义
    if ($winStart.Year -ne $winEnd.Year) { $rangeText = $winStart.ToString('MM-dd') + ' ~ ' + $winEnd.ToString('yyyy-MM-dd') }
    $subject   = $prefix + ' ' + (Format-Mail $T.subj @{ range = $rangeText; n = $rows.Count })
    $title     = $prefix + ' ' + (Format-Mail $T.title @{ range = $rangeText; date = $nowLocal.ToString('yyyy-MM-dd'); days = $Days })
    $textBody = Format-IcsDigestText -Rows $rows -Title $title -NowLocal $nowLocal -PreviewDeadlines $deadlinePreview -T $T

    # ---------- 3b. 变动提醒（放在正文最前，最醒目）----------
    if (@($feedChanges).Count -gt 0) {
        $kindText = @{ 'Added' = $T.chgAdded; 'Changed' = $T.chgChanged; 'Removed' = $T.chgRemoved }
        $sbChg = New-Object System.Text.StringBuilder
        [void]$sbChg.AppendLine('')
        [void]$sbChg.AppendLine('============================================================')
        [void]$sbChg.AppendLine($T.chgHead)
        [void]$sbChg.AppendLine('============================================================')
        foreach ($ch in @($feedChanges | Sort-Object Kind, Summary)) {
            $label = if ($kindText.ContainsKey($ch.Kind)) { $kindText[$ch.Kind] } else { $ch.Kind }
            [void]$sbChg.AppendLine('')
            [void]$sbChg.AppendLine('  ● ' + $ch.Summary)
            [void]$sbChg.AppendLine('    [' + $label + ']')
            if ($ch.Kind -eq 'Changed') {
                [void]$sbChg.AppendLine('    ' + $T.chgFields + ($ch.Fields -join '、'))
                [void]$sbChg.AppendLine('    ' + $T.chgOld + (Get-IcsRawValueText $ch.Old.Start) + $(if ($ch.Old.Location) { '  @ ' + $ch.Old.Location } else { '' }))
                [void]$sbChg.AppendLine('    ' + $T.chgNew + (Get-IcsRawValueText $ch.New.Start) + $(if ($ch.New.Location) { '  @ ' + $ch.New.Location } else { '' }))
            } elseif ($ch.Kind -eq 'Added') {
                [void]$sbChg.AppendLine('    ' + $T.chgTime + (Get-IcsRawValueText $ch.New.Start) + $(if ($ch.New.Location) { '  @ ' + $ch.New.Location } else { '' }))
            } elseif ($ch.Kind -eq 'Removed') {
                [void]$sbChg.AppendLine('    ' + $T.chgWas + (Get-IcsRawValueText $ch.Old.Start) + $(if ($ch.Old.Location) { '  @ ' + $ch.Old.Location } else { '' }))
            }
        }
        [void]$sbChg.AppendLine('')
        [void]$sbChg.AppendLine('  ' + $T.chgNote)
        [void]$sbChg.AppendLine('')
        $textBody = $sbChg.ToString() + $textBody
    }

    # ---- 新版排版（v14）----
    # DUE SOON（P0<=48h 红 / P1<=7d 琥珀）、TODAY & TOMORROW、FURTHER OUT（7-14d）
    . (Join-Path $ProjectRoot 'lib\render-v14.ps1')
    $classRows    = @($rows | Where-Object { $_.Source -eq 'Timetable' })
    $deadlineRows = @(@($rows | Where-Object { $_.Source -ne 'Timetable' }) + @($deadlinePreview))
    $htmlLang     = [string]$config['Language']
    if ([string]::IsNullOrEmpty($htmlLang) -or $htmlLang -eq 'auto') { $htmlLang = 'en' }
    $htmlBody = Render-V14Html -Classes $classRows -Deadlines $deadlineRows -Changes @($feedChanges) -Now $nowLocal -Lang $htmlLang

    # 落地一份纯文本存档
    $outDir = Join-Path $ProjectRoot 'out'
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $txtFile = Join-Path $outDir ('digest-' + (Get-Date -Format 'yyyy-MM-dd') + '.txt')
    Set-Content -LiteralPath $txtFile -Value $textBody -Encoding UTF8
    $htmlFile = Join-Path $outDir ('digest-' + (Get-Date -Format 'yyyy-MM-dd') + '.html')
    Set-Content -LiteralPath $htmlFile -Value $htmlBody -Encoding UTF8
    Write-Log ("  摘要已保存: " + $txtFile + "  |  邮件预览: " + $htmlFile)

    if ($Preview) {
        Write-Output '----- PREVIEW -----'
        Write-Output $textBody
        exit 0
    }

    if ($NoMail) { Write-Log '已指定 -NoMail，跳过发送。'; exit 0 }

    # ---------- 4. 发送邮件 ----------
    if (-not (Test-Path -LiteralPath $CredPath)) {
        Write-Log ("未找到凭据文件: " + $CredPath)
        Write-Log "请先在本机运行: pwsh -File .\scripts\init-credential.ps1"
        exit 2
    }
    $cred = Import-Clixml -LiteralPath $CredPath
    $mailParams = @{
        SmtpServer  = $config['SmtpServer']
        Port        = [int]$config['SmtpPort']
        Credential  = $cred
        From        = $config['From']
        To          = $config['To']
        Subject     = $subject
        Body        = $htmlBody
        BodyAsHtml  = $true
        Encoding    = [System.Text.Encoding]::UTF8
    }
    # -UseSsl 是 switch 型参数：仅在需要时加入，避免传位置值报错
    if ([bool]$config['UseSsl']) { $mailParams['UseSsl'] = $true }
    # 完整截止清单作为附件（很小，~10 KB），方便一键导入日历
    $attachIcs = Join-Path $ProjectRoot 'out\HKCC_deadlines.ics'
    if (Test-Path -LiteralPath $attachIcs) { $mailParams['Attachments'] = @($attachIcs) }

    # 发信重试：Office365 偶尔直接掐断 SMTP 连接（2026-11-13 就整封漏发过），
    # 隔一会儿再试通常就通了。
    $maxTry = 3
    $delays = @(60, 180)
    $sent = $false
    for ($i = 1; $i -le $maxTry; $i++) {
        try {
            Send-MailMessage @mailParams
            $sent = $true
            if ($i -gt 1) {
                Write-Log ("邮件已发送（第 " + $i + " 次尝试成功） -> " + ($config['To'] -join ', ') + "  共 " + $rows.Count + " 项")
            } else {
                Write-Log ("邮件已发送 -> " + ($config['To'] -join ', ') + "  共 " + $rows.Count + " 项")
            }
            break
        } catch {
            $msg = $_.Exception.Message
            if ($i -lt $maxTry) {
                Write-Log ("发信失败（第 " + $i + " 次）：" + $msg + " ；" + $delays[$i - 1] + " 秒后重试")
                Start-Sleep -Seconds $delays[$i - 1]
            } else {
                throw
            }
        }
    }
    if (-not $sent) { throw "发信未完成" }
    exit 0
} catch {
    Write-Log ('ERROR: ' + $_.Exception.Message)
    if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) {
        Write-Log ('  位置: 第 ' + $_.InvocationInfo.ScriptLineNumber + ' 行')
    }
    exit 1
}