<#
    ics.ps1 - PowerShell 5.1 兼容的 iCalendar 解析 / 展开 / 摘要工具
    供 HKCC 每日提醒使用。不依赖任何外部模块。
    时区处理：仅支持 +08:00 类（Asia/Hong_Kong）与 UTC(Z)，足够本用途。
#>

function Get-IcsUnfoldedLines {
    param([string[]]$RawLines)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($line in $RawLines) {
        if ($line.Length -gt 0 -and ($line[0] -eq ' ' -or $line[0] -eq "`t")) {
            if ($out.Count -gt 0) { $out[$out.Count - 1] = $out[$out.Count - 1] + $line.Substring(1) }
        } else {
            $out.Add($line)
        }
    }
    return $out
}

function Get-IcsTzidOf {
    <# 从块的 DTSTART 行取出 TZID 参数（无则返回空串，表示浮动/UTC 时间） #>
    param([string[]]$Lines)
    foreach ($l in $Lines) {
        if ($l -like 'DTSTART*' -and $l -match 'TZID=([^:;]+)') { return $Matches[1] }
    }
    return ''
}

function Format-IcsLine {
    <#
        按 RFC 5545 折行：单行不超过 75 字节，续行以单个空格开头。
        按 UTF-8 字节计算，且不会把一个多字节字符劈成两半。
    #>
    param([string]$Line, [int]$Limit = 75)
    $enc = New-Object System.Text.UTF8Encoding($false)
    if ($enc.GetByteCount($Line) -le $Limit) { return @($Line) }

    $chunks = New-Object System.Collections.Generic.List[string]
    $current = ''
    foreach ($ch in $Line.ToCharArray()) {
        $cand = $current + $ch
        $budget = if ($chunks.Count -eq 0) { $Limit } else { $Limit - 1 }  # 续行含前导空格
        if ($enc.GetByteCount($cand) -gt $budget -and $current.Length -gt 0) {
            $chunks.Add($current)
            $current = [string]$ch
        } else {
            $current = $cand
        }
    }
    if ($current.Length -gt 0) { $chunks.Add($current) }

    $out = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $chunks.Count; $i++) {
        if ($i -eq 0) { $out.Add($chunks[$i]) } else { $out.Add(' ' + $chunks[$i]) }
    }
    return $out
}

function Get-IcsProp {
    param([string[]]$Lines, [string]$Name)
    $prefix = $Name + ':'
    $prefixParam = $Name + ';'
    foreach ($l in $Lines) {
        if ($l.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            return $l.Substring($prefix.Length)
        }
        if ($l.StartsWith($prefixParam, [StringComparison]::OrdinalIgnoreCase)) {
            $idx = $l.IndexOf(':')
            if ($idx -ge 0) { return $l.Substring($idx + 1) }
        }
    }
    return $null
}

function Get-IcsPropParams {
    param([string[]]$Lines, [string]$Name)
    $prefix = $Name + ';'
    foreach ($l in $Lines) {
        if ($l.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            $idx = $l.IndexOf(':')
            $head = $l.Substring(0, $idx)
            $pairs = @{}
            foreach ($p in ($head -split ';')) {
                if ($p -match '^([^=]+)=(.*)$') { $pairs[$Matches[1].ToUpper()] = $Matches[2] }
            }
            return $pairs
        }
    }
    return @{}
}

function Get-AllIcsProps {
    param([string[]]$Lines, [string]$Name)
    $res = @()
    $prefix = $Name + ':'
    $prefixParam = $Name + ';'
    foreach ($l in $Lines) {
        if ($l.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $res += $l.Substring($prefix.Length) }
        elseif ($l.StartsWith($prefixParam, [StringComparison]::OrdinalIgnoreCase)) {
            $idx = $l.IndexOf(':'); if ($idx -ge 0) { $res += $l.Substring($idx + 1) }
        }
    }
    return $res
}

function Unescape-IcsText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text -replace '\\n', "`n" -replace '\\N', "`n"
    $t = $t -replace '\\,', ',' -replace '\\;', ';' -replace '\\\\', '\'
    return $t.Trim()
}

function Get-TzOffsetHours {
    param([string]$Tzid)
    if ([string]::IsNullOrEmpty($Tzid)) { return 8 }
    switch -Regex ($Tzid) {
        'Hong_Kong|Chongqing|Urumqi|Taipei|Singapore|Shanghai|China' { return 8 }
        'UTC|GMT' { return 0 }
        'Tokyo|Seoul' { return 9 }
        default { return 8 }
    }
}

function Convert-IcsDate {
    param([string]$Value, [string]$Tzid)
    if ([string]::IsNullOrEmpty($Value)) { return $null }
    $v = $Value.Trim()
    $isUtc = $false
    if ($v.EndsWith('Z')) { $isUtc = $true; $v = $v.Substring(0, $v.Length - 1) }
    $offset = 0
    if ($isUtc) { $offset = 0 } else { $offset = Get-TzOffsetHours -Tzid $Tzid }

    try {
        if ($v.Length -ge 15) {
            $dt = [datetime]::ParseExact($v.Substring(0, 15), 'yyyyMMddTHHmmss', $null)
        } elseif ($v.Length -eq 8) {
            $dt = [datetime]::ParseExact($v, 'yyyyMMdd', $null)
        } else {
            return $null
        }
    } catch { return $null }

    return $dt.AddHours(-$offset)   # 转成 UTC
}

function Convert-FromUtc {
    param([datetime]$Utc, [int]$OffsetHours)
    return $Utc.AddHours($OffsetHours)
}

function Expand-IcsEvent {
    <#
        输入：单个 VEVENT 的解析结果（含 DtStartUtc / DtEndUtc / Rrule / ExDatesUtc）
        输出：发生实例数组 @{ StartUtc; EndUtc }
        支持 FREQ=WEEKLY 的 INTERVAL / COUNT / UNTIL / BYDAY（本用途覆盖率 100%）
        其他 FREQ 仅返回 DTSTART 单次，并打标 Approx=$true
    #>
    param($Event, [int]$OffsetHours = 8)
    $occurrences = @()
    $rrule = $Event.Rrule
    if ([string]::IsNullOrEmpty($rrule)) {
        $occurrences += @{ StartUtc = $Event.DtStartUtc; EndUtc = $Event.DtEndUtc; Approx = $false }
        return $occurrences
    }

    $parts = @{}
    foreach ($p in ($rrule -split ';')) {
        if ($p -match '^([^=]+)=(.*)$') { $parts[$Matches[1].ToUpper()] = $Matches[2] }
    }
    $freq = if ($parts.ContainsKey('FREQ')) { $parts['FREQ'].ToUpper() } else { 'WEEKLY' }
    if ($freq -ne 'WEEKLY') {
        # 本任务数据内不出现；保守处理为单次并标记
        $occurrences += @{ StartUtc = $Event.DtStartUtc; EndUtc = $Event.DtEndUtc; Approx = $true }
        return $occurrences
    }

    $interval = 1
    if ($parts.ContainsKey('INTERVAL')) { [int]::TryParse($parts['INTERVAL'], [ref]$interval) | Out-Null }
    $count = -1
    if ($parts.ContainsKey('COUNT')) { [int]::TryParse($parts['COUNT'], [ref]$count) | Out-Null }
    $untilUtc = $null
    if ($parts.ContainsKey('UNTIL')) {
        $u = $parts['UNTIL']
        $untilUtc = Convert-IcsDate -Value $u -Tzid ''
        if ($null -eq $untilUtc -and $u.Length -ge 15) {
            try { $untilUtc = [datetime]::ParseExact($u.Substring(0, 15), 'yyyyMMddTHHmmss', $null) } catch {}
        }
    }

    # BYDAY：默认取 DTSTART 的星期
    $byDays = @()
    if ($parts.ContainsKey('BYDAY')) {
        foreach ($d in ($parts['BYDAY'] -split ',')) {
            $token = ($d.Trim() -replace '^[+-]?\d+', '').ToUpper()
            if ($token.Length -ge 2) { $byDays += $token.Substring(0, 2) }
        }
    }
    $dayMap = @{ 'SU' = 0; 'MO' = 1; 'TU' = 2; 'WE' = 3; 'TH' = 4; 'FR' = 5; 'SA' = 6 }

    $startLocal = Convert-FromUtc -Utc $Event.DtStartUtc -OffsetHours $OffsetHours
    $duration = $Event.DtEndUtc - $Event.DtStartUtc

    if ($byDays.Count -eq 0) {
        $code = @('SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA')[[int]$startLocal.DayOfWeek]
        $byDays = @($code)
    }

    # 以起始周为基准，逐周推进（覆盖 COUNT / UNTIL）
    $weekAnchor = $startLocal.Date
    $guard = 0
    $produced = 0
    while ($guard -lt 1200) {
        $guard++
        $weekStart = $weekAnchor.AddDays(7 * $interval * ($guard - 1))
        foreach ($bd in $byDays) {
            $target = $weekStart.AddDays((($dayMap[$bd] - [int]$weekAnchor.DayOfWeek) + 7) % 7)
            $occStart = $target.AddHours($startLocal.Hour).AddMinutes($startLocal.Minute).AddSeconds($startLocal.Second)
            if ($occStart -lt $startLocal) { continue }
            $occStartUtc = $occStart.AddHours(-$OffsetHours)
            if ($null -ne $untilUtc -and $occStartUtc -gt $untilUtc) { continue }
            if ($count -ge 0 -and $produced -ge $count) { break }
            # EXDATE 排除。注意：单个元素时 $Event.ExDatesUtc 会被解包成标量，
            # '-contains' 会退化为相等判断而失效，故显式包成数组逐项比较。
            $exList = @($Event.ExDatesUtc)
            if ($exList.Count -gt 0) {
                $isExcluded = $false
                foreach ($ex in $exList) { if ($ex -eq $occStartUtc) { $isExcluded = $true; break } }
                if ($isExcluded) { continue }
            }
            $occurrences += @{ StartUtc = $occStartUtc; EndUtc = $occStartUtc.Add($duration); Approx = $false }
            $produced++
        }
        if ($count -ge 0 -and $produced -ge $count) { break }
        # 无 COUNT/UNTIL 时，推进到 60 周后停止（防失控）
        if ($count -lt 0 -and $null -eq $untilUtc -and $guard -gt 60) { break }
        if ($null -ne $untilUtc -and $weekAnchor.AddDays(7 * $interval * $guard) -gt $untilUtc.AddHours($OffsetHours).AddDays(7)) { break }
    }
    return $occurrences
}

function Read-IcsEvents {
    <# 读取 .ics 文件（或已抓取的行数组），返回规范化事件列表 #>
    param([string[]]$Lines)
    $unfolded = Get-IcsUnfoldedLines -RawLines $Lines
    $events = @()
    $current = $null
    foreach ($line in $unfolded) {
        if ($line -eq 'BEGIN:VEVENT') { $current = New-Object System.Collections.Generic.List[string]; continue }
        if ($line -eq 'END:VEVENT') {
            if ($null -ne $current) {
                $l = $current.ToArray()
                $startRaw = Get-IcsProp -Lines $l -Name 'DTSTART'
                $endRaw = Get-IcsProp -Lines $l -Name 'DTEND'
                $startParams = Get-IcsPropParams -Lines $l -Name 'DTSTART'
                $endParams = Get-IcsPropParams -Lines $l -Name 'DTEND'
                $tz = if ($startParams.ContainsKey('TZID')) { $startParams['TZID'] } else { '' }
                $startUtc = Convert-IcsDate -Value $startRaw -Tzid $tz
                $endUtc = Convert-IcsDate -Value $endRaw -Tzid $tz
                if ($null -ne $startUtc) {
                    if ($null -eq $endUtc -or $endUtc -lt $startUtc) { $endUtc = $startUtc }
                    $exDates = @()
                    foreach ($ex in (Get-AllIcsProps -Lines $l -Name 'EXDATE')) {
                        foreach ($piece in ($ex -split ',')) {
                            $d = Convert-IcsDate -Value $piece -Tzid $tz
                            if ($null -ne $d) { $exDates += $d }
                        }
                    }
                    $events += [pscustomobject]@{
                        Summary    = Unescape-IcsText (Get-IcsProp -Lines $l -Name 'SUMMARY')
                        Location   = Unescape-IcsText (Get-IcsProp -Lines $l -Name 'LOCATION')
                        Description= Unescape-IcsText (Get-IcsProp -Lines $l -Name 'DESCRIPTION')
                        Uid        = Get-IcsProp -Lines $l -Name 'UID'
                        DtStartUtc = $startUtc
                        DtEndUtc   = $endUtc
                        Rrule      = Get-IcsProp -Lines $l -Name 'RRULE'
                        ExDatesUtc = $exDates
                    }
                }
            }
            $current = $null
            continue
        }
        if ($null -ne $current) { $current.Add($line) }
    }
    return $events
}

function Get-IcsTimeline {
    <#
        把 .ics 展开成"发生实例"时间线（不做窗口限制），供 Get-IcsDigest 与
        "后续截止预览"共用同一份解析结果，避免两套过滤逻辑。
        返回每项: Start/End(本地) Summary Location Source AllDayish Approx Expired
    #>
    param(
        [string]$Path,
        [datetime]$NowUtc = ([datetime]::UtcNow),
        [int]$OffsetHours = 8,
        [string]$Label = '',
        [Nullable[datetime]]$SinceLocalDate = $null,
        [Nullable[datetime]]$UntilLocalDate = $null
    )
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $events = Read-IcsEvents -Lines (Get-Content -LiteralPath $Path -Encoding UTF8)
    $sinceDate = $null
    if ($null -ne $SinceLocalDate) { $sinceDate = $SinceLocalDate.Date }
    $nowLocal = Convert-FromUtc -Utc $NowUtc -OffsetHours $OffsetHours

    $rows = @()
    foreach ($ev in $events) {
        foreach ($occ in (Expand-IcsEvent -Event $ev -OffsetHours $OffsetHours)) {
            $s = Convert-FromUtc -Utc $occ.StartUtc -OffsetHours $OffsetHours
            if ($null -ne $sinceDate -and $s.Date -lt $sinceDate) { continue }   # 旧学期
            if ($null -ne $UntilLocalDate -and $s.Date -ge $UntilLocalDate.Date) { continue }
            $e = Convert-FromUtc -Utc $occ.EndUtc -OffsetHours $OffsetHours
            # 课型只写在 DESCRIPTION 里（SUMMARY 只有课程代号+名称），
            # 摘要里补一个 [LEC]/[TUT]/[LAB] 标签，看邮件时一眼能分出是不是实验课。
            $desc = [string]$ev.Description
            $kind = if ($desc -match 'Laboratory') { 'LAB' }
                    elseif ($desc -match 'Tutorial') { 'TUT' }
                    elseif ($desc -match 'Lecture') { 'LEC' }
                    else { '' }
            $rows += [pscustomobject]@{
                Start     = $s
                End       = $e
                Summary   = $ev.Summary
                Location  = $ev.Location
                Source    = $Label
                AllDayish = ($e -le $s)
                Approx    = $occ.Approx
                Expired   = ($e -lt $nowLocal)
                Kind      = $kind
                # RRULE 非空 = 每周重复的常规课。窗口外的预警只看"一次性"条目
                # （截止、补课、手工新增的调课/测试），避免整学期课表刷屏。
                Recurring = (([string]$ev.Rrule).Trim() -ne '')
            }
        }
    }
    return ($rows | Sort-Object Start)
}

function Get-IcsDigest {
    <#
        返回未来 $Days 个自然日（含今天）内的发生实例列表。
        窗口起点是"今天 00:00"而非"此刻"，否则当天已上完的课会被排除。
    #>
    param(
        [string]$Path,
        [int]$Days = 7,
        [datetime]$NowUtc = ([datetime]::UtcNow),
        [int]$OffsetHours = 8,
        [string]$Label = '',
        [Nullable[datetime]]$SinceLocalDate = $null
    )
    $todayLocal = (Convert-FromUtc -Utc $NowUtc -OffsetHours $OffsetHours).Date
    $windowEnd  = $todayLocal.AddDays($Days)          # 含今天，共 Days 天

    # 预过滤区间下限取"旧学期下限"与"窗口起点"中较早者；
    # 但最终仍严格裁到 [今天 00:00, 窗口结束)，否则会把学期初的事件也列进来。
    $preSince = $todayLocal
    if ($null -ne $SinceLocalDate -and $SinceLocalDate.Date -lt $preSince) { $preSince = $SinceLocalDate.Date }

    $tl = Get-IcsTimeline -Path $Path -NowUtc $NowUtc -OffsetHours $OffsetHours -Label $Label `
                          -SinceLocalDate $preSince -UntilLocalDate $windowEnd
    return @($tl | Where-Object { $_.Start -ge $todayLocal -and $_.Start -lt $windowEnd })
}

function Get-UpcomingDeadlinePreview {
    <#
        严格位于窗口之后、且仍在未来的"截止"与"一次性日程"。
        截止 = Start == End；一次性日程 = 非每周重复（RRULE 为空）的课程，
        例如手工在 config\overrides.psd1 里新增的补课、调课、测试条目。
        每周重复的常规课不列入，避免把整学期课表灌进预警区。
        与窗口内内容不重叠，避免邮件里重复列出同一项。
    #>
    param($Timeline, [datetime]$WindowEndLocal, [int]$Max = 3)
    if (-not $Timeline) { return @() }
    return @($Timeline |
        Where-Object { -not $_.Expired -and $_.Start.Date -ge $WindowEndLocal.Date -and ($_.AllDayish -or -not $_.Recurring) } |
        Sort-Object Start | Select-Object -First $Max)
}


function Split-CoursePrefix {
    <#
        把 'CS101: T3 - Repetition ...' 拆成代号 + 标题；
        'CS101 INTRODUCTION TO PROGRAMMING' 也能识别出代号。
        'UNMAPPED-COURSE: xxx' 会被识别为未标注（Unmapped=$true）。
        返回 @{ Code; Title; Unmapped }
    #>
    param([string]$Summary)
    if ([string]::IsNullOrWhiteSpace($Summary)) { return @{ Code = ''; Title = ''; Unmapped = $false } }
    if ($Summary -match '^UNMAPPED-COURSE:\s*(.+)$') { return @{ Code = ''; Title = $Matches[1]; Unmapped = $true } }

    # 先按分隔符切出候选代号，再校验是否像课程码
    # （支持 8C112-ENG 这类数字开头 / 带连字符的代号；纯英文词如 Online 不会被误判）
    $head = $null; $rest = $null
    if ($Summary -match '^([^:\s]{2,14}):\s+(.+)$')    { $head = $Matches[1]; $rest = $Matches[2] }
    elseif ($Summary -match '^([^:\s]{2,14})\s+(.+)$') { $head = $Matches[1]; $rest = $Matches[2] }

    if ($head) {
        $looksLikeCode = (($head -match '^[A-Za-z0-9][A-Za-z0-9\-]*$') -and ($head -match '\d')) -or ($head -match '^[A-Za-z]{2,6}\d{4}$')
        if ($looksLikeCode) { return @{ Code = $head; Title = $rest; Unmapped = $false } }
    }
    return @{ Code = ''; Title = $Summary; Unmapped = $false }
}

function Format-IcsDigestText {
    param($Rows, [string]$Title = 'HKCC reminder', [datetime]$NowLocal, $PreviewDeadlines = @(), $T = $null)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$Title")
    [void]$sb.AppendLine((Format-Mail $T.generated @{ t = $NowLocal.ToString('yyyy-MM-dd HH:mm') }))
    [void]$sb.AppendLine(('=' * 60))
    if (-not $Rows -or $Rows.Count -eq 0) {
        [void]$sb.AppendLine($T.noEvents)
        return $sb.ToString()
    }
    $byDate = $Rows | Group-Object { ($_.Start.ToString('yyyy-MM-dd') + ' (' + (Get-WeekdayName $_.Start $T) + ')') }
    $todayKey = $NowLocal.ToString('yyyy-MM-dd')
    # 窗口含今天，且窗口内还有截止时，明确提示"今天已结束的部分不计入"
    $todayPast = @($Rows | Where-Object { $_.Start.ToString('yyyy-MM-dd') -eq $todayKey -and $_.Expired })
    foreach ($g in $byDate) {
        [void]$sb.AppendLine('')
        $dayMark = ''
        if ($g.Group[0].Start.ToString('yyyy-MM-dd') -eq $todayKey) { $dayMark = '   ' + $T.todayMark }
        [void]$sb.AppendLine('## ' + $g.Name + $dayMark)
        foreach ($r in ($g.Group | Sort-Object Start)) {
            if ($r.AllDayish) {
                $when = $r.Start.ToString('HH:mm') + $T.deadlineSuffix
            } else {
                $when = $r.Start.ToString('HH:mm') + '-' + $r.End.ToString('HH:mm')
            }
            $sp   = Split-CoursePrefix -Summary $r.Summary
            if ($sp.Code) { $code = ('[' + $sp.Code + ']').PadRight(14) }
            elseif ($sp.Unmapped) { $code = ('[' + $T.unmapped + ']').PadRight(14) }
            else { $code = ''.PadRight(14) }
            $line = '  ' + $when.PadRight(12) + $code + $sp.Title
            if ($r.Location) { $line += '  @ ' + $r.Location }
            if ($r.Kind) { $line += '  [' + $r.Kind + ']' }
            if ($r.Expired) { $line += '   ' + $T.expired }
            if ($r.Approx) { $line += '   ' + $T.approx }
            [void]$sb.AppendLine($line)
        }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('=' * 60))
    # 汇总：未标注项数量，方便知道还有多少待标
    $unmapped = @($Rows | Where-Object { $_.Summary -like 'UNMAPPED-COURSE:*' })
    $expired  = @($Rows | Where-Object { $_.Expired })
    $tail = (Format-Mail $T.total @{ n = $Rows.Count })
    if ($unmapped.Count -gt 0) { $tail += (Format-Mail $T.totalUnmapped @{ n = $unmapped.Count }) }
    if ($expired.Count -gt 0)  { $tail += (Format-Mail $T.totalExpired @{ n = $expired.Count }) }
    [void]$sb.AppendLine($tail + "。")
    # 窗口之外的更远截止提醒（避免"7 天窗口"漏掉下周一要交的作业）
    if ($PreviewDeadlines -and @($PreviewDeadlines).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine((Format-Mail $T.previewHead @{ days = 7 }) + ':')
        foreach ($r in ($PreviewDeadlines | Sort-Object Start)) {
            $sp = Split-CoursePrefix -Summary $r.Summary
            if ($sp.Code) { $code = ('[' + $sp.Code + ']').PadRight(14) } else { $code = ''.PadRight(14) }
            [void]$sb.AppendLine('  ' + ($r.Start.ToString('MM-dd') + ' (' + (Get-WeekdayName $r.Start $T) + ' ' + $r.Start.ToString('HH:mm') + ')').PadRight(18) + $code + $sp.Title + $(if ($r.Kind) { '  [' + $r.Kind + ']' } else { '' }))
        }
    }
    return $sb.ToString()
}
