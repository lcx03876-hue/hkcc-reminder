<#
    feed-monitor.ps1 - 监测 Blackboard 日历源的变动（不需要读邮件）

    原理：
      每次抓取源后，为每个事件记录一份"指纹"（UID + DTSTART + DTEND + SUMMARY）。
      与上一次快照对比，报告：
        新增 (Added)     —— 例如新发布的补课/调课
        修改 (Changed)   —— 例如时间/地点被改（最可能就是调课）
        消失 (Removed)   —— 例如被删除（可能是停课）
      快照存于 data\feed-state.json。

    局限（必须知情）：
      只能发现"已同步到 Blackboard 日历"的变动。
      只发邮件、没进日历的调课抓不到。
#>

function Get-IcsFingerprints {
    <#
        从 .ics 提取每个事件的指纹：UID -> @{ Start; End; Summary; Location }
        Start/End 用原始字符串，便于精确比较（避免时区解析差异）。
    #>
    param([string[]]$Lines)
    $unfolded = Get-IcsUnfoldedLines -RawLines $Lines
    $map = @{}
    $cur = $null
    foreach ($line in $unfolded) {
        if ($line -eq 'BEGIN:VEVENT') { $cur = New-Object System.Collections.Generic.List[string]; continue }
        if ($line -eq 'END:VEVENT') {
            if ($null -ne $cur) {
                $l = $cur.ToArray()
                $uid = Get-IcsProp -Lines $l -Name 'UID'
                if ($uid) {
                    $map[$uid] = @{
                        Start    = [string](Get-IcsProp -Lines $l -Name 'DTSTART')
                        End      = [string](Get-IcsProp -Lines $l -Name 'DTEND')
                        Summary  = [string](Get-IcsProp -Lines $l -Name 'SUMMARY')
                        Location = [string](Get-IcsProp -Lines $l -Name 'LOCATION')
                    }
                }
            }
            $cur = $null
            continue
        }
        if ($null -ne $cur) { $cur.Add($line) }
    }
    return $map
}

function Get-IcsDurationMinutes {
    <#
        由原始 DTSTART/DTEND 字符串算时长（分钟）。
        无法解析时返回 -1，让调用方按"不同"处理（宁可多报）。
    #>
    param([string]$Start, [string]$End)
    if ([string]::IsNullOrEmpty($Start) -or [string]::IsNullOrEmpty($End)) { return -1 }
    $s = $null; $e = $null
    if ($Start -match '(\d{8})T(\d{6})') { try { $s = [datetime]::ParseExact($Matches[1] + $Matches[2], 'yyyyMMddHHmmss', $null) } catch {} }
    if ($End   -match '(\d{8})T(\d{6})') { try { $e = [datetime]::ParseExact($Matches[1] + $Matches[2], 'yyyyMMddHHmmss', $null) } catch {} }
    if ($null -eq $s -or $null -eq $e) { return -1 }
    return [int]($e - $s).TotalMinutes
}

function Compare-IcsFingerprints {
    <#
        对比新旧指纹，返回变动列表。
        返回 @( @{ Kind='Added'|'Changed'|'Removed'; Uid; Summary; Old; New } )
    #>
    param($Old, $New)
    if ($null -eq $Old) { $Old = @{} }
    $changes = @()

    foreach ($uid in $New.Keys) {
        $n = $New[$uid]
        if (-not $Old.ContainsKey($uid)) {
            $changes += [pscustomobject]@{ Kind = 'Added'; Uid = $uid; Summary = $n.Summary; Old = $null; New = $n }
            continue
        }
        $o = $Old[$uid]
        $diff = @()
        if ($o.Start -ne $n.Start)     { $diff += '时间' }
        # 用"真实时长"判断结束时间是否额外变化：
        #   整体平移（09:00-10:00 → 11:00-12:00）时长不变，不重复报
        #   时长真变了（09:00-10:00 → 09:00-11:30）才报
        if ((Get-IcsDurationMinutes -Start $o.Start -End $o.End) -ne (Get-IcsDurationMinutes -Start $n.Start -End $n.End)) {
            $diff += '时长/结束时间'
        }
        if ($o.Summary -ne $n.Summary) { $diff += '标题' }
        if ($o.Location -ne $n.Location) { $diff += '地点' }
        if ($diff.Count -gt 0) {
            $changes += [pscustomobject]@{ Kind = 'Changed'; Uid = $uid; Summary = $n.Summary; Old = $o; New = $n; Fields = $diff }
        }
    }
    foreach ($uid in $Old.Keys) {
        if (-not $New.ContainsKey($uid)) {
            $o = $Old[$uid]
            $changes += [pscustomobject]@{ Kind = 'Removed'; Uid = $uid; Summary = $o.Summary; Old = $o; New = $null }
        }
    }
    return $changes
}

function Get-IcsRawValueText {
    <# 把 'TZID=Asia/Hong_Kong:20250101T083000' 变成可读 '2025-01-01 08:30' #>
    param([string]$Raw)
    if ([string]::IsNullOrEmpty($Raw)) { return '' }
    if ($Raw -match '(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})') {
        return ($Matches[1] + '-' + $Matches[2] + '-' + $Matches[3] + ' ' + $Matches[4] + ':' + $Matches[5])
    }
    if ($Raw -match '(\d{4})(\d{2})(\d{2})$') {
        return ($Matches[1] + '-' + $Matches[2] + '-' + $Matches[3])
    }
    return $Raw
}

function Update-FeedState {
    <#
        读取快照 -> 对比 -> 写回新快照。
        返回变动列表（首次运行时返回空，只建立基线）。
    #>
    param(
        [string]$IcsPath,
        [string]$StatePath,
        [switch]$Reset
    )
    if (-not (Test-Path -LiteralPath $IcsPath)) { return @() }
    $lines = Get-Content -LiteralPath $IcsPath -Encoding UTF8
    $current = Get-IcsFingerprints -Lines $lines

    $previous = $null
    if ((Test-Path -LiteralPath $StatePath) -and -not $Reset) {
        try {
            $json = Get-Content -LiteralPath $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            $previous = @{}
            foreach ($p in $json.PSObject.Properties) {
                $previous[$p.Name] = @{
                    Start    = [string]$p.Value.Start
                    End      = [string]$p.Value.End
                    Summary  = [string]$p.Value.Summary
                    Location = [string]$p.Value.Location
                }
            }
        } catch { $previous = $null }
    }

    # 落盘新快照
    $dir = Split-Path -Parent $StatePath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $obj = @{}
    foreach ($k in $current.Keys) { $obj[$k] = $current[$k] }
    ($obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $StatePath -Encoding UTF8

    if ($null -eq $previous) { return @() }   # 首次运行：只建立基线
    return (Compare-IcsFingerprints -Old $previous -New $current)
}
