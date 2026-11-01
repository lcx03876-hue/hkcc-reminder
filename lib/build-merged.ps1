# NOTE: keep the UTF-8 BOM on this file. Windows PowerShell 5.1 reads .ps1 using the
# ANSI code page when there is no BOM, which turns the Chinese comments below into
# mojibake and breaks parsing (the scheduled task then fails).
<#
    build-merged.ps1 - 把 Blackboard 日历源与 HKCC 课表合并成一个可导入 Outlook 的 .ics
    每天由 daily-reminder.ps1 自动调用，也可手动运行。
    输出: out\HKCC_merged.ics
         out\HKCC_deadlines.ics   (仅作业/考试截止)
         out\HKCC_classes.ics     (仅上课)
#>
[CmdletBinding()]
param(
    [string]$Root = ''
)

# $PSScriptRoot 在 param 默认值求值时可能为空（-File 调用）
if ([string]::IsNullOrEmpty($Root)) {
    $SelfDir = $PSScriptRoot
    if ([string]::IsNullOrEmpty($SelfDir)) { $SelfDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $Root = Split-Path -Parent $SelfDir
}
if ([string]::IsNullOrEmpty($Root)) { $Root = (Get-Location).Path }

$ErrorActionPreference = 'Stop'
. (Join-Path $Root 'lib\ics.ps1')

$dataDir = Join-Path $Root 'data'
$outDir  = Join-Path $Root 'out'
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

function Get-VeventBlocks {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $lines = Get-Content -LiteralPath $Path -Encoding UTF8
    $blocks = New-Object System.Collections.Generic.List[string]
    $cur = $null
    foreach ($l in $lines) {
        if ($l -eq 'BEGIN:VEVENT') { $cur = New-Object System.Collections.Generic.List[string]; continue }
        if ($l -eq 'END:VEVENT') {
            if ($null -ne $cur) { $blocks.Add(($cur -join "`r`n")) }
            $cur = $null
            continue
        }
        if ($null -ne $cur) { $cur.Add($l) }
    }
    return $blocks
}

# ==================== 科目标注引擎 ====================
# Blackboard 源不含任何科目信息，靠 config\courses.psd1 + config\uid-map.psd1 补上。
$CoursesConfigPath = Join-Path $Root 'config\courses.psd1'
$UidMapPath        = Join-Path $Root 'config\uid-map.psd1'

$courseNames = @{}
$labelRules  = @()
$unmatched   = 'UNMAPPED-COURSE'
$emitCsv     = $true
if (Test-Path -LiteralPath $CoursesConfigPath) {
    try {
        $cc = Import-PowerShellDataFile -LiteralPath $CoursesConfigPath
        if ($cc.ContainsKey('Courses') -and $cc['Courses']) {
            foreach ($k in $cc['Courses'].Keys) { $courseNames[[string]$k] = [string]$cc['Courses'][$k] }
        }
        if ($cc.ContainsKey('Rules') -and $cc['Rules']) { $labelRules = @($cc['Rules']) }
        if ($cc.ContainsKey('UnmatchedLabel') -and $cc['UnmatchedLabel']) { $unmatched = [string]$cc['UnmatchedLabel'] }
        if ($cc.ContainsKey('EmitCsv')) { $emitCsv = [bool]$cc['EmitCsv'] }
    } catch { Write-Output ('  [warn] courses.psd1 解析失败: ' + $_.Exception.Message) }
}
$uidMap = @{}
if (Test-Path -LiteralPath $UidMapPath) {
    try {
        $um = Import-PowerShellDataFile -LiteralPath $UidMapPath
        foreach ($k in $um.Keys) { $uidMap[[string]$k] = [string]$um[$k] }
    } catch { Write-Output ('  [warn] uid-map.psd1 解析失败: ' + $_.Exception.Message) }
}

# ==================== 旧学期过滤 + 标题屏蔽 ====================
# Blackboard 源会长期携带历史学期事件；按 config.psd1 的 HideBeforeDate 过滤掉。
# HideTitles 用于永久屏蔽某些条目（如非学分的普通话网课申请）。
$hideBefore = $null
$hideTitles = @()
$mainConfigPath = Join-Path $Root 'config\config.psd1'
if (Test-Path -LiteralPath $mainConfigPath) {
    try {
        $mc = Import-PowerShellDataFile -LiteralPath $mainConfigPath
        if ($mc.ContainsKey('HideBeforeDate') -and -not [string]::IsNullOrWhiteSpace([string]$mc['HideBeforeDate'])) {
            $hideBefore = [datetime]::ParseExact([string]$mc['HideBeforeDate'], 'yyyy-MM-dd', $null)
        }
        if ($mc.ContainsKey('HideTitles') -and $mc['HideTitles']) { $hideTitles = @($mc['HideTitles']) }
    } catch { Write-Output ('  [warn] config.psd1 解析失败: ' + $_.Exception.Message) }
}

function Test-TitleHidden {
    <# 标题命中 HideTitles（通配符 * ?）则返回 $true #>
    param([string]$Title)
    if (-not $Title) { return $false }
    foreach ($pat in $hideTitles) {
        if (-not [string]::IsNullOrWhiteSpace($pat) -and $Title -like $pat) { return $true }
    }
    return $false
}

function Get-BlockStartLocalDate {
    <# 从 VEVENT 块取出本地（UTC+8）开始日期，供旧学期过滤使用 #>
    param([string]$Block)
    $line = ''
    foreach ($l in ($Block -split "`r`n")) {
        if ($l -like 'DTSTART*') { $line = $l; break }
    }
    if (-not $line -or $line -notmatch '(\d{8})') { return $null }
    try { return [datetime]::ParseExact($Matches[1], 'yyyyMMdd', $null) } catch { return $null }
}

function Get-BlockProp {
    param([string[]]$Lines, [string]$Name)
    foreach ($l in $Lines) {
        if ($l -like ($Name + ':*')) { return $l.Substring($Name.Length + 1) }
        if ($l -like ($Name + ';*')) { $i = $l.IndexOf(':'); if ($i -ge 0) { return $l.Substring($i + 1) } }
    }
    return ''
}

function Get-CourseForEvent {
    param([string]$Uid, [string]$Title)
    if ($Uid -and $uidMap.ContainsKey($Uid)) { return $uidMap[$Uid] }        # 1) UID 精确映射（最可靠）
    foreach ($r in $labelRules) {                                            # 2) 标题规则，首个命中生效
        if (-not $r -or -not $r.ContainsKey('Pattern') -or -not $r.ContainsKey('Course')) { continue }
        $pat = [string]$r['Pattern']
        $mode = if ($r.ContainsKey('Match')) { [string]$r['Match'] } else { 'Like' }
        $hit = $false
        if ($mode -eq 'Regex') { $hit = ($Title -match $pat) }
        else { $hit = ($Title -like $pat) }
        if ($hit) { return [string]$r['Course'] }
    }
    return ''
}

function Get-DisplayCourse {
    param([string]$Code)
    if ([string]::IsNullOrEmpty($Code)) { return $unmatched }
    if ($courseNames.ContainsKey($Code)) { return ($Code + ' ' + $courseNames[$Code]) }
    return $Code
}

function Add-CourseLabelToBlock {
    param([string]$Block)
    $lines = @($Block -split "`r`n")
    $uid   = Get-BlockProp -Lines $lines -Name 'UID'
    $title = Get-BlockProp -Lines $lines -Name 'SUMMARY'
    $code  = Get-CourseForEvent -Uid $uid -Title $title
    $label = Get-DisplayCourse -Code $code

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($l -like 'SUMMARY:*') {
            $raw = $l.Substring(8)
            if ($code) {
                # 代号优先：'CS101: T3 - Repetition ...'
                # 若原标题已以课程名开头（部分课程如此），先剥掉，避免重复
                $name = ''
                if ($courseNames.ContainsKey($code)) { $name = [string]$courseNames[$code] }
                $t = $raw
                if ($name) { $t = $t -replace ('^\s*' + [regex]::Escape($name) + '\s*[:\-\u2013]?\s*'), '' }
                if ([string]::IsNullOrWhiteSpace($t)) { $t = $raw }
                $out.Add('SUMMARY:' + $code + ': ' + $t)
            } else {
                # 未标注：明确前缀，提示去标注
                $out.Add('SUMMARY:' + $unmatched + ': ' + $raw)
            }
        } elseif ($l -like 'DESCRIPTION:*') {
            if ($code) { $out.Add($l + '\nCourse: ' + $label) }
            else { $out.Add($l + '\nCourse: (unlabelled - run scripts\label-events.ps1)') }
        } else {
            $out.Add($l)
        }
    }
    if ($code) { $out.Add('CATEGORIES:' + $label) }
    return ($out -join "`r`n")
}

$header = @(
    'BEGIN:VCALENDAR'
    'VERSION:2.0'
    'PRODID:-//HKCC Reminder//Merged//EN'
    'CALSCALE:GREGORIAN'
    'METHOD:PUBLISH'
    'X-WR-CALNAME:HKCC 课程与截止日期（自动更新）'
    'X-WR-TIMEZONE:Asia/Hong_Kong'
)

# 抽取一份权威 VTIMEZONE（Blackboard 源自带的定义最完整），否则 Outlook 可能按本地时区猜
foreach ($cand in @((Join-Path $dataDir 'blackboard.ics'), (Join-Path $dataDir 'timetable.ics'))) {
    if (-not (Test-Path -LiteralPath $cand)) { continue }
    $raw = Get-Content -LiteralPath $cand -Raw -Encoding UTF8
    $m = [regex]::Match($raw, '(?ms)^BEGIN:VTIMEZONE\r?\n.*?^END:VTIMEZONE\r?$')
    if ($m.Success) {
        $vt = ($m.Value -split "\r?\n")
        $header += $vt
        Write-Output ("  VTIMEZONE 取自 " + (Split-Path -Leaf $cand) + "：" + $vt.Count + " 行")
        break
    }
}

$bbPath = Join-Path $dataDir 'blackboard.ics'
$ttPath = Join-Path $dataDir 'timetable.ics'
$bbBlocks = Get-VeventBlocks -Path $bbPath
$ttBlocks = Get-VeventBlocks -Path $ttPath

# ==================== 调课覆盖（来自 Blackboard 公告）====================
# 公告不会进日历源，必须手工登记；这里把 Exclude（EXDATE 跳过某一次）
# 和 Add（补课/调课新事件）应用到课表事件上。
$OverridesPath = Join-Path $Root 'config\overrides.psd1'
$overrides = @()
if (Test-Path -LiteralPath $OverridesPath) {
    try {
        $ov = Import-PowerShellDataFile -LiteralPath $OverridesPath
        if ($ov.ContainsKey('Overrides') -and $ov['Overrides']) { $overrides = @($ov['Overrides']) }
    } catch { Write-Output ('  [warn] overrides.psd1 解析失败: ' + $_.Exception.Message) }
}

function Add-IcsLine {
    param([string]$Block, [string]$Line)
    return ($Block.TrimEnd("`r`n") + "`r`n" + $Line)
}

function Apply-OverridesToTimetable {
    <#
        返回 @{ Blocks = <处理后的课表块>; Added = <新增的一次性事件块>; Excluded = <跳过次数> }
    #>
    param([string[]]$Blocks, $Overrides)

    $out = New-Object System.Collections.Generic.List[string]
    $added = New-Object System.Collections.Generic.List[string]
    $excludedCount = 0

    foreach ($b in $Blocks) {
        $lines = @($b -split "`r`n")
        $sum   = [string](Get-BlockProp -Lines $lines -Name 'SUMMARY')
        $desc  = [string](Get-BlockProp -Lines $lines -Name 'DESCRIPTION')
        $uid   = [string](Get-BlockProp -Lines $lines -Name 'UID')
        $newBlock = $b

        foreach ($o in $Overrides) {
            if (-not $o -or -not $o.ContainsKey('Action')) { continue }
            if ([string]$o['Action'] -ne 'Exclude') { continue }
            if (-not $o.ContainsKey('Match') -or -not $o.ContainsKey('Date')) { continue }

            if ($sum -notlike ('*' + [string]$o['Match'] + '*')) { continue }
            # 可选 Kind 过滤：课型（Lecture/Tutorial/Laboratory）在 DESCRIPTION 里，不在 SUMMARY
            if ($o.ContainsKey('Kind') -and $o['Kind']) {
                $k = [string]$o['Kind']
                if (($sum -notlike ('*' + $k + '*')) -and ($desc -notlike ('*' + $k + '*'))) { continue }
            }
            $d = ([string]$o['Date']) -replace '-', ''
            if ($d -notmatch '^\d{8}$') { continue }
            # 构造 EXDATE：必须精确匹配该次事件的开始"时刻"（不能只写当天 00:00），
            # 否则 RRULE 展开时匹配不上，排除不会生效。
            # 本项目的重复课表事件每周同一天同一时间，所以直接取 DTSTART 的 HHMMSS。
            $dtLine = ($lines | Where-Object { $_ -like 'DTSTART*' } | Select-Object -First 1)
            $hhmmss = '000000'
            if ($dtLine -match 'T(\d{6})\s*$') { $hhmmss = $Matches[1] }
            $exLine = if ($dtLine -match 'TZID=([^:;]+):') { 'EXDATE;TZID=' + $Matches[1] + ':' + $d + 'T' + $hhmmss }
                      else { 'EXDATE:' + $d + 'T' + $hhmmss }
            $newBlock = Add-IcsLine -Block $newBlock -Line $exLine
            $excludedCount++
            Write-Output ('  调课覆盖[跳过一次]: ' + $sum + '  ' + $o['Date'] + $(if ($o.ContainsKey('Note')) { '  (' + $o['Note'] + ')' } else { '' }))
        }
        $out.Add($newBlock)
    }

    # 生成新增事件
    foreach ($o in $Overrides) {
        if (-not $o -or -not $o.ContainsKey('Action') -or [string]$o['Action'] -ne 'Add') { continue }
        if ($o.ContainsKey('Enabled') -and -not [bool]$o['Enabled']) { continue }
        if (-not $o.ContainsKey('Date') -or -not $o.ContainsKey('Start') -or -not $o.ContainsKey('End')) { continue }
        $code = if ($o.ContainsKey('Code')) { [string]$o['Code'] } else { 'OVERRIDE' }
        $sum  = if ($o.ContainsKey('Summary')) { [string]$o['Summary'] } else { 'Make-up Class' }
        $d    = ([string]$o['Date']) -replace '-', ''
        $sRaw = ([string]$o['Start']) -replace ':', ''
        $eRaw = ([string]$o['End'])   -replace ':', ''
        if ($d -notmatch '^\d{8}$' -or $sRaw -notmatch '^\d{4}$' -or $eRaw -notmatch '^\d{4}$') { continue }
        $loc  = if ($o.ContainsKey('Location')) { [string]$o['Location'] } else { '' }
        $desc = if ($o.ContainsKey('Description')) { [string]$o['Description'] } else { 'From Blackboard announcement' }

        # 查重：用"已展开的实例"判断该日期是否已有同课程事件。
        # 不能只比对静态 DTSTART —— 第 4 周的课是 RRULE 展开出来的，
        # 静态 DTSTART 往往是第 1 周，永远匹配不上。
        $dup = $false
        foreach ($b in $Blocks) {
            $ls = @($b -split "`r`n")
            $bSum = [string](Get-BlockProp -Lines $ls -Name 'SUMMARY')
            if ($bSum -notlike ('*' + $code + '*')) { continue }
            $bEv = [pscustomobject]@{
                Summary     = $bSum
                Location    = [string](Get-BlockProp -Lines $ls -Name 'LOCATION')
                Description = ''
                Uid         = [string](Get-BlockProp -Lines $ls -Name 'UID')
                DtStartUtc  = Convert-IcsDate -Value ([string](Get-BlockProp -Lines $ls -Name 'DTSTART')) `
                                              -Tzid (Get-IcsTzidOf -Lines $ls)
                DtEndUtc    = Convert-IcsDate -Value ([string](Get-BlockProp -Lines $ls -Name 'DTEND')) `
                                              -Tzid (Get-IcsTzidOf -Lines $ls)
                Rrule       = [string](Get-BlockProp -Lines $ls -Name 'RRULE')
                ExDatesUtc  = @()
            }
            if ($null -eq $bEv.DtStartUtc) { continue }
            if ($null -eq $bEv.DtEndUtc)   { $bEv.DtEndUtc = $bEv.DtStartUtc }
            foreach ($occ in (Expand-IcsEvent -Event $bEv -OffsetHours 8)) {
                $occLocal = $occ.StartUtc.AddHours(8)
                if ($occLocal.ToString('yyyyMMdd') -eq $d) { $dup = $true; break }
            }
            if ($dup) { break }
        }
        if ($dup -and -not ($o.ContainsKey('AllowDuplicate') -and [bool]$o['AllowDuplicate'])) {
            Write-Output ('  调课覆盖[跳过重复]: ' + $code + ' ' + $o['Date'] + ' ' + $o['Start'] + ' —— 该时段已有同课程事件（如需强制保留请加 AllowDuplicate = $true）')
            continue
        }

        $evt = New-Object System.Collections.Generic.List[string]
        $evt.Add('UID:hkcc-override-' + $d + 'T' + $sRaw + '-' + $code + '@hkcc.local')
        $evt.Add('DTSTAMP:20260101T000000Z')
        $evt.Add('SUMMARY:' + $code + ': ' + $sum)
        $evt.Add('DTSTART;TZID=Asia/Hong_Kong:' + $d + 'T' + $sRaw + '00')
        $evt.Add('DTEND;TZID=Asia/Hong_Kong:' + $d + 'T' + $eRaw + '00')
        if ($loc)  { $evt.Add('LOCATION:' + $loc) }
        $evt.Add('DESCRIPTION:' + $desc)
        $evt.Add('CATEGORIES:' + $code + ' OVERRIDE')
        $added.Add(($evt -join "`r`n"))
        Write-Output ('  调课覆盖[新增事件]: ' + $code + ' ' + $sum + '  ' + $o['Date'] + ' ' + $o['Start'] + '-' + $o['End'] + $(if ($loc) { '  @ ' + $loc } else { '' }))
    }

    return [pscustomobject]@{ Blocks = $out.ToArray(); Added = $added.ToArray(); Excluded = $excludedCount }
}

$ovResult = Apply-OverridesToTimetable -Blocks $ttBlocks -Overrides $overrides
$ttBlocks = $ovResult.Blocks
$overrideBlocks = @($ovResult.Added | Where-Object { $_ })
if ($overrides.Count -gt 0) {
    Write-Output ('  调课覆盖: 共 ' + $overrides.Count + ' 条规则，跳过 ' + $ovResult.Excluded + ' 次，新增 ' + $overrideBlocks.Count + ' 个事件')
}

# ---------- 合并（TT 在前，BB 在后）----------
$seenUid = @{}
$merged = New-Object System.Collections.Generic.List[string]
foreach ($b in @($ttBlocks + $overrideBlocks + $bbBlocks)) {
    $uidLine = ($b -split "`r`n") | Where-Object { $_ -like 'UID:*' } | Select-Object -First 1
    $uid = if ($uidLine) { $uidLine.Substring(4) } else { '' }
    if ($uid -and $seenUid.ContainsKey($uid)) { continue }
    if ($uid) { $seenUid[$uid] = $true }
    $merged.Add($b)
}

function Convert-BlocksToUtc {    <#
        把 DTSTART/DTEND 的 TZID 形式统一转成 UTC(Z)。
        Asia/Hong_Kong 恒为 UTC+8（1979 年后无夏令时），转换无歧义；
        这样交付的 .ics 不依赖任何时区库，Outlook 不会报"未知时区"或错位。
    #>
    param([string[]]$Blocks)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($b in $Blocks) {
        $newLines = New-Object System.Collections.Generic.List[string]
        foreach ($line in ($b -split "`r`n")) {
            if ($line -match '^(DTSTART|DTEND)(;TZID=([^:;]+))?:(\d{8}T\d{6})$') {
                $prop = $Matches[1]
                $tzid = $Matches[3]
                $val = $Matches[4]
                $offset = Get-TzOffsetHours -Tzid $tzid
                try {
                    $dt = [datetime]::ParseExact($val, 'yyyyMMddTHHmmss', $null)
                    $utc = $dt.AddHours(-$offset)
                    $newLines.Add($prop + ':' + $utc.ToString('yyyyMMddTHHmmss') + 'Z')
                    continue
                } catch { }
            }
            $newLines.Add($line)
        }
        $out.Add(($newLines -join "`r`n"))
    }
    return $out
}


function Write-IcsFile {
    param([string]$Path, [string[]]$Blocks, [string[]]$ExtraHeader = @())
    if ($Blocks.Count -eq 0) { return 0 }
    # 先标注科目（仅对 Blackboard 事件生效；课表事件 SUMMARY 已含代号，原样保留）
    $labeled = New-Object System.Collections.Generic.List[string]
    foreach ($b in $Blocks) {
        if ($b -match 'UID:_blackboard\.') { $labeled.Add((Add-CourseLabelToBlock -Block $b)) }
        else { $labeled.Add($b) }
    }
    $Blocks = Convert-BlocksToUtc -Blocks $labeled.ToArray()
    $content = @()
    $content += $header
    $content += $ExtraHeader
    foreach ($b in $Blocks) {
        $content += 'BEGIN:VEVENT'
        $content += ($b -split "`r`n")
        $content += 'END:VEVENT'
    }
    $content += 'END:VCALENDAR'
    # RFC 5545 折行（75 字节上限），避免长标题被严格解析器截断
    $folded = New-Object System.Collections.Generic.List[string]
    foreach ($l in $content) {
        foreach ($piece in (Format-IcsLine -Line $l)) { $folded.Add($piece) }
    }
    Set-Content -LiteralPath $Path -Value $folded -Encoding UTF8
    return $Blocks.Count
}

# ==================== 分类、旧学期过滤与导出 ====================
# Blackboard 的 gradebook 项 = 作业/测验截止；其余（如 Examination Day）= 考试
$bbDeadlines = @(); $bbExams = @(); $bbHidden = 0; $ttHidden = 0; $bbTitleHidden = 0
foreach ($b in $bbBlocks) {
    $ls0 = @($b -split "`r`n")
    $t0  = Get-BlockProp -Lines $ls0 -Name 'SUMMARY'
    if (Test-TitleHidden -Title $t0) { $bbTitleHidden++; continue }
    $sd = Get-BlockStartLocalDate -Block $b
    if ($null -ne $hideBefore -and $null -ne $sd -and $sd -lt $hideBefore) { $bbHidden++; continue }
    if ($b -match 'UID:_blackboard\.platform\.gradebook2\.') { $bbDeadlines += $b } else { $bbExams += $b }
}
$bbKept = @($bbDeadlines + $bbExams)

# ==================== 手工登记的截止（日历源里没有的）====================
# 教训：有些作业只出现在 Blackboard Content 页，日历源里没有它 —— 而 data\blackboard.ics 每天定时会被
# Blackboard 的 Content 页，日历源里没有它 —— 而 data\blackboard.ics 每天定时会被
# sync-feed.ps1 用源文件整体覆盖，所以"直接改 blackboard.ics"不是持久方案。
# 登记处：config\extra-deadlines.psd1（ASCII only，字段见文件头注释）。
$extraDeadlineBlocks = @()
$extraPath = Join-Path $Root 'config\extra-deadlines.psd1'
if (Test-Path -LiteralPath $extraPath) {
    try {
        $ex = Import-PowerShellDataFile -LiteralPath $extraPath
        foreach ($e in @($ex['Deadlines'])) {
            if (-not $e -or -not $e.ContainsKey('Due')) { continue }
            if ($e.ContainsKey('Enabled') -and -not $e['Enabled']) { continue }
            $due  = [datetime]::Parse([string]$e['Due'], [Globalization.CultureInfo]::InvariantCulture)
            $code = [string]$e['Code']
            $sum  = [string]$e['Summary']
            $name = ''
            if ($courseNames.ContainsKey($code)) { $name = [string]$courseNames[$code] }
            $label = $code
            if ($name) { $label = $code + ' ' + $name }
            $key = $code + '-' + $due.ToString('yyyyMMddTHHmmss')
            if ($e.ContainsKey('Key') -and $e['Key']) { $key = [string]$e['Key'] }
            $uid = 'hkcc-extra-deadline-' + $key + '@hkcc.local'
            # 注意：不要在这里加 BEGIN:VEVENT/END:VEVENT —— Write-IcsFile 会统一包一层。
            $evt = New-Object System.Collections.Generic.List[string]
            $evt.Add('DTSTAMP:' + (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmss') + 'Z')
            $evt.Add('DTSTART;TZID=Asia/Hong_Kong:' + $due.ToString('yyyyMMddTHHmmss'))
            $evt.Add('DTEND;TZID=Asia/Hong_Kong:'   + $due.ToString('yyyyMMddTHHmmss'))
            $evt.Add('SUMMARY:' + $code + ': ' + $sum + ' (Due: ' + $due.ToString('yyyy-MM-dd HH:mm') + ')')
            $desc = 'Manually registered in config\extra-deadlines.psd1 (not in the Blackboard calendar feed).'
            if ($e.ContainsKey('Description') -and $e['Description']) { $desc = [string]$e['Description'] + ' | ' + $desc }
            $evt.Add('DESCRIPTION:' + $desc)
            $evt.Add('CATEGORIES:' + $label)
            $evt.Add('UID:' + $uid)
            $extraDeadlineBlocks += ($evt -join "`r`n")
        }
        if ($extraDeadlineBlocks.Count -gt 0) {
            Write-Output ('  手工登记截止: 加入 ' + $extraDeadlineBlocks.Count + ' 条（config\extra-deadlines.psd1）')
        }
    } catch {
        Write-Output ('  [warn] extra-deadlines.psd1 解析失败: ' + $_.Exception.Message)
    }
}
$bbKept = @($bbDeadlines + $bbExams + $extraDeadlineBlocks)
$ttKept = @()
foreach ($b in $ttBlocks) {
    $sd = Get-BlockStartLocalDate -Block $b
    if ($null -ne $hideBefore -and $null -ne $sd -and $sd -lt $hideBefore) { $ttHidden++; continue }
    $ttKept += $b
}
$mergedKept = @($ttKept) + @($overrideBlocks) + @($bbKept)

if ($null -ne $hideBefore) {
    Write-Output ('  旧学期过滤: 保留 ' + $bbKept.Count + ' / ' + $bbBlocks.Count + ' 条 Blackboard 事件，隐藏 ' + $bbHidden + ' 条（早于 ' + $hideBefore.ToString('yyyy-MM-dd') + '）')
} else {
    Write-Output '  旧学期过滤: 未启用（保留全部事件）'
}
if ($bbTitleHidden -gt 0) {
    Write-Output ('  标题屏蔽: 隐藏 ' + $bbTitleHidden + ' 条（HideTitles）')
}

$nMerged = Write-IcsFile -Path (Join-Path $outDir 'HKCC_merged.ics')    -Blocks $mergedKept
# 调课新增的事件本质上是"课程事件"，必须一并写进 classes.ics，
# 否则邮件摘要（读 classes.ics）看不到当天/本周新增的补课与调课。
$nClass  = Write-IcsFile -Path (Join-Path $outDir 'HKCC_classes.ics')   -Blocks (@($ttKept) + @($overrideBlocks))
$nDead   = Write-IcsFile -Path (Join-Path $outDir 'HKCC_deadlines.ics') -Blocks $bbKept

Write-Output ("  合并导出: merged=" + $nMerged + " (classes=" + $nClass + ", deadlines=" + $nDead + ") -> " + $outDir)

# ==================== 科目标注状况报表 ====================
$report = New-Object System.Collections.Generic.List[object]
foreach ($b in $bbKept) {
    $ls    = @($b -split "`r`n")
    $uid   = Get-BlockProp -Lines $ls -Name 'UID'
    $title = Get-BlockProp -Lines $ls -Name 'SUMMARY'
    $dt    = Get-BlockProp -Lines $ls -Name 'DTSTART'
    # 先看 SUMMARY 是否已带课程代号前缀（手工登记的截止自带 'CS101: ...'，
    # 补课/调课事件自带 'EE201: ...'）—— 这类事件不在 UID 映射与标题规则里，
    # 只看 Get-CourseForEvent 会被误判成"未标注"。
    $code = ''
    if ($title -match '^([A-Za-z][A-Za-z0-9\-]*)\s*:') { $code = $Matches[1] }
    if (-not $code) { $code = Get-CourseForEvent -Uid $uid -Title $title }
    $when  = ''
    $dtv = $dt -replace 'Z$', ''
    if ($dtv -match '(\d{8})T(\d{6})') {
        try { $when = [datetime]::ParseExact($Matches[1] + $Matches[2], 'yyyyMMddHHmmss', $null).ToString('yyyy-MM-dd HH:mm') } catch { $when = $dt }
    }
    $report.Add([pscustomobject]@{
        Uid        = $uid
        Due        = $when
        Label      = (Get-DisplayCourse -Code $code)
        CourseCode = $code
        Title      = $title
    })
}
$unlabeled = @($report | Where-Object { -not $_.CourseCode })
$unlabeledTitles = @($unlabeled | Group-Object Title)
Write-Output ('  科目标注: 已标注 ' + ($report.Count - $unlabeled.Count) + ' / ' + $report.Count + ' 条；未标注 ' + $unlabeled.Count + ' 条，涉及 ' + $unlabeledTitles.Count + ' 个不同标题')
if ($unlabeledTitles.Count -gt 0) {
    Write-Output '  未标注标题（运行 scripts\label-events.ps1 可一次性标注）:'
    foreach ($g in ($unlabeledTitles | Sort-Object Name)) { Write-Output ('    - ' + $g.Name) }
}
if ($emitCsv) {
    $csvPath = Join-Path $outDir 'label-report.csv'
    $report | Sort-Object Due | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    Write-Output ('  报表: ' + $csvPath)
}
