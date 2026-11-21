<#
    render-v14.ps1 - 每日提醒邮件的新版排版（深色 / 分级）
    约定：本文件只管排版；数据由调用方传入（均为 ICS 行对象）。
#>

function Get-V14Strings {
    param([string]$Lang = 'en')
    $L = @{
        'en' = @{
            dueSoon = 'DUE SOON'; todayTmrw = 'TODAY & TOMORROW'; further = 'FURTHER OUT · 7-14 DAYS'
            today = 'Today'; tomorrow = 'Tomorrow'; noClasses = 'No classes.'; ended = 'ended'
            nothing48 = 'Nothing else within 48 hours.'; nothing7 = 'Nothing due in the next 7 days.'
            attachNote = 'Everything else is in the attached HkccDeadlines.ics.'
            footer1 = 'Built from your own Blackboard calendar feed and timetable. Everything stays on this machine.'
            footer2 = 'No fetch errors today.'
            thing = 'thing'; things = 'things'; classes = 'classes'; dueThisWeek = 'due this week.'; todayHas = 'today.'
            inDays = 'in {n} days'; inDay = 'in {n} day'; inHours = 'in {n}h'
            changes = 'CHANGES'
        }
        'zh-CN' = @{
            dueSoon = '要交的东西'; todayTmrw = '今天和明天'; further = '更远的 · 7–14 天'
            today = '今天'; tomorrow = '明天'; noClasses = '没有课。'; ended = '已结束'
            nothing48 = '48 小时内没有别的了。'; nothing7 = '7 天内没有要交的。'
            attachNote = '其余的事都在附件 HkccDeadlines.ics 里。'
            footer1 = '数据来自你自己的 Blackboard 日历源与课表，全部在本机生成，不上传任何地方。'
            footer2 = '今天没有抓取失败。'
            thing = '件'; things = '件'; classes = '节课'; dueThisWeek = '要交。'; todayHas = '要上。'
            inDays = '{n} 天后'; inDay = '{n} 天后'; inHours = '{n} 小时后'
            changes = '变动'
        }
        'zh-TW' = @{
            dueSoon = '要交的東西'; todayTmrw = '今天和明天'; further = '更遠的 · 7–14 天'
            today = '今天'; tomorrow = '明天'; noClasses = '沒有課。'; ended = '已結束'
            nothing48 = '48 小時內沒有別的了。'; nothing7 = '7 天內沒有要交的。'
            attachNote = '其餘的事都在附件 HkccDeadlines.ics 裡。'
            footer1 = '資料來自你自己的 Blackboard 日曆源與課表，全部在本機產生，不上傳任何地方。'
            footer2 = '今天沒有抓取失敗。'
            thing = '件'; things = '件'; classes = '節課'; dueThisWeek = '要交。'; todayHas = '要上。'
            inDays = '{n} 天後'; inDay = '{n} 天後'; inHours = '{n} 小時後'
            changes = '變動'
        }
    }
    if (-not $L.ContainsKey($Lang)) { return $L['en'] }
    return $L[$Lang]
}

function Clean-V14Title {
    # 去掉标题里残留的 (Due Date: ...)、(deadline: ...) 之类的后缀
    param([string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return '' }
    $t = [regex]::Replace($Title, '\s*\((?:Due Date|Due date|Due|deadline)[^)]*\)', '')
    $t = [regex]::Replace($t, ':\s*Deadline.*$', '')
    $t = [regex]::Replace($t, '\s*\(PDF version only\)', '')
    $t = [regex]::Replace($t, '[,\s]*\b\d{1,2}:\d{2}\b\s*$', '')
    $t = [regex]::Replace($t, '\s*\[(?:by|By)[^\]]*\]', '')
    $t = [regex]::Replace($t, ',\s*submission of both.*$', '')
    return $t.Trim()
}

function Format-V14Day {
    param([datetime]$Date, [string]$Lang)
    if ($Lang -eq 'en') {
        $ci = [System.Globalization.CultureInfo]::GetCultureInfo('en-GB')
        return $Date.ToString('ddd d MMM', $ci)
    }
    $wd = @('周日', '周一', '周二', '周三', '周四', '周五', '周六')[[int]$Date.DayOfWeek]
    return ('{0:MM-dd} {1}' -f $Date, $wd)
}

function Get-V14Countdown {
    param([datetime]$Due, [datetime]$Now, [hashtable]$S)
    $hours = ($Due - $Now).TotalHours
    if ($hours -lt 0) { return '' }
    if ($hours -le 48) { return ($S.inHours -replace '\{n\}', [math]::Max(1, [math]::Round($hours))) }
    $days = [int]($Due.Date - $Now.Date).TotalDays
    $tpl = if ($days -le 1) { $S.inDay } else { $S.inDays }
    return ($tpl -replace '\{n\}', $days)
}

function Render-V14Html {
    param(
        [object[]]$Classes = @(),
        [object[]]$Deadlines = @(),
        [object[]]$Changes = @(),
        [datetime]$Now,
        [string]$Lang = 'en'
    )

    $S = Get-V14Strings -Lang $Lang

    $p0 = @(); $p1 = @(); $p3 = @()
    foreach ($d in $Deadlines) {
        $h = ($d.Start - $Now).TotalHours
        if ($h -lt 0) { continue }
        $days = [int]($d.Start.Date - $Now.Date).TotalDays
        if ($h -le 48) { $p0 += $d }
        elseif ($days -le 7) { $p1 += $d }
        elseif ($days -le 14) { $p3 += $d }
    }
    $p0 = @($p0 | Sort-Object Start)
    $p1 = @($p1 | Sort-Object Start)
    $p3 = @($p3 | Sort-Object Start)
    if ($p0.Count -gt 1) { $p1 = @($p0[1..($p0.Count - 1)]) + $p1; $p0 = @($p0[0]) }

    $today = @($Classes | Where-Object { $_.Start.Date -eq $Now.Date } | Sort-Object Start)
    $tmrw = @($Classes | Where-Object { $_.Start.Date -eq $Now.AddDays(1).Date } | Sort-Object Start)

    $dueCount = $p0.Count + $p1.Count
    $classCount = $today.Count + $tmrw.Count
    $amber = '#e0a861'
    $grey = '#9a958f'
    $font = 'font-family:-apple-system,BlinkMacSystemFont,sans-serif;'
    $gap = if ($Lang -eq 'en') { ' ' } else { '' }

    $parts = New-Object System.Collections.Generic.List[string]
    $add = { param($s) $parts.Add($s) }

    $add.Invoke('<!DOCTYPE html><html lang="' + $Lang + '"><head><meta charset="utf-8">')
    $add.Invoke('<meta name="viewport" content="width=device-width,initial-scale=1">')
    $add.Invoke('<meta name="color-scheme" content="dark">')
    $add.Invoke('<meta name="format-detection" content="telephone=no,date=no,address=no,email=no,url=no">')
    $add.Invoke('<meta name="x-apple-disable-message-reformatting">')
    $add.Invoke('<style>a[x-apple-data-detectors]{color:inherit!important;text-decoration:none!important;}</style>')
    $add.Invoke('</head><body style="margin:0;padding:0;background:#121214;">')
    $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#121214;"><tr><td align="center" style="padding:18px 10px;">')
    $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#18181b;border:1px solid #26262b;">')

    # 抬头
    $add.Invoke('<tr><td style="padding:26px 22px 2px 22px;' + $font + '">')
    $add.Invoke('<div style="font-size:12px;color:#8b8b93;letter-spacing:1.5px;">' + '<a href="#" style="color:#5aa9e6;text-decoration:underline;">' + (Format-V14Day -Date $Now -Lang $Lang) + '</a>' + '</div>')
    $add.Invoke('<div style="font-size:23px;line-height:1.4;color:#f4f1ec;font-weight:700;margin-top:8px;">')
    $dueWord = if ($dueCount -eq 1) { $S.thing } else { $S.things }
    $clsWord = $S.classes
    $add.Invoke('<span style="color:' + $amber + ';">' + $dueCount + $gap + $dueWord + '</span> ' + $S.dueThisWeek + '<br>')
    $add.Invoke('<span style="color:' + $grey + ';">' + $classCount + $gap + $clsWord + '</span> ' + $S.todayHas)
    $add.Invoke('</div></td></tr>')

    # 变动
    if (@($Changes).Count -gt 0) {
        $add.Invoke('<tr><td style="padding:0 22px 8px 22px;' + $font + '">')
        $add.Invoke('<div style="background:#2b1517;border-left:4px solid #e04a4a;padding:10px 12px;">')
        $add.Invoke('<div style="font-size:12px;color:#f08a8a;font-weight:700;letter-spacing:1px;">' + $S.changes + '</div>')
        foreach ($c in $Changes) {
            $add.Invoke('<div style="font-size:13px;color:#e8c9c9;margin-top:4px;">' + [System.Net.WebUtility]::HtmlEncode([string]$c.Summary) + '</div>')
        }
        $add.Invoke('</div></td></tr>')
    }

    # DUE SOON
    $add.Invoke('<tr><td style="padding:20px 22px 0 22px;' + $font + '">')
    $add.Invoke('<div style="font-size:12px;color:#8b8b93;letter-spacing:1.5px;font-weight:600;margin-bottom:10px;">' + $S.dueSoon + '</div>')

    $renderCard = {
        param($d, $urgent)
        $sp = Split-CoursePrefix -Summary $d.Summary
        $bar = if ($urgent) { '#e04a4a' } else { '#d98b3f' }
        $box = if ($urgent) { '#2b1517' } else { '#2b2316' }
        $tagBg = if ($urgent) { '#4a1a1d' } else { '#4a3418' }
        $tagFg = if ($urgent) { '#ffb3b3' } else { '#f0c48a' }
        $sub = if ($urgent) { '#d9a2a2' } else { '#b39a78' }
        $when = (Format-V14Day -Date $d.Start -Lang $Lang) + ' · ' + $d.Start.ToString('HH:mm')
        $cd = Get-V14Countdown -Due $d.Start -Now $Now -S $S
        $code = if ($sp.Code) { [System.Net.WebUtility]::HtmlEncode([string]$sp.Code) } else { '' }
        $title = [System.Net.WebUtility]::HtmlEncode((Clean-V14Title $sp.Title))

        $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + $box + ';border-left:4px solid ' + $bar + ';margin-bottom:12px;">')
        $add.Invoke('<tr><td style="padding:14px 14px 14px 12px;' + $font + '">')
        $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>')
        $add.Invoke('<td style="font-size:13px;font-weight:600;"><a href="#" style="color:#5aa9e6;text-decoration:underline;">' + $when + '</a></td>')
        $add.Invoke('<td align="right"><span style="display:inline-block;background:' + $tagBg + ';color:' + $tagFg + ';font-size:12px;font-weight:600;padding:3px 9px;border-radius:10px;">' + $cd + '</span></td>')
        $add.Invoke('</tr></table>')
        $add.Invoke('<div style="font-size:19px;line-height:1.4;color:#f6f1e8;font-weight:700;margin-top:6px;">' + $title + '</div>')
        $add.Invoke('<div style="font-size:13px;color:' + $sub + ';margin-top:4px;">' + $code + '</div>')
        $add.Invoke('</td></tr></table>')
    }

    foreach ($d in $p0) { $renderCard.Invoke($d, $true) }
    foreach ($d in $p1) { $renderCard.Invoke($d, $false) }

    $emptyNote = if ($dueCount -eq 0) { $S.nothing7 } else { $S.nothing48 }
    $add.Invoke('<div style="font-size:13px;color:#8b8b93;">' + $emptyNote + '</div></td></tr>')

    # TODAY & TOMORROW
    $add.Invoke('<tr><td style="padding:22px 22px 0 22px;' + $font + '">')
    $add.Invoke('<div style="font-size:12px;color:#8b8b93;letter-spacing:1.5px;font-weight:600;margin-bottom:10px;">' + $S.todayTmrw + '</div>')
    $add.Invoke('<div style="font-size:13px;color:#b9b9c0;font-weight:600;margin-bottom:6px;">' + $S.today + ' · ' + '<a href="#" style="color:#5aa9e6;text-decoration:underline;">' + (Format-V14Day -Date $Now -Lang $Lang) + '</a>' + '</div>')

    if ($today.Count -eq 0) {
        $add.Invoke('<div style="font-size:13px;color:#8b8b93;">' + $S.noClasses + '</div>')
    } else {
        $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14px;">')
        foreach ($c in $today) {
            $sp = Split-CoursePrefix -Summary $c.Summary
            $kind = if ($c.Kind) { ' ' + $c.Kind } else { '' }
            $place = if ($c.Location) { ' · ' + [System.Net.WebUtility]::HtmlEncode([string]$c.Location) } else { '' }
            $state = if ($c.Expired) { '<td align="right" style="padding:4px 0;color:#7c7c84;font-size:12px;">' + $S.ended + '</td>' } else { '<td></td>' }
            $add.Invoke('<tr><td width="128" style="padding:4px 0;"><a href="#" style="display:inline-block;background:#16222e;color:#5aa9e6;font-size:13px;font-weight:600;padding:3px 8px;border-radius:10px;text-decoration:underline;">' + $c.Start.ToString('HH:mm') + '–' + $c.End.ToString('HH:mm') + '</a></td>')
            $add.Invoke('<td style="padding:4px 0;color:#d8d8dd;">' + [System.Net.WebUtility]::HtmlEncode([string]$sp.Title) + $kind + $place + '</td>' + $state + '</tr>')
        }
        $add.Invoke('</table>')
    }

    $add.Invoke('<div style="font-size:13px;color:#b9b9c0;font-weight:600;margin:14px 0 4px 0;">' + $S.tomorrow + ' · ' + '<a href="#" style="color:#5aa9e6;text-decoration:underline;">' + (Format-V14Day -Date $Now.AddDays(1) -Lang $Lang) + '</a>' + '</div>')
    if ($tmrw.Count -eq 0) {
        $add.Invoke('<div style="font-size:13px;color:#8b8b93;">' + $S.noClasses + '</div>')
    } else {
        $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14px;">')
        foreach ($c in $tmrw) {
            $sp = Split-CoursePrefix -Summary $c.Summary
            $kind = if ($c.Kind) { ' ' + $c.Kind } else { '' }
            $place = if ($c.Location) { ' · ' + [System.Net.WebUtility]::HtmlEncode([string]$c.Location) } else { '' }
            $add.Invoke('<tr><td width="128" style="padding:4px 0;"><a href="#" style="display:inline-block;background:#16222e;color:#5aa9e6;font-size:13px;font-weight:600;padding:3px 8px;border-radius:10px;text-decoration:underline;">' + $c.Start.ToString('HH:mm') + '–' + $c.End.ToString('HH:mm') + '</a></td>')
            $add.Invoke('<td style="padding:4px 0;color:#d8d8dd;">' + [System.Net.WebUtility]::HtmlEncode([string]$sp.Title) + $kind + $place + '</td><td></td></tr>')
        }
        $add.Invoke('</table>')
    }
    $add.Invoke('</td></tr>')

    # FURTHER OUT
    if ($p3.Count -gt 0) {
        $add.Invoke('<tr><td style="padding:22px 22px 6px 22px;' + $font + '">')
        $add.Invoke('<div style="font-size:12px;color:#8b8b93;letter-spacing:1.5px;font-weight:600;margin-bottom:10px;">' + $S.further + '</div>')
        $add.Invoke('<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14px;line-height:1.9;">')
        foreach ($d in $p3) {
            $sp = Split-CoursePrefix -Summary $d.Summary
            $cd = Get-V14Countdown -Due $d.Start -Now $Now -S $S
            $line = '<a href="#" style="color:#5aa9e6;text-decoration:underline;">' + (Format-V14Day -Date $d.Start -Lang $Lang) + '</a> · <b>' + [System.Net.WebUtility]::HtmlEncode([string]$sp.Code) + '</b> ' + [System.Net.WebUtility]::HtmlEncode((Clean-V14Title $sp.Title)) + ' <span style="color:#7c7c84;">(' + $cd + ')</span>'
            $add.Invoke('<tr><td width="10" valign="top" style="padding-top:8px;"><span style="display:inline-block;width:6px;height:6px;border-radius:3px;background:#c9a06a;"></span></td>')
            $add.Invoke('<td style="color:#a9a9b1;">' + $line + '</td></tr>')
        }
        $add.Invoke('</table>')
        $add.Invoke('<div style="font-size:12px;color:#7c7c84;margin-top:8px;">' + $S.attachNote + '</div></td></tr>')
    }

    # 页脚
    $add.Invoke('<tr><td style="padding:16px 22px 20px 22px;border-top:1px solid #26262b;' + $font + '">')
    $add.Invoke('<div style="font-size:12px;line-height:1.8;color:#7c7c84;">' + $S.footer1 + '<br>' + $S.footer2 + '</div>')
    $add.Invoke('</td></tr>')
    $add.Invoke('</table></td></tr></table></body></html>')

    return ($parts -join "`n")
}
