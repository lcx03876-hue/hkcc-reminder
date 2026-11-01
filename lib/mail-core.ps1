<#
    mail-core.ps1 - 读取收件箱（Microsoft Graph）+ 摘要 + 调课候选提取

    设计原则：全部基于确定性规则，不做"AI 猜测"。
      - 调课候选必须命中"调课类关键词"，并且尽量带上时间/日期线索
      - 候选只写入 pending 清单，绝不自动改动课表（fail-safe）
#>

function Get-MailMessages {
    <#
        取收件箱最近 $Hours 小时的邮件。
        返回: @{ Subject; From; FromAddress; Received; Preview; HasAttachments; WebLink; InternetId }
    #>
    param(
        [int]$Hours = 24,
        [int]$Max = 40,
        [string]$Token = '',
        [switch]$IncludeAllFolders
    )
    if ([string]::IsNullOrEmpty($Token)) { $Token = Get-GraphAccessToken }
    if ([string]::IsNullOrEmpty($Token)) { throw '未登录，无法读取邮件（先运行 scripts\mail-login.ps1）' }

    $since = (Get-Date).ToUniversalTime().AddHours(-$Hours).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $select = 'subject,from,receivedDateTime,bodyPreview,hasAttachments,webLink,internetMessageId'
    $uri = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?`$top=$Max&`$select=$select&`$orderby=receivedDateTime desc&`$filter=receivedDateTime ge $since"

    $msgs = @()
    $resp = Get-GraphJson -Uri $uri -Token $Token
    foreach ($m in $resp.value) {
        $addr = ''
        if ($m.from -and $m.from.emailAddress) { $addr = [string]$m.from.emailAddress.address }
        $msgs += [pscustomobject]@{
            Subject        = [string]$m.subject
            From           = if ($m.from -and $m.from.emailAddress) { [string]$m.from.emailAddress.name } else { '' }
            FromAddress    = $addr
            Received       = ([datetime]$m.receivedDateTime).ToLocalTime()
            Preview        = [string]$m.bodyPreview
            HasAttachments = [bool]$m.hasAttachments
            WebLink        = [string]$m.webLink
            InternetId     = [string]$m.internetMessageId
        }
    }
    return $msgs
}

function Test-ScheduleChangeMail {
    <#
        规则判定：这封邮件是否疑似"调课/改期/取消/换教室"通知。
        返回 @{ Hit=$true/$false; Reason='...'; Clues=@(...) }
        宁可多报（交给人工确认），不可漏报。
    #>
    param($Message)

    $text = ([string]$Message.Subject + ' ' + [string]$Message.Preview)
    $hits = @()

    # 1) 调课类关键词（中英 + 简繁）
    $kwPatterns = @(
        @{ p = '調課|调课|調堂|调堂|調時間|改期|改時間|更改時間|時間調動|時間變更|改上課|補課|补课|停課|停课|取消課|课堂取消|順延|顺延|延期'; r = '调课/改期类关键词' },
        @{ p = 'reschedul|postpone|cancel(l)?ed|class\s+cancel|change\s+of\s+(class|venue|room)|venue\s+change|room\s+change|make-?up\s+class'; r = '英文调课/改期关键词' },
        @{ p = 'timetable\s+change|schedule\s+change|revised\s+timetable|new\s+arrangement'; r = 'timetable/schedule change' }
    )
    foreach ($k in $kwPatterns) { if ($text -match $k.p) { $hits += $k.r } }

    # 2) 时间/教室线索（有线索说明更像具体安排变更）
    $cluePatterns = @(
        @{ p = '\b([01]?\d|2[0-3])[:.][0-5]\d\b'; r = '含具体时间' },
        @{ p = '星期[一二三四五六日天]|週[一二三四五六日天]|周[一二三四五六日天]|\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\b'; r = '含星期' },
        @{ p = '\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?|\d{4}年\d{1,2}月\d{1,2}日'; r = '含日期' },
        @{ p = '課室|教室|演講廳|演讲厅|實驗室|实验室|room|venue|lecture\s+theatre|\bWK-|\bN\d{3}\b|\bS\d{3}\b'; r = '含地点/课室' }
    )
    $clues = @()
    foreach ($c in $cluePatterns) { if ($text -match $c.p) { $clues += $c.r } }

    # 判定：命中调课关键词，且至少有一个时间/地点线索（或关键词非常明确）
    $strongKw = ($text -match '調課|调课|調堂|reschedul|class\s+cancel|timetable\s+change|schedule\s+change')
    $hit = ($hits.Count -gt 0) -and (($clues.Count -gt 0) -or $strongKw)

    return [pscustomobject]@{ Hit = $hit; Reason = ($hits -join ' + '); Clues = $clues }
}

function Get-MailSummary {
    <#
        对邮件列表做摘要分类（确定性规则）：
          action   需要你行动（截止/提交/回复/缴费/报名）——最需要关注
          change   疑似调课/改期（待确认）
          notice   一般通知（公告/活动/系统通知）
        返回按类别分组的对象
    #>
    param($Messages)

    $actionPattern = '截止|deadline|due\s|due\b|提交|繳交|缴交|submission|submit|回覆|回复|reply\s+by|確認|确认|confirm|報名|报名|register|enrol|缴费|繳費|payment|fees?\b|申請|申请|apply|interview|面試|面试'
    $noticePattern = '通知|公告|notice|announcement|newsletter|活動|活动|event|seminar|講座|讲座|webinar|圖書館|图书馆|library|maintenance|維護|维护'

    $action = @(); $change = @(); $notice = @(); $other = @()
    foreach ($m in $Messages) {
        $chg = Test-ScheduleChangeMail -Message $m
        if ($chg.Hit) { $change += [pscustomobject]@{ Msg = $m; Reason = $chg.Reason; Clues = $chg.Clues }; continue }
        $t = [string]$m.Subject + ' ' + [string]$m.Preview
        if ($t -match $actionPattern) { $action += $m; continue }
        if ($t -match $noticePattern) { $notice += $m; continue }
        $other += $m
    }
    return [pscustomobject]@{
        Action = $action
        Change = $change
        Notice = $notice
        Other  = $other
        Total  = @($Messages).Count
    }
}
