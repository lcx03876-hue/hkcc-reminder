# HKCC daily reminder - configuration EXAMPLE
# Copy this file to config.psd1 and fill in your own values.
#
# Keep this file ASCII-only: Windows PowerShell 5.1 cannot reliably parse
# non-ASCII .psd1 files without a UTF-8 BOM.
#
# SMTP reference:
#   PolyU / CPCE  : smtp.office365.com  587
#   Gmail         : smtp.gmail.com      587  (App Password required)
#   QQ Mail       : smtp.qq.com         587  (authorization code required)
#   163 Mail      : smtp.163.com        465
@{
    SmtpServer    = 'smtp.office365.com'
    SmtpPort      = 587
    UseSsl        = $true
    From          = 'you@example.com'
    To            = @('you@example.com')

    # ---------------------------------------------------------------
    # Blackboard calendar feed
    #   Blackboard -> Calendar -> "Get external calendar link" / iCal.
    #   The URL contains a personal token; treat it like a password.
    #   Leave empty to skip the download and use data\blackboard.ics as-is.
    # ---------------------------------------------------------------
    FeedUrl       = ''

    SubjectPrefix = 'cpce'
    Days          = 7

    # Mail body language: zh-CN (default) / zh-TW / en.
    Language      = 'zh-CN'

    # ---------------------------------------------------------------
    # Reading mail (Microsoft Graph, OAuth2 device-code flow) -- optional
    #   Register a public client app in Entra (free) and paste its
    #   Application (client) ID below. Leave empty to disable mail reading.
    # ---------------------------------------------------------------
    ClientId    = ''
    Tenant      = 'common'
    GraphScopes = @('https://graph.microsoft.com/Mail.Read')
    MailLookbackHours = 24
    MailMaxMessages   = 40

    # Hide Blackboard events that started before this date (old semesters).
    # Format: yyyy-MM-dd (Asia/Hong_Kong). Set to '' to disable filtering.
    HideBeforeDate = ''

    # Permanently hide events whose title matches these wildcard patterns.
    HideTitles = @(
        # 'Application for LC0036*'
    )
}
