[CmdletBinding()]
param(
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Dist = Join-Path $Root 'dist'
$Zip = Join-Path $Dist ("hkcc-reminder-v$Version.zip")
$Prefix = "hkcc-reminder-v$Version/"

if (-not (Test-Path -LiteralPath (Join-Path $Root '.git'))) {
    throw "build-release.ps1 需要在 git 仓库里运行: $Root"
}

$dirty = & git -C $Root status --porcelain
if ($dirty) {
    Write-Warning '工作区有未提交改动；git archive 只会打包 HEAD，请先提交。'
}

if (-not (Test-Path -LiteralPath $Dist)) {
    New-Item -ItemType Directory -Path $Dist -Force | Out-Null
}
if (Test-Path -LiteralPath $Zip) { Remove-Item -LiteralPath $Zip -Force }

Push-Location $Root
try {
    & git archive --format=zip "--prefix=$Prefix" --output="$Zip" HEAD
    if ($LASTEXITCODE -ne 0) { throw 'git archive 失败' }
} finally {
    Pop-Location
}
if (-not (Test-Path -LiteralPath $Zip)) { throw '没有生成 zip' }

# ---- 校验：不能把个人配置/数据打进去 ----
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($Zip)
try {
    $names = @($archive.Entries | ForEach-Object { $_.FullName })
} finally {
    $archive.Dispose()
}

$bad = @($names | Where-Object {
    $_ -match '(^|/)(config\.psd1|smtp\.cred\.xml|graph\.cred\.xml|blackboard\.ics|timetable\.ics|feed-state\.json)$' -or
    $_ -match '(^|/)(logs|out|acl-recovery)/'
})
if ($bad.Count -gt 0) {
    throw ('zip 里出现不应打包的文件: ' + ($bad -join ', '))
}

$required = @(
    'install.bat',
    'install.ps1',
    'INSTALL.txt',
    'README.md',
    'config/config.example.psd1',
    'data/timetable.sample.ics',
    'scripts/daily-reminder.ps1'
)
foreach ($rel in $required) {
    if (-not ($names -contains ($Prefix + $rel))) {
        throw ('zip 缺少必要文件: ' + $rel)
    }
}

$hash = (Get-FileHash -LiteralPath $Zip -Algorithm SHA256).Hash
$hashFile = Join-Path $Dist ("hkcc-reminder-v$Version.sha256")
Set-Content -LiteralPath $hashFile -Value ($hash + '  hkcc-reminder-v' + $Version + '.zip') -Encoding ASCII

$sizeMb = [math]::Round((Get-Item -LiteralPath $Zip).Length / 1MB, 2)
Write-Host ''
Write-Host ('已生成: ' + $Zip) -ForegroundColor Green
Write-Host ('大小: ' + $sizeMb + ' MB · 文件数: ' + $names.Count)
Write-Host ('SHA256: ' + $hash)
Write-Host ('校验文件: ' + $hashFile)
