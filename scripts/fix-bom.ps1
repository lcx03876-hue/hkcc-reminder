<#
    fix-bom.ps1 - 给本目录下所有 .ps1 补 UTF-8 BOM
    原因：Windows PowerShell 5.1 会把"无 BOM 的 UTF-8"当成 GBK 读取，
          脚本里的中文注释/字符串会变成乱码并导致语法错误。
    任何编辑过 .ps1 之后，运行本脚本即可修回。
#>
$root = Split-Path -Parent $PSScriptRoot
$fixed = 0
Get-ChildItem -Path $root -Recurse -Filter *.ps1 | ForEach-Object {
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191)
    if (-not $hasBom) {
        $text = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8
        [System.IO.File]::WriteAllText($_.FullName, $text, (New-Object System.Text.UTF8Encoding($true)))
        Write-Host ("BOM added: " + $_.Name) -ForegroundColor Green
        $fixed++
    }
}
if ($fixed -eq 0) { Write-Host '所有 .ps1 均已带 BOM，无需修改。' -ForegroundColor Cyan }
