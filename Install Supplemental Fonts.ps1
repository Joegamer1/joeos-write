# Requires an administrator PowerShell. Downloads only Windows font capabilities.
$ErrorActionPreference = 'Stop'
$report = Join-Path $PSScriptRoot 'audit-output\supplemental-fonts.json'
$results = @()
@{ complete=$false; status='Checking available font packs'; fonts=@() } | ConvertTo-Json | Set-Content -LiteralPath $report -Encoding utf8
try {
    $fonts = @(Get-WindowsCapability -Online | Where-Object Name -Like 'Language.Fonts.*')
    foreach ($font in $fonts) {
        try {
            if ($font.State -ne 'Installed') {
                $result = Add-WindowsCapability -Online -Name $font.Name
                $state = (Get-WindowsCapability -Online -Name $font.Name).State.ToString()
                $results += @{ name=$font.Name; state=$state; restartNeeded=$result.RestartNeeded }
            } else { $results += @{ name=$font.Name; state='Installed'; restartNeeded=$false } }
        } catch { $results += @{ name=$font.Name; error=$_.Exception.Message } }
        @{ complete=$false; fonts=$results } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding utf8
    }
    @{ complete=$true; fonts=$results } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding utf8
} catch {
    @{ complete=$false; error=$_.Exception.Message; fonts=$results } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding utf8
    exit 1
}
