. "$PSScriptRoot\..\src\QuotaRecon.Auth.ps1"
. "$PSScriptRoot\..\src\QuotaRecon.Input.ps1"

# Read the same subs the main script would use.
$scriptRoot = Split-Path -Parent $PSScriptRoot
$subs = Read-QRList -CsvPath (Join-Path $scriptRoot 'templates\subscriptions.csv') -ColumnName 'SubscriptionId'
Write-Host ("Testing {0} subs from templates/subscriptions.csv" -f $subs.Count)
Write-Host ("Ids: {0}" -f ($subs -join ', '))
Write-Host ""

$map = Get-QRSubscriptionNames -SubscriptionIds $subs

foreach ($id in $subs) {
    Write-Host ("  {0} -> {1}" -f $id, $map[$id])
}
