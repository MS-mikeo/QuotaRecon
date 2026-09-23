$ErrorActionPreference = 'Stop'
$files = Get-ChildItem -Recurse -Filter *.ps1 -Path 'c:\GitHub\PowerShell\Capacity\QuotaRecon'
$hadError = $false
foreach ($f in $files) {
    $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errs)
    if ($errs -and $errs.Count) {
        $hadError = $true
        Write-Host $f.FullName -ForegroundColor Red
        foreach ($e in $errs) {
            Write-Host ("  line {0}: {1}" -f $e.Extent.StartLineNumber, $e.Message)
        }
    } else {
        Write-Host ("OK  {0}" -f $f.Name) -ForegroundColor Green
    }
}
if ($hadError) { exit 1 } else { exit 0 }
