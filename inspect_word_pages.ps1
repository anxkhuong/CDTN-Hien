param([Parameter(Mandatory = $true)][string[]]$InputDocx)

$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    foreach ($path in $InputDocx) {
        $document = $word.Documents.Open($path, $false, $true)
        $document.Repaginate()
        $pages = $document.ComputeStatistics(2)
        $words = $document.ComputeStatistics(0)
        Write-Output "$path`tpages=$pages`twords=$words"
        $document.Close($false)
        $document = $null
    }
}
finally {
    if ($document) { $document.Close($false) }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
