param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf
)

$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $document = $word.Documents.Open($InputDocx, $false, $false)
    foreach ($toc in $document.TablesOfContents) {
        $toc.Update()
    }
    foreach ($field in $document.Fields) {
        $field.Update() | Out-Null
    }
    foreach ($section in $document.Sections) {
        foreach ($footer in $section.Footers) {
            foreach ($field in $footer.Range.Fields) {
                $field.Update() | Out-Null
            }
        }
    }
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.Save()
    $document.ExportAsFixedFormat($OutputPdf, 17)
    Write-Output "pages=$pages"
}
finally {
    if ($document) {
        $document.Close($false)
    }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
