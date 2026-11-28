param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$OutputPdf
)

$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $inputPath = [System.IO.Path]::GetFullPath($InputDocx)
    $outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    $document.ExportAsFixedFormat($outputPath, 17)
}
finally {
    if ($document -ne $null) {
        try { $document.Close(0) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
