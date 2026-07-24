$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$pptx = (Resolve-Path (Join-Path $root 'outputs\zw1f25_MATH6185_presentation\zw1f25_MATH6185_Presentation.pptx')).Path
$pdf = Join-Path (Split-Path $pptx) 'zw1f25_MATH6185_Presentation.pdf'

$app = New-Object -ComObject PowerPoint.Application

try {
    $presentation = $app.Presentations.Open($pptx, $true, $false, $false)
    try {
        $presentation.SaveAs($pdf, 32)
    }
    finally {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
}
finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "PRESENTATION PDF EXPORT PASS: $pdf"
