Set-StrictMode -Version Latest

Add-Type -AssemblyName Office
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

Write-Output "ppt2pdf v0.1 by Shinichi Akiyama"

if ($args.Length -eq 0) {
    Write-Output "[Usage] ppt2pdf [FILE]..."
    exit 1
}

$ppt = New-Object -ComObject PowerPoint.Application
Write-Output "PowoerPoint version $($ppt.version)"

Get-ChildItem -Path $args | ForEach-Object {
    $presentation = $ppt.Presentations.Open($_)
    Write-Output "$_ was opened."

    $pdf = Join-Path (Split-Path $_) ([System.IO.Path]::GetFileNameWithoutExtension($_) + ".pdf")

    $presentation.PrintOptions.Ranges.ClearAll()
    [void] $presentation.PrintOptions.Ranges.Add(1, $presentation.Slides.Count)

    $presentation.ExportAsFixedFormat($pdf,
        [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF,
        [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        [Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder]::ppPrintHandoutVerticalFirst,
        [Microsoft.Office.Interop.PowerPoint.PpPrintOutputType]::ppPrintOutputSlides,
        [Microsoft.Office.Core.MsoTriState]::msoFalse,
        $presentation.PrintOptions.Ranges.Item(1),
        [Microsoft.Office.Interop.PowerPoint.PpPrintRangeType]::ppPrintSlideRange)

    $presentation.Close()
    Write-Output "$pdf was created."
}
Write-Output "All files were converted."
