param([Parameter(Mandatory=$true)][string]$Deck, [Parameter(Mandatory=$true)][string]$Output)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
    $presentation = $app.Presentations.Open((Resolve-Path -LiteralPath $Deck).Path, -1, 0, 0)
    Write-Output "PowerPoint version: $($app.Version)"
    $presentation.Slides.Item(1).Export([IO.Path]::GetFullPath($Output), 'PNG', 1600, 900)
    Write-Output "Exported: $Output"
} finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
