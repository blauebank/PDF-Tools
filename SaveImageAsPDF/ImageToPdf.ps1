# ImageToPdf.ps1
# Erstellt aus einer Bilddatei eine PDF im gleichen Ordner mit gleichem Namen.
# Verwendet Microsoft Print to PDF.
# Kein Logging.

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [switch]$Force
)

$ErrorActionPreference = "Stop"

try {

    $InputFile = (Resolve-Path -LiteralPath $InputFile).Path
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".pdf")

    if ((Test-Path $OutputFile) -and (-not $Force)) {
        exit 0
    }

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

    $printerName = "Microsoft Print to PDF"

    $installedPrinters = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
    if ($installedPrinters -notcontains $printerName) {
        throw "Microsoft Print to PDF nicht gefunden."
    }

    $img = [System.Drawing.Image]::FromFile($InputFile)

    $printDoc = New-Object System.Drawing.Printing.PrintDocument
    $printDoc.PrinterSettings.PrinterName = $printerName
    $printDoc.PrinterSettings.PrintToFile = $true
    $printDoc.PrinterSettings.PrintFileName = $OutputFile
    $printDoc.PrintController = New-Object System.Drawing.Printing.StandardPrintController
    $printDoc.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(0,0,0,0)
    $printDoc.DefaultPageSettings.Landscape = ($img.Width -gt $img.Height)

    $printDoc.add_PrintPage({
        param($sender, $e)

        $rect = $e.MarginBounds

        $imgRatio = $img.Width / $img.Height
        $rectRatio = $rect.Width / $rect.Height

        if ($imgRatio -gt $rectRatio) {
            $newHeight = [int]($rect.Width / $imgRatio)
            $y = $rect.Top + (($rect.Height - $newHeight) / 2)
            $rect = New-Object System.Drawing.Rectangle($rect.Left, $y, $rect.Width, $newHeight)
        } else {
            $newWidth = [int]($rect.Height * $imgRatio)
            $x = $rect.Left + (($rect.Width - $newWidth) / 2)
            $rect = New-Object System.Drawing.Rectangle($x, $rect.Top, $newWidth, $rect.Height)
        }

        $e.Graphics.DrawImage($img, $rect)
        $e.HasMorePages = $false
    })

    $printDoc.Print()

    $printDoc.Dispose()
    $img.Dispose()

}
catch {
    # stiller Fehler – kein Popup
}