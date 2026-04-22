param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$PrinterName,

    [int]$Copies = 1,
    [string]$PageRange,
    [string]$WorkingRoot = 'C:\Users\Public\Documents\hermes-print-jobs',
    [switch]$SetDefaultPrinterForHwp,
    [int]$PostSubmitWaitSeconds = 10
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"
    Write-Host $line
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
    }
}

function Get-PrinterStrict {
    param([string]$Name)
    $printer = Get-Printer -Name $Name -ErrorAction SilentlyContinue
    if (-not $printer) {
        throw "Printer not found: $Name"
    }
    return $printer
}

function New-WorkDir {
    param([string]$Root)
    $jobId = Get-Date -Format 'yyyyMMdd_HHmmss'
    $dir = Join-Path $Root $jobId
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

function Copy-SourceToWorkDir {
    param(
        [string]$LiteralSourcePath,
        [string]$WorkDir
    )
    if (-not (Test-Path -LiteralPath $LiteralSourcePath)) {
        throw "Source file not found: $LiteralSourcePath"
    }
    $leaf = Split-Path -Leaf $LiteralSourcePath
    $dest = Join-Path $WorkDir $leaf
    Copy-Item -LiteralPath $LiteralSourcePath -Destination $dest -Force
    return $dest
}

function Get-DefaultPrinterName {
    $defaultPrinter = Get-CimInstance Win32_Printer | Where-Object { $_.Default } | Select-Object -First 1
    if ($defaultPrinter) { return $defaultPrinter.Name }
    return $null
}

function Set-DefaultPrinterName {
    param([string]$Name)
    $network = New-Object -ComObject WScript.Network
    $network.SetDefaultPrinter($Name)
}

function Wait-ForQueueObservation {
    param(
        [string]$Name,
        [int]$Seconds
    )
    Start-Sleep -Seconds $Seconds
    return @(Get-PrintJob -PrinterName $Name -ErrorAction SilentlyContinue)
}

function Get-RecentAcrobatCrash {
    param([datetime]$Since)
    return Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $Since } -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.ProviderName -eq 'Application Error' -or $_.ProviderName -eq 'Windows Error Reporting') -and
            $_.Message -match 'Acrobat.exe'
        } |
        Select-Object -First 1
}

function Submit-PdfPrintAsImages {
    param(
        [string]$PrintablePath,
        [object]$Printer,
        [int]$Copies,
        [string]$WorkDir,
        [int]$WaitSeconds
    )

    Add-Type -AssemblyName System.Drawing
    $pdfToCairo = 'C:\poppler-23.01.0\Library\bin\pdftocairo.exe'
    if (-not (Test-Path $pdfToCairo)) {
        throw "pdftocairo not found: $pdfToCairo"
    }

    $renderDir = Join-Path $WorkDir 'rendered-pages'
    New-Item -ItemType Directory -Path $renderDir -Force | Out-Null
    $renderBase = Join-Path $renderDir 'page'

    Write-Log 'Rendering PDF pages to PNG for fallback printing'
    & $pdfToCairo -png -r 150 -- $PrintablePath $renderBase
    if ($LASTEXITCODE -ne 0) {
        throw "pdftocairo image render failed with exit code $LASTEXITCODE"
    }

    $pages = @(Get-ChildItem -LiteralPath $renderDir -Filter 'page-*.png' | Sort-Object Name)
    if ($pages.Count -eq 0) {
        throw 'No rendered PNG pages found for PDF fallback printing.'
    }

    for ($copy = 1; $copy -le $Copies; $copy++) {
        Write-Log "Submitting PDF print job $copy/$Copies via image-render fallback"
        $pageIndex = 0
        $printDoc = New-Object System.Drawing.Printing.PrintDocument
        $printDoc.PrinterSettings.PrinterName = $Printer.Name
        $printDoc.DocumentName = [System.IO.Path]::GetFileName($PrintablePath)

        $handler = [System.Drawing.Printing.PrintPageEventHandler]{
            param($sender, $e)
            $imgPath = $pages[$pageIndex].FullName
            $img = [System.Drawing.Image]::FromFile($imgPath)
            try {
                $margin = $e.MarginBounds
                $ratioX = $margin.Width / $img.Width
                $ratioY = $margin.Height / $img.Height
                $ratio = [Math]::Min($ratioX, $ratioY)
                $width = [int]($img.Width * $ratio)
                $height = [int]($img.Height * $ratio)
                $x = $margin.X + [int](($margin.Width - $width) / 2)
                $y = $margin.Y + [int](($margin.Height - $height) / 2)
                $e.Graphics.DrawImage($img, $x, $y, $width, $height)
            }
            finally {
                $img.Dispose()
            }

            $pageIndex++
            $e.HasMorePages = $pageIndex -lt $pages.Count
        }

        $printDoc.add_PrintPage($handler)
        $printDoc.Print()
        $jobs = Wait-ForQueueObservation -Name $Printer.Name -Seconds $WaitSeconds
        if ($jobs.Count -gt 0) {
            Write-Log "Observed print queue entries after fallback submission: $($jobs.Count)"
        } else {
            Write-Log 'No queue entry observed after fallback submission. Physical confirmation still required.'
        }
    }
}

function Submit-PdfPrint {
    param(
        [string]$PrintablePath,
        [object]$Printer,
        [int]$Copies,
        [int]$WaitSeconds,
        [string]$WorkDir
    )

    $acroCandidates = @(
        'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
        'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe'
    )
    $acro = $acroCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($acro) {
        $started = Get-Date
        for ($i = 1; $i -le $Copies; $i++) {
            Write-Log "Submitting PDF print job $i/$Copies via Acrobat"
            Start-Process -FilePath $acro -ArgumentList @('/t', $PrintablePath, $Printer.Name, $Printer.DriverName, $Printer.PortName) | Out-Null
            $jobs = Wait-ForQueueObservation -Name $Printer.Name -Seconds $WaitSeconds
            if ($jobs.Count -gt 0) {
                Write-Log "Observed print queue entries after PDF submission: $($jobs.Count)"
                return
            }
        }

        $crash = Get-RecentAcrobatCrash -Since $started
        if ($crash) {
            Write-Log 'Detected Acrobat crash during PDF print submission.'
            throw 'Acrobat crashed during PDF printing. Automatic fallback printing is temporarily disabled until the fallback path is verified safe.'
        }

        Write-Log 'Acrobat submission gave no observable queue activity.'
        throw 'PDF print submission produced no observable queue activity. Automatic fallback printing is temporarily disabled until verified safe.'
    }

    Write-Log 'Adobe Acrobat/Reader not found.'
    throw 'Adobe Acrobat/Reader not found, and automatic fallback printing is temporarily disabled until verified safe.'
}

function Submit-HwpPrint {
    param(
        [string]$PrintablePath,
        [object]$Printer,
        [int]$Copies,
        [switch]$AllowSetDefault,
        [int]$WaitSeconds
    )

    $hwpPrintManager = 'C:\Program Files (x86)\Hnc\Office 2018\HOffice100\Bin\HwpPrnMng.exe'
    if (-not (Test-Path $hwpPrintManager)) {
        throw "HWP print manager not found: $hwpPrintManager"
    }

    $originalDefault = Get-DefaultPrinterName
    $changedDefault = $false

    try {
        if ($originalDefault -ne $Printer.Name) {
            if ($AllowSetDefault) {
                Write-Log "Changing default printer temporarily for HWP printing: $originalDefault -> $($Printer.Name)"
                Set-DefaultPrinterName -Name $Printer.Name
                $changedDefault = $true
                Start-Sleep -Seconds 2
            } else {
                throw "HWP printing uses the Windows default printer in this environment. Current default is '$originalDefault', target is '$($Printer.Name)'. Re-run with -SetDefaultPrinterForHwp if you want the script to switch the default printer temporarily."
            }
        }

        for ($i = 1; $i -le $Copies; $i++) {
            Write-Log "Submitting HWP print job $i/$Copies via HwpPrnMng.exe"
            Start-Process -FilePath $hwpPrintManager -ArgumentList @('/p', $PrintablePath) | Out-Null
            $jobs = Wait-ForQueueObservation -Name $Printer.Name -Seconds $WaitSeconds
            if ($jobs.Count -gt 0) {
                Write-Log "Observed print queue entries after HWP submission: $($jobs.Count)"
            } else {
                Write-Log 'No queue entry observed after HWP submission. HWP can fail silently, so physical confirmation may still be required.'
            }
        }
    }
    finally {
        if ($changedDefault -and $originalDefault) {
            Write-Log "Restoring original default printer: $originalDefault"
            Set-DefaultPrinterName -Name $originalDefault
        }
    }
}

$sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
$extension = $sourceItem.Extension.ToLowerInvariant()
$printer = Get-PrinterStrict -Name $PrinterName

$workDir = New-WorkDir -Root $WorkingRoot
$script:LogFile = Join-Path $workDir 'print-log.txt'
Write-Log "Work directory created: $workDir"
Write-Log "Source: $SourcePath"
Write-Log "Printer: $($printer.Name)"
Write-Log "Copies: $Copies"
Write-Log "Extension: $extension"

$workingCopy = $null
$printablePath = $SourcePath

switch ($extension) {
    '.pdf' {
        if ($PageRange) {
            $workingCopy = Copy-SourceToWorkDir -LiteralSourcePath $SourcePath -WorkDir $workDir
            Write-Log "Working copy created: $workingCopy"
            $printablePath = $workingCopy

            if ($PageRange -notmatch '^(\d+)(-(\d+))?$') {
                throw "Invalid PageRange format: $PageRange. Use '1' or '1-3'."
            }

            $firstPage = [int]$Matches[1]
            $lastPage = if ($Matches[3]) { [int]$Matches[3] } else { $firstPage }
            $pdfToCairo = 'C:\poppler-23.01.0\Library\bin\pdftocairo.exe'
            if (-not (Test-Path $pdfToCairo)) {
                throw "pdftocairo not found: $pdfToCairo"
            }

            $derivedBase = Join-Path $workDir 'derived-pages'
            Write-Log "Creating derived PDF for page range $firstPage-$lastPage"
            & $pdfToCairo -pdf -f $firstPage -l $lastPage -- $workingCopy $derivedBase
            if ($LASTEXITCODE -ne 0) {
                throw "pdftocairo failed with exit code $LASTEXITCODE"
            }

            $derivedRaw = $derivedBase
            $derivedPdf = "$derivedBase.pdf"
            if (Test-Path -LiteralPath $derivedRaw) {
                Copy-Item -LiteralPath $derivedRaw -Destination $derivedPdf -Force
            }
            if (-not (Test-Path -LiteralPath $derivedPdf)) {
                throw "Derived PDF was not created: $derivedPdf"
            }
            $printablePath = $derivedPdf
            Write-Log "Derived printable PDF created: $printablePath"
        }

        Submit-PdfPrint -PrintablePath $printablePath -Printer $printer -Copies $Copies -WaitSeconds $PostSubmitWaitSeconds -WorkDir $workDir
    }

    '.hwp' {
        $workingCopy = Copy-SourceToWorkDir -LiteralSourcePath $SourcePath -WorkDir $workDir
        Write-Log "Working copy created: $workingCopy"
        Submit-HwpPrint -PrintablePath $workingCopy -Printer $printer -Copies $Copies -AllowSetDefault:$SetDefaultPrinterForHwp -WaitSeconds $PostSubmitWaitSeconds
    }

    default {
        throw "Unsupported file type: $extension. Supported types are .pdf and .hwp"
    }
}

Write-Log 'Print submission step completed. Physical printer confirmation is still recommended.'
Write-Host "`nWORKDIR=$workDir"
Write-Host "LOGFILE=$script:LogFile"