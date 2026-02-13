param(
    [string]$SourceWorkbook = ".\AI_Assistant_latest_dev.xlsm",
    [string]$BaseAddin = ".\AI_Assistant.xlam",
    [string]$OutputAddin = ".\AI_Assistant.xlam",
    [switch]$KeepTemp
)

$ErrorActionPreference = "Stop"

function Resolve-FullPathOrThrow {
    param([string]$PathValue, [string]$Label)
    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "$Label not found: $PathValue"
    }
    (Resolve-Path -LiteralPath $PathValue).Path
}

$sourcePath = Resolve-FullPathOrThrow -PathValue $SourceWorkbook -Label "Source workbook"
$basePath = Resolve-FullPathOrThrow -PathValue $BaseAddin -Label "Base add-in"

$sevenZip = Get-Command 7z -ErrorAction SilentlyContinue
if (-not $sevenZip) {
    throw "7z was not found in PATH. Install 7-Zip CLI and try again."
}

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ai_assistant_sync_" + [guid]::NewGuid().ToString("N"))
$baseDir = Join-Path $tmpRoot "base"
$srcDir = Join-Path $tmpRoot "src"

New-Item -ItemType Directory -Path $baseDir -Force | Out-Null
New-Item -ItemType Directory -Path $srcDir -Force | Out-Null

try {
    & 7z x $basePath ("-o" + $baseDir) -y | Out-Null
    & 7z x $sourcePath ("-o" + $srcDir) -y | Out-Null

    $srcVba = Join-Path $srcDir "xl\vbaProject.bin"
    $dstVba = Join-Path $baseDir "xl\vbaProject.bin"

    if (-not (Test-Path -LiteralPath $srcVba)) {
        throw "Source workbook does not contain xl/vbaProject.bin: $sourcePath"
    }
    if (-not (Test-Path -LiteralPath $dstVba)) {
        throw "Base add-in does not contain xl/vbaProject.bin: $basePath"
    }

    Copy-Item -LiteralPath $srcVba -Destination $dstVba -Force

    $zipOut = Join-Path $tmpRoot "addin.zip"
    Push-Location $baseDir
    try {
        & 7z a -tzip $zipOut .\* -mx=9 | Out-Null
    }
    finally {
        Pop-Location
    }

    $outputFull = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputAddin))
    Copy-Item -LiteralPath $zipOut -Destination $outputFull -Force
    Write-Output "Updated add-in: $outputFull"
}
finally {
    if (-not $KeepTemp) {
        Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Output "Temp kept: $tmpRoot"
    }
}
