[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$d2mScript = Join-Path $scriptRoot "d2m.ps1"
$m2dScript = Join-Path $scriptRoot "m2d.ps1"

function Test-IsDirectoryOutputPath {
  param([string]$PathValue)

  if ([string]::IsNullOrWhiteSpace($PathValue)) {
    return $false
  }

  $trimmed = $PathValue.Trim()

  if ($trimmed -eq "." -or $trimmed -eq "..") {
    return $true
  }

  if (Test-Path -LiteralPath $trimmed -PathType Container) {
    return $true
  }

  return $trimmed.EndsWith([System.IO.Path]::DirectorySeparatorChar) -or
    $trimmed.EndsWith([System.IO.Path]::AltDirectorySeparatorChar)
}

function Resolve-FinalOutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  $inputDir = Split-Path -Parent $InputPath
  $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    return [System.IO.Path]::GetFullPath((Join-Path $inputDir ($inputBase + $DefaultExtension)))
  }

  $workingDirectory = (Get-Location).ProviderPath
  $candidate = $OutputPath.Trim()

  if (Test-IsDirectoryOutputPath -PathValue $candidate) {
    $outputDirectory = if ([System.IO.Path]::IsPathRooted($candidate)) {
      $candidate
    } else {
      Join-Path $workingDirectory $candidate
    }

    if (Test-Path -LiteralPath $outputDirectory -PathType Container) {
      $outputDirectory = (Resolve-Path -LiteralPath $outputDirectory -ErrorAction Stop).Path
    } else {
      $outputDirectory = [System.IO.Path]::GetFullPath($outputDirectory)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $outputDirectory ($inputBase + $DefaultExtension)))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($candidate))) {
    $candidate = $candidate + $DefaultExtension
  }

  if (-not [System.IO.Path]::IsPathRooted($candidate)) {
    $candidate = Join-Path $workingDirectory $candidate
  }

  return [System.IO.Path]::GetFullPath($candidate)
}

# Resolve the original input so we know its directory
$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path

# Step 1: docx -> md (into a temp file)
$tempMd = Join-Path ([System.IO.Path]::GetTempPath()) ("dmd_" + [System.Guid]::NewGuid().ToString("N") + ".md")

# Determine final output path from the original input, not the temp markdown file.
$finalOutput = Resolve-FinalOutputPath `
  -InputPath $resolvedInput `
  -OutputPath $OutputFile `
  -DefaultExtension ".docx"

try {
  & $d2mScript $InputFile $tempMd
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

  # Step 2: md -> docx (output next to the original file)
  & $m2dScript $tempMd $finalOutput
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
  if (Test-Path -LiteralPath $tempMd) {
    Remove-Item -LiteralPath $tempMd -Force
  }
}
