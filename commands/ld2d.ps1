[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "..\filters\remap.lua"
$referenceDoc = Join-Path $scriptRoot "..\styles\flexup_template.docx"
$defaultInputExtension = ".docx"
$defaultOutputExtension = ".docx"

function Resolve-InputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  if (Test-Path -LiteralPath $PathValue) {
    return (Resolve-Path -LiteralPath $PathValue -ErrorAction Stop).Path
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($PathValue))) {
    $withExt = $PathValue + $DefaultExtension
    if (Test-Path -LiteralPath $withExt) {
      return (Resolve-Path -LiteralPath $withExt -ErrorAction Stop).Path
    }
  }

  throw "Input file not found: $PathValue"
}

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

function Resolve-ExplicitOutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$InputBase,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

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

    return (Join-Path $outputDirectory ($InputBase + $DefaultExtension))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($candidate))) {
    $candidate = $candidate + $DefaultExtension
  }

  if (-not [System.IO.Path]::IsPathRooted($candidate)) {
    $candidate = Join-Path $workingDirectory $candidate
  }

  return [System.IO.Path]::GetFullPath($candidate)
}

function Resolve-OutputPath {
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
    $resolved = Join-Path $inputDir ($inputBase + "_remapped" + $DefaultExtension)
  } else {
    $resolved = Resolve-ExplicitOutputPath `
      -OutputPath $OutputPath `
      -InputBase ($inputBase + "_remapped") `
      -DefaultExtension $DefaultExtension
  }

  $resolved = [System.IO.Path]::GetFullPath($resolved)

  if (Test-Path -LiteralPath $resolved) {
    $dir = Split-Path -Parent $resolved
    $ext = [System.IO.Path]::GetExtension($resolved)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolved)

    if ($baseName -match '^(.+)_(\d+)$') {
      $root = $Matches[1]
      $counter = [int]$Matches[2] + 1
    } else {
      $root = $baseName
      $counter = 1
    }

    do {
      $resolved = Join-Path $dir "${root}_${counter}${ext}"
      $counter++
    } while (Test-Path -LiteralPath $resolved)
  }

  return $resolved
}

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

if (-not (Test-Path -LiteralPath $referenceDoc)) {
  throw "Missing reference doc: $referenceDoc"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ld2d_" + [System.Guid]::NewGuid().ToString("N"))
$tempInput = Join-Path $tempDir "input.docx"

try {
  New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
  Copy-Item -LiteralPath $resolvedInput -Destination $tempInput -Force

  & pandoc `
    -f "docx+styles" `
    -t "docx" `
    --lua-filter="$filterPath" `
    --reference-doc="$referenceDoc" `
    --no-highlight `
    "$tempInput" `
    -o "$resolvedOutput"

  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }
}
finally {
  if (Test-Path -LiteralPath $tempDir) {
    Remove-Item -LiteralPath $tempDir -Force -Recurse
  }
}

Write-Output "Created: $resolvedOutput"
