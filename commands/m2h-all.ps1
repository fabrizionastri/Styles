[CmdletBinding()]
param(
  [Parameter(Position = 0)]
  [string]$TargetFolder
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Source the main commands file so m2h is available
. (Join-Path $scriptRoot "commands.ps1")

# Default to current directory if no folder specified
if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
  $TargetFolder = Get-Location
}

# Resolve to absolute path
$TargetFolder = (Resolve-Path -LiteralPath $TargetFolder -ErrorAction Stop).Path

if (-not (Test-Path -LiteralPath $TargetFolder -PathType Container)) {
  Write-Error "Target folder does not exist: $TargetFolder"
  return
}

# Find all .md files in the top-level folder only
$mdFiles = @(Get-ChildItem -LiteralPath $TargetFolder -Filter "*.md" -File)

if ($mdFiles.Count -eq 0) {
  Write-Host "No md file found in this directory"
  return
}

function Get-UniqueOutputPath {
  param(
    [string]$BasePath,
    [string]$Extension
  )

  $dir  = Split-Path -Parent $BasePath
  $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
  $candidate = Join-Path $dir "$name$Extension"

  if (-not (Test-Path -LiteralPath $candidate)) {
    return $candidate
  }

  if ($name -match '^(.+)_(\d+)$') {
    $root = $Matches[1]
    $counter = [int]$Matches[2] + 1
  } else {
    $root = $name
    $counter = 1
  }

  do {
    $candidate = Join-Path $dir "${root}_${counter}${Extension}"
    $counter++
  } while (Test-Path -LiteralPath $candidate)

  return $candidate
}

$total   = $mdFiles.Count
$current = 0
$errors  = 0

foreach ($file in $mdFiles) {
  $current++
  $outputPath = Get-UniqueOutputPath -BasePath $file.FullName -Extension ".html"
  $outputName = Split-Path -Leaf $outputPath

  Write-Host "[$current/$total] Converting $($file.Name) -> $outputName"

  try {
    m2h $file.FullName $outputPath
  }
  catch {
    Write-Warning "  Failed: $_"
    $errors++
  }
}

Write-Host ""
if ($errors -eq 0) {
  Write-Host "Done: $total file(s) converted."
} else {
  Write-Host "Done: $($total - $errors) of $total file(s) converted ($errors error(s))."
}
