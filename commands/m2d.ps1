[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "filters\compact_to_docx.lua"
$referenceDoc = Join-Path $scriptRoot "styles.docx"
$defaultInputExtension = ".md"
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

function Resolve-OutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputDir = Split-Path -Parent $InputPath
    $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    return (Join-Path $inputDir ($inputBase + $DefaultExtension))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($OutputPath))) {
    return ($OutputPath + $DefaultExtension)
  }

  return $OutputPath
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

& pandoc `
  -f "markdown+fancy_lists+lists_without_preceding_blankline" `
  -t "docx" `
  --no-highlight `
  --reference-doc="$referenceDoc" `
  --lua-filter="$filterPath" `
  "$resolvedInput" `
  -o "$resolvedOutput"

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

Write-Output "Created: $resolvedOutput"
