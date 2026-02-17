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

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

if (-not (Test-Path -LiteralPath $referenceDoc)) {
  throw "Missing reference doc: $referenceDoc"
}

$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
  $inputDir = Split-Path -Parent $resolvedInput
  $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)
  $OutputFile = Join-Path $inputDir ($inputBase + ".docx")
}

& pandoc `
  -f "markdown+fancy_lists+lists_without_preceding_blankline" `
  -t "docx" `
  --no-highlight `
  --reference-doc="$referenceDoc" `
  --lua-filter="$filterPath" `
  "$resolvedInput" `
  -o "$OutputFile"

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

Write-Output "Created: $OutputFile"
