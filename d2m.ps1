[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "filters\docx_to_compact.lua"

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
  $inputDir = Split-Path -Parent $resolvedInput
  $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)
  $OutputFile = Join-Path $inputDir ($inputBase + ".md")
}

& pandoc `
  -f "docx+styles" `
  -t "markdown" `
  --lua-filter="$filterPath" `
  "$resolvedInput" `
  -o "$OutputFile"

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

Write-Output "Created: $OutputFile"
