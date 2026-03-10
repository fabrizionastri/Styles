[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$m2dScript = Join-Path $scriptRoot "m2d.ps1"
$d2mScript = Join-Path $scriptRoot "d2m.ps1"

# Resolve the original input so we know its directory
$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path
$inputDir = Split-Path -Parent $resolvedInput
$inputBase = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)

# Step 1: md -> docx (into a temp file)
$tempDocx = Join-Path ([System.IO.Path]::GetTempPath()) ("mdm_" + [System.Guid]::NewGuid().ToString("N") + ".docx")

# Determine final output path (same directory as original, .md extension)
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
  $finalOutput = Join-Path $inputDir ($inputBase + ".md")
} else {
  $finalOutput = $OutputFile
}

try {
  & $m2dScript $InputFile $tempDocx
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

  # Step 2: docx -> md (output next to the original file)
  & $d2mScript $tempDocx $finalOutput
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
  if (Test-Path -LiteralPath $tempDocx) {
    Remove-Item -LiteralPath $tempDocx -Force
  }
}
