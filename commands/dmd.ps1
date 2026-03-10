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

# Resolve the original input so we know its directory
$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path
$inputDir = Split-Path -Parent $resolvedInput
$inputBase = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)

# Step 1: docx -> md (into a temp file)
$tempMd = Join-Path ([System.IO.Path]::GetTempPath()) ("dmd_" + [System.Guid]::NewGuid().ToString("N") + ".md")

# Determine final output path (same directory as original, .docx extension)
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
  $finalOutput = Join-Path $inputDir ($inputBase + ".docx")
} else {
  $finalOutput = $OutputFile
}

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
