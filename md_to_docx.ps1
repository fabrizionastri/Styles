[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $scriptRoot "m2d.ps1") $InputFile $OutputFile

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}
