Set-StrictMode -Version Latest
$script:commandsRoot = $PSScriptRoot

function d2m {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "d2m.ps1") $InputFile $OutputFile
}

function m2d {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "m2d.ps1") $InputFile $OutputFile
}

function ld2m {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "legacy_d2m.ps1") $InputFile $OutputFile
}
