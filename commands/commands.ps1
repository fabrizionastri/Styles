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

function ld2d {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "ld2d.ps1") $InputFile $OutputFile
}

function m2d-all {
  [CmdletBinding()]
  param(
    [Parameter(Position = 0)]
    [string]$TargetFolder
  )

  & (Join-Path $script:commandsRoot "m2d-all.ps1") $TargetFolder
}

function dmd {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "dmd.ps1") $InputFile $OutputFile
}

function mdm {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Position = 1)]
    [string]$OutputFile
  )

  & (Join-Path $script:commandsRoot "mdm.ps1") $InputFile $OutputFile
}

function d2m-all {
  [CmdletBinding()]
  param(
    [Parameter(Position = 0)]
    [string]$TargetFolder
  )

  & (Join-Path $script:commandsRoot "d2m-all.ps1") $TargetFolder
}
