[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# This installer lives in:
#   <FlexUp - Documents>\Legal\Styles\commands\install_commands.ps1
# We walk back up the tree so we can infer the shared FlexUp documents root
# without hard-coding the current Windows username.
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$stylesRoot = Split-Path -Parent $scriptRoot
$legalRoot = Split-Path -Parent $stylesRoot
$documentsRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $legalRoot)).Path
$commandsPath = Join-Path $scriptRoot "commands.ps1"
$linkPath = "C:\FlexUp"
$envVarName = "FLEXUP"

if (-not (Test-Path -LiteralPath $commandsPath -PathType Leaf)) {
  throw "Missing commands file: $commandsPath"
}

function Convert-ToSingleQuotedPowerShellString {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  return "'" + $Value.Replace("'", "''") + "'"
}

function Ensure-DirectoryJunction {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Target
  )

  # Windows cannot create a true hard link to a directory, so for folders
  # the closest equivalent is a junction.
  if (Test-Path -LiteralPath $Path) {
    $existingItem = Get-Item -LiteralPath $Path -Force
    $existingTarget = $null

    if ($existingItem.PSObject.Properties.Name -contains "Target") {
      $rawTarget = $existingItem.Target
      if ($rawTarget -is [System.Array]) {
        $existingTarget = [string]$rawTarget[0]
      } elseif ($null -ne $rawTarget) {
        $existingTarget = [string]$rawTarget
      }
    }

    if (-not [string]::IsNullOrWhiteSpace($existingTarget)) {
      try {
        $resolvedExistingTarget = (Resolve-Path -LiteralPath $existingTarget).Path
        if ($resolvedExistingTarget -eq $Target) {
          Write-Output "Link already configured: $Path -> $Target"
          return
        }
      }
      catch {
        # If the target cannot be resolved, we fall through to the safety check below.
      }
    }

    throw "Path already exists and does not point to $Target : $Path"
  }

  try {
    # Create the folder shortcut using the exact mklink /J form:
    #   cmd /c mklink /J C:\FlexUp "C:\Users\<username>\FlexUp\FlexUp - Documents"
    $mklinkOutput = & cmd.exe /c mklink /J $Path $Target 2>&1
    if ($LASTEXITCODE -ne 0) {
      throw (($mklinkOutput | ForEach-Object { $_.ToString().Trim() }) -join [Environment]::NewLine)
    }

    Write-Output (($mklinkOutput | ForEach-Object { $_.ToString().Trim() }) -join [Environment]::NewLine)
  }
  catch {
    throw "Could not create $Path. Try running PowerShell as Administrator. $($_.Exception.Message)"
  }
}

function Ensure-ProfileBlock {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ProfilePath,

    [Parameter(Mandatory = $true)]
    [string]$CommandsPath,

    [Parameter(Mandatory = $true)]
    [string]$FlexUpPath,

    [Parameter(Mandatory = $true)]
    [string]$EnvironmentVariableName
  )

  $profileDir = Split-Path -Parent $ProfilePath
  if (-not (Test-Path -LiteralPath $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
  }

  if (-not (Test-Path -LiteralPath $ProfilePath)) {
    New-Item -ItemType File -Path $ProfilePath -Force | Out-Null
  }

  # We write a clearly delimited block so the installer can be re-run safely:
  # it replaces the FlexUp block instead of appending duplicate lines.
  $startMarker = "# >>> FlexUp commands >>>"
  $endMarker = "# <<< FlexUp commands <<<"
  $quotedCommandsPath = Convert-ToSingleQuotedPowerShellString -Value $CommandsPath
  $quotedFlexUpPath = Convert-ToSingleQuotedPowerShellString -Value $FlexUpPath

  $block = @(
    $startMarker
    "# Keep the path available as a normal environment variable too."
    'if ([string]::IsNullOrWhiteSpace($env:' + $EnvironmentVariableName + ')) {'
    '  $env:' + $EnvironmentVariableName + ' = ' + $quotedFlexUpPath
    '}'
    ""
    "# PowerShell shortcut to the shared FlexUp documents root."
    '$Global:FlexUp = $env:' + $EnvironmentVariableName
    ""
    "# Load the conversion commands (d2m, m2d, ld2d, etc.) in every session."
    ". $quotedCommandsPath"
    $endMarker
  ) -join [Environment]::NewLine

  $content = Get-Content -LiteralPath $ProfilePath -Raw
  $escapedStart = [regex]::Escape($startMarker)
  $escapedEnd = [regex]::Escape($endMarker)
  $pattern = "(?s)$escapedStart.*?$escapedEnd"

  if ($content -match $pattern) {
    $updatedContent = [regex]::Replace($content, $pattern, $block)
  }
  elseif ([string]::IsNullOrWhiteSpace($content)) {
    $updatedContent = $block + [Environment]::NewLine
  }
  else {
    $separator = if ($content.EndsWith("`r`n") -or $content.EndsWith("`n")) {
      ""
    } else {
      [Environment]::NewLine
    }
    $updatedContent = $content + $separator + [Environment]::NewLine + $block + [Environment]::NewLine
  }

  [System.IO.File]::WriteAllText(
    $ProfilePath,
    $updatedContent,
    (New-Object System.Text.UTF8Encoding($false))
  )
}

function Ensure-UserEnvironmentVariable {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  # Persist the FlexUp path at the Windows user level so it survives reboot
  # and is available to future PowerShell sessions and other processes.
  [System.Environment]::SetEnvironmentVariable($Name, $Value, "User")
}

# Step 1: create the stable C:\FlexUp shortcut to the shared documents folder.
Ensure-DirectoryJunction -Path $linkPath -Target $documentsRoot

# Step 2: store the path as a persistent Windows user environment variable.
Ensure-UserEnvironmentVariable -Name $envVarName -Value $linkPath

# Step 3: update the user's PowerShell profile so future shells always know:
# - where FlexUp lives ($FlexUp)
# - how to load the helper commands
$profilePath = $PROFILE
Ensure-ProfileBlock `
  -ProfilePath $profilePath `
  -CommandsPath $commandsPath `
  -FlexUpPath $linkPath `
  -EnvironmentVariableName $envVarName

# Step 4: make the current PowerShell session usable immediately as well.
Set-Item -Path ("Env:" + $envVarName) -Value $linkPath
$Global:FlexUp = $linkPath
. $commandsPath

Write-Output "Profile configured: $profilePath"
Write-Output "User environment variable configured: $envVarName=$linkPath"
Write-Output '$FlexUp now points to: C:\FlexUp'
Write-Output "Close and reopen PowerShell so every new session picks up the updated profile."
