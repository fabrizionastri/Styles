[CmdletBinding()]
param()

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$commandsPath = Join-Path $scriptRoot "commands.ps1"

if (-not (Test-Path -LiteralPath $commandsPath)) {
  throw "Missing commands file: $commandsPath"
}

$profilePath = $PROFILE
$profileDir = Split-Path -Parent $profilePath

if (-not (Test-Path -LiteralPath $profileDir)) {
  New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $profilePath)) {
  New-Item -ItemType File -Path $profilePath -Force | Out-Null
}

$line = ". `"$commandsPath`""
$content = Get-Content -LiteralPath $profilePath -ErrorAction Stop

if ($content -contains $line) {
  Write-Output "Profile already configured: $profilePath"
  exit 0
}

Add-Content -LiteralPath $profilePath -Value ""
Add-Content -LiteralPath $profilePath -Value $line
Write-Output "Added d2m/m2d/ld2m loader to profile: $profilePath"
