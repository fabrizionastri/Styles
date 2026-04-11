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

function Test-IsDirectoryOutputPath {
  param([string]$PathValue)

  if ([string]::IsNullOrWhiteSpace($PathValue)) {
    return $false
  }

  $trimmed = $PathValue.Trim()

  if ($trimmed -eq "." -or $trimmed -eq "..") {
    return $true
  }

  if (Test-Path -LiteralPath $trimmed -PathType Container) {
    return $true
  }

  return $trimmed.EndsWith([System.IO.Path]::DirectorySeparatorChar) -or
    $trimmed.EndsWith([System.IO.Path]::AltDirectorySeparatorChar)
}

function Resolve-FinalOutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  $inputDir = Split-Path -Parent $InputPath
  $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    return [System.IO.Path]::GetFullPath((Join-Path $inputDir ($inputBase + $DefaultExtension)))
  }

  $workingDirectory = (Get-Location).ProviderPath
  $candidate = $OutputPath.Trim()

  if (Test-IsDirectoryOutputPath -PathValue $candidate) {
    $outputDirectory = if ([System.IO.Path]::IsPathRooted($candidate)) {
      $candidate
    } else {
      Join-Path $workingDirectory $candidate
    }

    if (Test-Path -LiteralPath $outputDirectory -PathType Container) {
      $outputDirectory = (Resolve-Path -LiteralPath $outputDirectory -ErrorAction Stop).Path
    } else {
      $outputDirectory = [System.IO.Path]::GetFullPath($outputDirectory)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $outputDirectory ($inputBase + $DefaultExtension)))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($candidate))) {
    $candidate = $candidate + $DefaultExtension
  }

  if (-not [System.IO.Path]::IsPathRooted($candidate)) {
    $candidate = Join-Path $workingDirectory $candidate
  }

  return [System.IO.Path]::GetFullPath($candidate)
}

function Get-TrailingLineBreakSuffix {
  param([string]$Text)

  if ($null -eq $Text) {
    return ""
  }

  $i = $Text.Length
  while ($i -gt 0) {
    $ch = $Text[$i - 1]
    if ($ch -eq "`n" -or $ch -eq "`r") {
      $i--
      continue
    }
    break
  }

  return $Text.Substring($i)
}

function Sync-TerminalLineBreaks {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
  )

  if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    return
  }
  if (-not (Test-Path -LiteralPath $TargetPath -PathType Leaf)) {
    return
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $sourceText = [System.IO.File]::ReadAllText($SourcePath, $utf8NoBom)
  $targetText = [System.IO.File]::ReadAllText($TargetPath, $utf8NoBom)

  $sourceUsesCrLf = $sourceText.Contains("`r`n")
  $targetNormalized = $targetText -replace "`r`n", "`n"
  $targetNormalized = $targetNormalized -replace "`r", "`n"

  if ($sourceUsesCrLf) {
    $targetText = $targetNormalized -replace "`n", "`r`n"
  } else {
    $targetText = $targetNormalized
  }

  $sourceSuffix = Get-TrailingLineBreakSuffix -Text $sourceText
  $targetCore = $targetText.TrimEnd("`r", "`n")
  $normalized = $targetCore + $sourceSuffix

  [System.IO.File]::WriteAllText($TargetPath, $normalized, $utf8NoBom)
}

# Resolve the original input so we know its directory
$resolvedInput = (Resolve-Path -LiteralPath $InputFile -ErrorAction Stop).Path

# Step 1: md -> docx (into a temp file)
$tempDocx = Join-Path ([System.IO.Path]::GetTempPath()) ("mdm_" + [System.Guid]::NewGuid().ToString("N") + ".docx")

# Determine final output path from the original input, not the temp docx file.
$finalOutput = Resolve-FinalOutputPath `
  -InputPath $resolvedInput `
  -OutputPath $OutputFile `
  -DefaultExtension ".md"

try {
  & $m2dScript $InputFile $tempDocx
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

  # Step 2: docx -> md (output next to the original file)
  $d2mLines = @(& $d2mScript $tempDocx $finalOutput)
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
  $d2mLines | ForEach-Object { Write-Output $_ }

  $createdOutput = $null
  foreach ($line in $d2mLines) {
    if ($line -match '^Created:\s+(.+)$') {
      $createdOutput = $Matches[1].Trim()
    }
  }

  if ([string]::IsNullOrWhiteSpace($createdOutput)) {
    $createdOutput = $finalOutput
  }

  if (-not [System.IO.Path]::IsPathRooted($createdOutput)) {
    $createdOutput = Join-Path (Get-Location).ProviderPath $createdOutput
  }

  Sync-TerminalLineBreaks -SourcePath $resolvedInput -TargetPath $createdOutput
}
finally {
  if (Test-Path -LiteralPath $tempDocx) {
    Remove-Item -LiteralPath $tempDocx -Force
  }
}
