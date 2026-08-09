[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "..\filters\compact_to_html.lua"
$cssPath = Join-Path $scriptRoot "..\styles\flexup_styles.css"
$ErrorActionPreference = 'Stop'
$defaultInputExtension = ".md"
$defaultOutputExtension = ".html"

function Resolve-InputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  $resolved = $null

  if (Test-Path -LiteralPath $PathValue -PathType Leaf) {
    $resolved = (Resolve-Path -LiteralPath $PathValue -ErrorAction Stop).Path
  }
  elseif ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($PathValue))) {
    $withExt = $PathValue + $DefaultExtension
    if (Test-Path -LiteralPath $withExt -PathType Leaf) {
      $resolved = (Resolve-Path -LiteralPath $withExt -ErrorAction Stop).Path
    }
  }

  if (-not $resolved) {
    # Report where a relative path actually landed. The usual cause is running
    # the command from inside the folder that the path itself already names.
    $attempted = $PathValue
    try {
      $attempted = [System.IO.Path]::GetFullPath(
        [System.IO.Path]::Combine($PWD.ProviderPath, $PathValue))
    }
    catch {
      $attempted = $PathValue
    }
    throw "Input file not found: $PathValue (looked for '$attempted')"
  }

  $actualExtension = [System.IO.Path]::GetExtension($resolved)
  if ($actualExtension -ne $DefaultExtension) {
    throw "This command reads $DefaultExtension files, but got '$actualExtension': $resolved"
  }

  return $resolved
}

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

function Resolve-ExplicitOutputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$InputBase,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

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

    return (Join-Path $outputDirectory ($InputBase + $DefaultExtension))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($candidate))) {
    $candidate = $candidate + $DefaultExtension
  }

  if (-not [System.IO.Path]::IsPathRooted($candidate)) {
    $candidate = Join-Path $workingDirectory $candidate
  }

  return [System.IO.Path]::GetFullPath($candidate)
}

function Resolve-OutputPath {
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
    $resolved = Join-Path $inputDir ($inputBase + $DefaultExtension)
  } else {
    $resolved = Resolve-ExplicitOutputPath `
      -OutputPath $OutputPath `
      -InputBase $inputBase `
      -DefaultExtension $DefaultExtension
  }

  $resolved = [System.IO.Path]::GetFullPath($resolved)

  if (Test-Path -LiteralPath $resolved) {
    $dir = Split-Path -Parent $resolved
    $ext = [System.IO.Path]::GetExtension($resolved)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolved)

    if ($baseName -match '^(.+)_(\d+)$') {
      $root = $Matches[1]
      $counter = [int]$Matches[2] + 1
    } else {
      $root = $baseName
      $counter = 1
    }

    do {
      $resolved = Join-Path $dir "${root}_${counter}${ext}"
      $counter++
    } while (Test-Path -LiteralPath $resolved)
  }

  return $resolved
}

function Convert-OffsetPrefixesToStyleSuffix {
  param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath
  )

  $raw = [System.IO.File]::ReadAllText($InputPath, [System.Text.Encoding]::UTF8)
  $normalized = $raw -replace "`r`n", "`n"
  $lines = $normalized -split "`n", -1
  $outLines = New-Object 'System.Collections.Generic.List[string]'

  $inFence = $false
  $fenceChar = ""
  # Inside a "::: Definitions" block the indentation is structural: it binds
  # each block to its definition-list item. Rewriting indented lines there
  # would dedent them out of the definition, so pass the block through as-is.
  $inDefinitions = $false

  for ($i = 0; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]

    if (-not $inFence) {
      if ($inDefinitions) {
        # Everything belonging to a definition is indented, so the only
        # unindented closing fence is the one that ends the block.
        if ($line -match '^:::+\s*$') {
          $inDefinitions = $false
        }
        $outLines.Add($line)
        continue
      }
      if ($line -match '^:::+\s*\{?\.?Definitions\b') {
        $inDefinitions = $true
        $outLines.Add($line)
        continue
      }
    }

    if ($line -match '^\s*(```+|~~~+)') {
      $Token = $Matches[1]
      $char = $Token.Substring(0, 1)
      if (-not $inFence) {
        $inFence = $true
        $fenceChar = $char
      }
      elseif ($fenceChar -eq $char) {
        $inFence = $false
        $fenceChar = ""
      }
      $outLines.Add($line)
      continue
    }

    if ($inFence) {
      $outLines.Add($line)
      continue
    }

    # Indented article markers should remain semantic list markers, not code blocks.
    if ($line -match '^    ([ivxlcdmIVXLCDM]+\.\s+.+)$') {
      $outLines.Add($Matches[1])
      continue
    }

    if ($line -match '^  ([A-Za-z]\)\s+.+)$') {
      $outLines.Add($Matches[1])
      continue
    }

    $prevBlank = ($i -eq 0) -or [string]::IsNullOrWhiteSpace($lines[$i - 1])
    $nextBlank = ($i -eq ($lines.Length - 1)) -or [string]::IsNullOrWhiteSpace($lines[$i + 1])

    # Standalone indented bullets map explicitly to List 2 / List 3.
    if ($prevBlank -and $line -match '^  -\s+(.+)$') {
      $outLines.Add("- $($Matches[1]) {.List-2}")
      continue
    }

    if ($prevBlank -and $line -match '^    -\s+(.+)$') {
      $outLines.Add("- $($Matches[1]) {.List-3}")
      continue
    }

    if ($line -match '^  (\S.*)$') {
      $body = $Matches[1]
      $isListLike = ($body -match '^[-*+]\s') -or
      ($body -match '^\d+[.)]\s') -or
      ($body -match '^[A-Za-z][.)]\s') -or
      ($body -match '^[ivxlcdmIVXLCDM]+\.\s')
      $isTableLike = $body -match '^\|'
      $isDivFence = $body -match '^:::'
      $isHeadingLike = $body -match '^#{1,6}\s'

      if (-not $isListLike -and -not $isTableLike -and -not $isDivFence -and -not $isHeadingLike) {
        $outLines.Add("$body {.Offset}")
        continue
      }
    }

    $outLines.Add($line)
  }

  return ($outLines -join "`n")
}

function Convert-LargeTableToGfm {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Markdown
  )

  $lines = $Markdown -replace "`r`n", "`n" -split "`n", -1
  $outLines = [System.Collections.Generic.List[string]]::new()
  $i = 0

  while ($i -lt $lines.Length) {

    if ($lines[$i] -match '^:::\s*LargeTable\s*$') {
      $i++

      $allRows    = [System.Collections.Generic.List[object]]::new()
      $cellLines  = [System.Collections.Generic.List[string]]::new()
      $rowCells   = [System.Collections.Generic.List[object]]::new()

      while ($i -lt $lines.Length -and $lines[$i] -notmatch '^:::\s*$') {
        $l = $lines[$i]; $i++

        if ($l -match '^(.*) \|\|$') {
          $cellLines.Add($Matches[1])
          $rowCells.Add($cellLines.ToArray())
          $allRows.Add($rowCells.ToArray())
          $cellLines = [System.Collections.Generic.List[string]]::new()
          $rowCells  = [System.Collections.Generic.List[object]]::new()
        }
        elseif ($l -match '^(.*) \|$') {
          $cellLines.Add($Matches[1])
          $rowCells.Add($cellLines.ToArray())
          $cellLines = [System.Collections.Generic.List[string]]::new()
        }
        else {
          $cellLines.Add($l)
        }
      }
      if ($i -lt $lines.Length) { $i++ }

      if ($allRows.Count -eq 0) { continue }

      $colCount = ([object[]]$allRows[0]).Count

      $flattenCell = [scriptblock]{
        param([string[]]$cellLineArr)
        $groups = [System.Collections.Generic.List[string]]::new()
        $group  = [System.Collections.Generic.List[string]]::new()
        foreach ($ln in $cellLineArr) {
          if ([string]::IsNullOrWhiteSpace($ln)) {
            if ($group.Count -gt 0) {
              $groups.Add(($group.ToArray() -join '<br>'))
              $group = [System.Collections.Generic.List[string]]::new()
            }
          } else {
            $group.Add($ln)
          }
        }
        if ($group.Count -gt 0) { $groups.Add(($group.ToArray() -join '<br>')) }
        $flat = $groups.ToArray() -join '<br>'
        $flat = $flat.Replace('|', '\|')
        if ([string]::IsNullOrWhiteSpace($flat)) { return ' ' }
        return $flat.Trim()
      }

      $flatRows = foreach ($row in $allRows) {
        $cells = foreach ($cellArr in [object[]]$row) {
          & $flattenCell ([string[]]$cellArr)
        }
        ,@($cells)
      }

      $colWidths = @(3) * $colCount
      foreach ($row in $flatRows) {
        $rowArr = [string[]]$row
        for ($c = 0; $c -lt [Math]::Min($colCount, $rowArr.Length); $c++) {
          $w = $rowArr[$c].Length
          if ($w -gt $colWidths[$c]) { $colWidths[$c] = $w }
        }
      }

      $hArr   = [string[]]$flatRows[0]
      $hParts = for ($c = 0; $c -lt $colCount; $c++) {
        $v = if ($c -lt $hArr.Length) { $hArr[$c] } else { ' ' }
        $v.PadRight($colWidths[$c])
      }
      $outLines.Add('| ' + ($hParts -join ' | ') + ' |')

      $sepParts = $colWidths | ForEach-Object { '-' * $_ }
      $outLines.Add('| ' + ($sepParts -join ' | ') + ' |')

      for ($r = 1; $r -lt $flatRows.Count; $r++) {
        $rArr   = [string[]]$flatRows[$r]
        $rParts = for ($c = 0; $c -lt $colCount; $c++) {
          $v = if ($c -lt $rArr.Length) { $rArr[$c] } else { ' ' }
          $v.PadRight($colWidths[$c])
        }
        $outLines.Add('| ' + ($rParts -join ' | ') + ' |')
      }

      continue
    }

    $outLines.Add($lines[$i])
    $i++
  }

  return $outLines -join "`n"
}

function Get-UniqueFilePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$CandidatePath
  )

  if (-not (Test-Path -LiteralPath $CandidatePath)) {
    return $CandidatePath
  }

  $dir = Split-Path -Parent $CandidatePath
  $ext = [System.IO.Path]::GetExtension($CandidatePath)
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($CandidatePath)

  if ($baseName -match '^(.+)_(\d+)$') {
    $root = $Matches[1]
    $counter = [int]$Matches[2] + 1
  } else {
    $root = $baseName
    $counter = 1
  }

  do {
    $candidate = Join-Path $dir "${root}_${counter}${ext}"
    $counter++
  } while (Test-Path -LiteralPath $candidate)

  return $candidate
}

function Stage-MarkdownImages {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Markdown,

    [Parameter(Mandatory = $true)]
    [string]$MarkdownPath,

    [Parameter(Mandatory = $true)]
    [string]$TempRoot
  )

  $sourceDirectory = Split-Path -Parent $MarkdownPath
  $stagedMediaDirectory = Join-Path $TempRoot "_media"
  if (-not (Test-Path -LiteralPath $stagedMediaDirectory)) {
    New-Item -ItemType Directory -Path $stagedMediaDirectory -Force | Out-Null
  }

  $stagedBySourcePath = @{}
  $pattern = [regex]'!\[(?<alt>[^\]]*)\]\((?<target>[^)\r\n]+)\)'

  return $pattern.Replace($Markdown, {
    param($match)

    $altText = $match.Groups['alt'].Value
    $rawTarget = $match.Groups['target'].Value.Trim()

    if ([string]::IsNullOrWhiteSpace($rawTarget)) {
      return $match.Value
    }

    $pathPart = $rawTarget
    $titlePart = ""

    if ($rawTarget.StartsWith("<")) {
      $closing = $rawTarget.IndexOf(">")
      if ($closing -gt 0) {
        $pathPart = $rawTarget.Substring(1, $closing - 1)
        $titlePart = $rawTarget.Substring($closing + 1).TrimStart()
      }
    } elseif ($rawTarget -match "^(?<path>\S+)(?<rest>\s+[""'].*)?$") {
      $pathPart = $Matches["path"]
      if ($Matches.ContainsKey("rest") -and $null -ne $Matches["rest"]) {
        $titlePart = $Matches["rest"].TrimStart()
      } else {
        $titlePart = ""
      }
    }

    if ($pathPart -match '^(?i)(https?|ftp|ftps|mailto|tel|data|file):') {
      return $match.Value
    }

    if ($pathPart.StartsWith("#")) {
      return $match.Value
    }

    $relativeSourcePath = $pathPart -replace '/', [System.IO.Path]::DirectorySeparatorChar
    $resolvedSourcePath = if ([System.IO.Path]::IsPathRooted($relativeSourcePath)) {
      $relativeSourcePath
    } else {
      Join-Path $sourceDirectory $relativeSourcePath
    }

    if (-not (Test-Path -LiteralPath $resolvedSourcePath -PathType Leaf)) {
      throw "Image file referenced in markdown was not found: $pathPart (resolved from $MarkdownPath)"
    }

    $resolvedSourcePath = [System.IO.Path]::GetFullPath($resolvedSourcePath)
    $cacheKey = $resolvedSourcePath.ToLowerInvariant()

    if (-not $stagedBySourcePath.ContainsKey($cacheKey)) {
      $fileName = [System.IO.Path]::GetFileName($resolvedSourcePath)
      if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "image.bin"
      }

      $destinationPath = Get-UniqueFilePath -CandidatePath (Join-Path $stagedMediaDirectory $fileName)
      Copy-Item -LiteralPath $resolvedSourcePath -Destination $destinationPath
      $stagedBySourcePath[$cacheKey] = "_media/" + [System.IO.Path]::GetFileName($destinationPath)
    }

    $stagedRelativePath = $stagedBySourcePath[$cacheKey]
    if ($stagedRelativePath -match '\s') {
      $stagedRelativePath = "<$stagedRelativePath>"
    }

    $titleSuffix = if ([string]::IsNullOrWhiteSpace($titlePart)) { "" } else { " $titlePart" }
    return "![${altText}]($stagedRelativePath$titleSuffix)"
  })
}

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

if (-not (Test-Path -LiteralPath $cssPath)) {
  throw "Missing stylesheet: $cssPath"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension

$preparedMarkdown = Convert-OffsetPrefixesToStyleSuffix -InputPath $resolvedInput
$preparedMarkdown = Convert-LargeTableToGfm -Markdown $preparedMarkdown
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("m2h_" + [System.Guid]::NewGuid().ToString("N"))
$tempInput = Join-Path $tempRoot "input.md"
$tempHeader = Join-Path $tempRoot "header.html"
$preparedMarkdown = Stage-MarkdownImages -Markdown $preparedMarkdown -MarkdownPath $resolvedInput -TempRoot $tempRoot
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
[System.IO.File]::WriteAllText($tempInput, $preparedMarkdown, $utf8NoBom)

# The stylesheet is inlined via --include-in-header rather than --css so the
# output is a single self-contained .html file, matching how m2d/m2p produce
# a single .docx/.pdf with no side files to keep track of.
$cssContent = [System.IO.File]::ReadAllText($cssPath, [System.Text.Encoding]::UTF8)
$headerHtml = "<style>`n$cssContent`n</style>"
[System.IO.File]::WriteAllText($tempHeader, $headerHtml, $utf8NoBom)

$titleMetadata = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)

try {
  & pandoc `
    -f "markdown+fancy_lists+lists_without_preceding_blankline+fenced_divs-smart" `
    -t "html5" `
    --standalone `
    --embed-resources `
    --no-highlight `
    --resource-path="$tempRoot" `
    --include-in-header="$tempHeader" `
    --lua-filter="$filterPath" `
    --metadata="title:$titleMetadata" `
    "$tempInput" `
    -o "$resolvedOutput"

  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }
}
finally {
  # Cleanup must never mask the failure that got us here.
  if (Test-Path -LiteralPath $tempRoot) {
    try { Remove-Item -LiteralPath $tempRoot -Force -Recurse } catch { }
  }
}

Write-Output "Created: $resolvedOutput"
