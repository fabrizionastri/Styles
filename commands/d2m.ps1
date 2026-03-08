[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "..\filters\docx_to_compact.lua"
$defaultInputExtension = ".docx"
$defaultOutputExtension = ".md"

function Resolve-InputPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue,

    [Parameter(Mandatory = $true)]
    [string]$DefaultExtension
  )

  if (Test-Path -LiteralPath $PathValue) {
    return (Resolve-Path -LiteralPath $PathValue -ErrorAction Stop).Path
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($PathValue))) {
    $withExt = $PathValue + $DefaultExtension
    if (Test-Path -LiteralPath $withExt) {
      return (Resolve-Path -LiteralPath $withExt -ErrorAction Stop).Path
    }
  }

  throw "Input file not found: $PathValue"
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

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $resolved = Join-Path $inputDir ($inputBase + $DefaultExtension)
  } else {
    if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($OutputPath))) {
      $OutputPath = $OutputPath + $DefaultExtension
    }
    if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
      $OutputPath = Join-Path $inputDir $OutputPath
    }
    $resolved = $OutputPath
  }

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

function ConvertTo-AnchorId {
  param([string]$DisplayText)
  $slug = $DisplayText.ToLower()
  $slug = $slug -replace '([0-9])([a-z])', '$1-$2'
  $slug = $slug -replace '([a-z])([0-9])', '$1-$2'
  $slug = $slug -replace '[^a-z0-9]+', '-'
  $slug = $slug.Trim('-')
  return "ref-$slug"
}

function Extract-CrossReferences {
  param([string]$DocxPath)

  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
  $nsUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

  try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $content = $reader.ReadToEnd()
    $reader.Close()
    $stream.Close()
  }
  finally {
    $zip.Dispose()
  }

  $xml = [xml]$content
  $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace('w', $nsUri)

  $bookmarks = @{}
  $refs = [System.Collections.ArrayList]::new()

  foreach ($para in $xml.SelectNodes('//w:p', $ns)) {
    foreach ($bm in $para.SelectNodes('.//w:bookmarkStart', $ns)) {
      $name = $bm.GetAttribute('name', $nsUri)
      if ($name -and $name -like '_Ref*') {
        $texts = $para.SelectNodes('.//w:t', $ns)
        $paraText = ($texts | ForEach-Object { $_.InnerText }) -join ''
        $bookmarks[$name] = @{ ParagraphText = $paraText.Trim() }
      }
    }

    $depth = 0
    $instrAccum = ''
    $displayAccum = ''
    $inDisplay = $false

    foreach ($run in $para.ChildNodes) {
      if ($run.LocalName -ne 'r') { continue }

      $fc = $run.SelectSingleNode('w:fldChar', $ns)
      if ($fc) {
        $fcType = $fc.GetAttribute('fldCharType', $nsUri)
        if ($fcType -eq 'begin') {
          $depth++
          if ($depth -eq 1) {
            $instrAccum = ''
            $displayAccum = ''
            $inDisplay = $false
          }
        }
        elseif ($fcType -eq 'separate' -and $depth -eq 1) {
          $inDisplay = $true
        }
        elseif ($fcType -eq 'end') {
          if ($depth -eq 1 -and $instrAccum -match 'REF\s+(_Ref\w+)(.*)') {
            $null = $refs.Add(@{
              BookmarkName = $Matches[1]
              DisplayText  = $displayAccum.Trim()
              Switches     = $Matches[2].Trim()
            })
          }
          $depth = [Math]::Max(0, $depth - 1)
          if ($depth -eq 0) { $inDisplay = $false }
        }
        continue
      }

      $instrEl = $run.SelectSingleNode('w:instrText', $ns)
      if ($instrEl -and $depth -ge 1) {
        $instrAccum += $instrEl.InnerText
      }

      if ($inDisplay -and $depth -eq 1) {
        $tEl = $run.SelectSingleNode('w:t', $ns)
        if ($tEl) { $displayAccum += $tEl.InnerText }
      }
    }
  }

  return @{ Bookmarks = $bookmarks; Refs = $refs }
}

function Find-AnchorLine {
  param(
    [string[]]$Lines,
    [string]$DisplayText
  )

  $remaining = $DisplayText
  $patterns = @()

  if ($remaining -match '^(\d+\.\d+)(.*)$') {
    $num = $Matches[1]
    $remaining = $Matches[2]
    $escaped = [regex]::Escape($num)
    $patterns += "^${escaped}\.?\s"
  }
  else {
    return -1
  }

  if ($remaining -match '^([a-z])\)(.*)$') {
    $letter = $Matches[1]
    $remaining = $Matches[2]
    $patterns += "^${letter}\)\s"
  }

  if ($remaining -match '^([ivxlcdm]+)$') {
    $roman = $Matches[1]
    $patterns += "^${roman}\.\s"
  }

  $searchFrom = 0
  $lastFound = -1

  foreach ($pat in $patterns) {
    $found = $false
    for ($i = $searchFrom; $i -lt $Lines.Count; $i++) {
      if ($Lines[$i].TrimStart() -match $pat) {
        $lastFound = $i
        $searchFrom = $i + 1
        $found = $true
        break
      }
    }
    if (-not $found) { return -1 }
  }

  return $lastFound
}

function Add-CrossReferences {
  param(
    [string]$Markdown,
    [hashtable]$XRefData
  )

  if ($XRefData.Refs.Count -eq 0) { return $Markdown }

  $refsByBookmark = @{}
  foreach ($ref in $XRefData.Refs) {
    $bm = $ref.BookmarkName
    if (-not $XRefData.Bookmarks.ContainsKey($bm)) { continue }
    if (-not $refsByBookmark.ContainsKey($bm)) {
      $refsByBookmark[$bm] = [System.Collections.ArrayList]::new()
    }
    $null = $refsByBookmark[$bm].Add($ref)
  }

  if ($refsByBookmark.Count -eq 0) { return $Markdown }

  $anchorMap = @{}
  foreach ($bmName in $refsByBookmark.Keys) {
    $display = $refsByBookmark[$bmName][0].DisplayText
    $anchorMap[$bmName] = ConvertTo-AnchorId $display
  }

  $sep = if ($Markdown.Contains("`r`n")) { "`r`n" } else { "`n" }
  $lines = $Markdown -split $sep

  foreach ($bmName in $refsByBookmark.Keys) {
    $display = $refsByBookmark[$bmName][0].DisplayText
    $anchorId = $anchorMap[$bmName]
    $idx = Find-AnchorLine -Lines $lines -DisplayText $display
    if ($idx -ge 0) {
      $lines[$idx] = $lines[$idx].TrimEnd() + " []{#$anchorId}"
    }
  }

  $result = $lines -join $sep

  $sortedRefs = $XRefData.Refs | Sort-Object { $_.DisplayText.Length } -Descending

  foreach ($ref in $sortedRefs) {
    $anchorId = $anchorMap[$ref.BookmarkName]
    if (-not $anchorId) { continue }

    $escaped = [regex]::Escape($ref.DisplayText)
    $pattern = "(Article\s+${escaped})(?![a-zA-Z0-9\)\.])"
    $replacement = '[$1](#' + $anchorId + ')'
    $newResult = [regex]::Replace($result, $pattern, $replacement)

    if ($newResult -ne $result) {
      $result = $newResult
    }
    else {
      $pattern = "(?<=See\s+)(${escaped})(?![a-zA-Z0-9\)\.])"
      $replacement = '[$1](#' + $anchorId + ')'
      $result = [regex]::Replace($result, $pattern, $replacement)
    }
  }

  return $result
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

function Get-RelativeMarkdownPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$FromDirectory,

    [Parameter(Mandatory = $true)]
    [string]$ToPath
  )

  $fromFull = [System.IO.Path]::GetFullPath($FromDirectory)
  if (-not $fromFull.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
    $fromFull += [System.IO.Path]::DirectorySeparatorChar
  }

  $toFull = [System.IO.Path]::GetFullPath($ToPath)
  $fromUri = [System.Uri]$fromFull
  $toUri = [System.Uri]$toFull
  $relative = $fromUri.MakeRelativeUri($toUri).ToString()

  return [System.Uri]::UnescapeDataString($relative).Replace('\', '/')
}

function Rewrite-ExtractedImageLinks {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Markdown,

    [Parameter(Mandatory = $true)]
    [string]$ExtractRoot,

    [Parameter(Mandatory = $true)]
    [string]$TargetMediaDirectory,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
  )

  $outputDir = Split-Path -Parent $OutputPath
  $copiedPaths = @{}
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

    $relativeExtractPath = $pathPart -replace '/', [System.IO.Path]::DirectorySeparatorChar
    $sourcePath = if ([System.IO.Path]::IsPathRooted($relativeExtractPath)) {
      $relativeExtractPath
    } else {
      Join-Path $ExtractRoot $relativeExtractPath
    }

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
      return $match.Value
    }

    $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
    $cacheKey = $sourcePath.ToLowerInvariant()

    if (-not $copiedPaths.ContainsKey($cacheKey)) {
      if (-not (Test-Path -LiteralPath $TargetMediaDirectory)) {
        New-Item -ItemType Directory -Path $TargetMediaDirectory -Force | Out-Null
      }

      $fileName = [System.IO.Path]::GetFileName($sourcePath)
      if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "image.bin"
      }

      $targetPath = Get-UniqueFilePath -CandidatePath (Join-Path $TargetMediaDirectory $fileName)
      Copy-Item -LiteralPath $sourcePath -Destination $targetPath

      $relativeTarget = Get-RelativeMarkdownPath -FromDirectory $outputDir -ToPath $targetPath
      $copiedPaths[$cacheKey] = @{
        Path = $relativeTarget
        Name = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
      }
    }

    $mapped = $copiedPaths[$cacheKey]
    $newAlt = if ([string]::IsNullOrWhiteSpace($altText)) { $mapped.Name } else { $altText }
    $newTarget = $mapped.Path
    if ($newTarget -match '\s') {
      $newTarget = "<$newTarget>"
    }

    $titleSuffix = if ([string]::IsNullOrWhiteSpace($titlePart)) { "" } else { " $titlePart" }
    return "![${newAlt}]($newTarget$titleSuffix)"
  })
}

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension
$sourceDirectory = Split-Path -Parent $resolvedInput
$targetMediaDirectory = Join-Path $sourceDirectory "_media"

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("d2m_" + [System.Guid]::NewGuid().ToString("N"))
$tempOutput = Join-Path $tempRoot "output.md"
$tempExtract = Join-Path $tempRoot "media_extract"

try {
  New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

  $xrefData = Extract-CrossReferences -DocxPath $resolvedInput

  & pandoc `
    -f "docx+styles" `
    -t "markdown+fenced_divs" `
    --wrap=none `
    --extract-media="$tempExtract" `
    --lua-filter="$filterPath" `
    "$resolvedInput" `
    -o "$tempOutput"

  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $rendered = [System.IO.File]::ReadAllText($tempOutput, $utf8NoBom)
  $rendered = Add-CrossReferences -Markdown $rendered -XRefData $xrefData
  $rendered = Rewrite-ExtractedImageLinks `
    -Markdown $rendered `
    -ExtractRoot $tempExtract `
    -TargetMediaDirectory $targetMediaDirectory `
    -OutputPath $resolvedOutput
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($resolvedOutput, $rendered, $utf8Bom)
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Force -Recurse
  }
}

Write-Output "Created: $resolvedOutput"
