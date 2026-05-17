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
      if ($name -and $name -notmatch '^_GoBack') {
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
          if ($depth -eq 1 -and $instrAccum -match '(?i)\bREF\s+([A-Za-z_][A-Za-z0-9_]*)\b(.*)') {
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

  $remaining = $DisplayText.Trim()
  $hadArticlePrefix = $false

  if ($remaining -match '^(?i)Article\s+(.+)$') {
    $hadArticlePrefix = $true
    $remaining = $Matches[1].Trim()
  }

  $patterns = @()

  if ($remaining -match '^(\d+\.\d+)\.?(.*)$') {
    $num = $Matches[1]
    $remaining = $Matches[2]
    $escaped = [regex]::Escape($num)
    if ($hadArticlePrefix) {
      $patterns += "^(?:#{1,6}\s+)?Article\s+${escaped}\.?\b"
    } else {
      $patterns += "^${escaped}\.?\s"
    }
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

  function Convert-SwitchesToMarkdownXref {
    param([string]$Switches)

    $tokens = [regex]::Matches($Switches, '\\[A-Za-z]+') | ForEach-Object { $_.Value }
    if ($tokens.Count -eq 0) {
      return '\h\n'
    }
    return ($tokens -join '')
  }

  function Normalize-XrefDisplayText {
    param([string]$DisplayText)

    $text = if ($null -eq $DisplayText) { "" } else { $DisplayText.Trim() }
    if ([string]::IsNullOrWhiteSpace($text)) {
      return ""
    }

    # Defensive normalization: some editors/renderers may substitute square brackets.
    $text = $text.Replace('‹', '[').Replace('›', ']')

    if ($text -match '\\?\[(?<label>[^\]]+)\]\\?\]\(#') {
      return $Matches['label'].Trim()
    }

    return $text
  }

  $anchorMap = @{}
  foreach ($bmName in $refsByBookmark.Keys) {
    $display = Normalize-XrefDisplayText -DisplayText $refsByBookmark[$bmName][0].DisplayText
    if ([string]::IsNullOrWhiteSpace($display)) {
      continue
    }
    $anchorMap[$bmName] = $bmName
  }

  $sep = if ($Markdown.Contains("`r`n")) { "`r`n" } else { "`n" }
  $lines = $Markdown -split $sep

  foreach ($bmName in $refsByBookmark.Keys) {
    $display = Normalize-XrefDisplayText -DisplayText $refsByBookmark[$bmName][0].DisplayText
    $anchorId = $anchorMap[$bmName]
    if (-not $anchorId) { continue }
    if ([string]::IsNullOrWhiteSpace($display)) { continue }
    $idx = Find-AnchorLine -Lines $lines -DisplayText $display
    if ($idx -ge 0) {
      $anchorToken = "[]{#$anchorId}"
      if ($lines[$idx] -notmatch [regex]::Escape($anchorToken)) {
        $lines[$idx] = $lines[$idx].TrimEnd() + " $anchorToken"
      }
    }
  }

  $result = $lines -join $sep

  $sortedRefs = $XRefData.Refs | Sort-Object { $_.DisplayText.Length } -Descending

  foreach ($ref in $sortedRefs) {
    $anchorId = $anchorMap[$ref.BookmarkName]
    if (-not $anchorId) { continue }
    $displayText = Normalize-XrefDisplayText -DisplayText $ref.DisplayText
    if ([string]::IsNullOrWhiteSpace($displayText)) { continue }
    $xrefSwitches = Convert-SwitchesToMarkdownXref -Switches $ref.Switches
    $linkSuffix = ' "xref:' + $xrefSwitches + '")'

    $escaped = [regex]::Escape($displayText)
    if ($displayText -notmatch '^(?i)Article\s+') {
      $pattern = "(Article\s+${escaped})(?![a-zA-Z0-9\)])"
      $replacement = '[$1](#' + $anchorId + $linkSuffix
      $newResult = [regex]::Replace($result, $pattern, $replacement)

      if ($newResult -ne $result) {
        $result = $newResult
        continue
      }
    }

    $pattern = "(?<=See\s+)(${escaped})(?![a-zA-Z0-9\)])"
    $replacement = '[$1](#' + $anchorId + $linkSuffix
    $newResult = [regex]::Replace($result, $pattern, $replacement)
    if ($newResult -ne $result) {
      $result = $newResult
      continue
    }

    $pattern = "(?<=in\s+)(${escaped})(?![a-zA-Z0-9\)])"
    $replacement = '[$1](#' + $anchorId + $linkSuffix
    $result = [regex]::Replace($result, $pattern, $replacement)
  }

  return $result
}

function Remove-MarkdownEscapes {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Markdown
  )

  # Remove backslash escapes that pandoc adds but are unnecessary in our
  # contract-template format: [ ] < and - . Grid-table rows are skipped:
  # stripping escapes there would shrink cells and break the column
  # alignment that grid tables require, corrupting the table on reparse.
  $sep = if ($Markdown.Contains("`r`n")) { "`r`n" } else { "`n" }
  $lines = $Markdown -split $sep, -1
  $inGrid = $false

  for ($i = 0; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]

    if ($line -match '^\s*\+[-=+]+\+\s*$') {
      $inGrid = $true
      continue
    }

    if ($inGrid) {
      if ($line -match '^\s*[|+]') {
        continue
      }
      $inGrid = $false
    }

    $line = $line -replace '\\\[', '['
    $line = $line -replace '\\\]', ']'
    $line = $line -replace '\\<',  '<'
    $line = $line -replace '\\\-', '-'
    $lines[$i] = $line
  }

  return ($lines -join $sep)
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

function Convert-ToMediaFileStem {
  param(
    [string]$Value,
    [string]$Fallback
  )

  if ($null -eq $Value) {
    $candidate = ""
  } else {
    $candidate = [string]$Value
  }
  $candidate = $candidate.ToLowerInvariant()
  $candidate = $candidate -replace '[^a-z0-9]+', '_'
  $candidate = $candidate.Trim('_')

  if ([string]::IsNullOrWhiteSpace($candidate)) {
    return $Fallback
  }

  return $candidate
}

function Get-WordImageNameMap {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath
  )

  $map = @{}

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)

  try {
    $docEntry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    if (-not $docEntry) { return $map }

    $docStream = $docEntry.Open()
    $docReader = New-Object System.IO.StreamReader($docStream)
    $docXmlRaw = $docReader.ReadToEnd()
    $docReader.Close()
    $docStream.Close()

    $docXml = [xml]$docXmlRaw
    $ns = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
    $ns.AddNamespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

    foreach ($drawing in $docXml.SelectNodes('//w:drawing', $ns)) {
      $blip = $drawing.SelectSingleNode('.//a:blip[@r:embed]', $ns)
      if (-not $blip) { continue }

      $relId = $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      if ([string]::IsNullOrWhiteSpace($relId)) { continue }

      $docPr = $drawing.SelectSingleNode('.//wp:docPr', $ns)
      $cNvPr = $drawing.SelectSingleNode('.//pic:cNvPr', $ns)

      $descrCandidates = @(
        $(if ($docPr) { $docPr.GetAttribute('descr') } else { $null }),
        $(if ($cNvPr) { $cNvPr.GetAttribute('descr') } else { $null })
      )
      $nameCandidates = @(
        $(if ($docPr) { $docPr.GetAttribute('name') } else { $null }),
        $(if ($cNvPr) { $cNvPr.GetAttribute('name') } else { $null })
      )

      $chosen = $null
      foreach ($value in $descrCandidates + $nameCandidates) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($value -match '^(?i)(picture|image)\s+\d+$') { continue }
        $chosen = $value
        break
      }

      if ([string]::IsNullOrWhiteSpace($chosen)) { continue }
      $map[$relId.ToLowerInvariant()] = $chosen
    }
  }
  finally {
    $zip.Dispose()
  }

  return $map
}

function Resolve-TargetMediaPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetMediaDirectory,

    [Parameter(Mandatory = $true)]
    [string]$PreferredFileName
  )

  return (Join-Path $TargetMediaDirectory $PreferredFileName)
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
    [hashtable]$ImageNameMap,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
  )

  $outputDir = Split-Path -Parent $OutputPath
  $copiedPaths = @{}
  $altMarkers = @{}
  $markerCounter = 0
  $pattern = [regex]'!\[(?<alt>[^\]]*)\]\((?<target>[^)\r\n]+)\)(?<attrs>\{[^\r\n]*\})?'
  $existingByHash = @{}

  if (Test-Path -LiteralPath $TargetMediaDirectory -PathType Container) {
    foreach ($existingFile in Get-ChildItem -LiteralPath $TargetMediaDirectory -File) {
      $existingHash = (Get-FileHash -LiteralPath $existingFile.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
      if (-not $existingByHash.ContainsKey($existingHash)) {
        $existingByHash[$existingHash] = $existingFile.Name
      }
    }
  }

  $rewritten = $pattern.Replace($Markdown, {
    param($match)

    $altText = $match.Groups['alt'].Value
    $rawTarget = $match.Groups['target'].Value.Trim()
    $attrs = $match.Groups['attrs'].Value

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

      $sourceStem = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)
      $sourceExt = [System.IO.Path]::GetExtension($sourcePath)
      if ([string]::IsNullOrWhiteSpace($sourceExt)) {
        $sourceExt = ".bin"
      }

      $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
      $preferredFileName = $null

      if ($existingByHash.ContainsKey($sourceHash)) {
        $preferredFileName = [string]$existingByHash[$sourceHash]
      }

      if ([string]::IsNullOrWhiteSpace($preferredFileName)) {
        $nameHint = $null
        if (-not [string]::IsNullOrWhiteSpace($sourceStem)) {
          $key = $sourceStem.ToLowerInvariant()
          if ($ImageNameMap.ContainsKey($key)) {
            $nameHint = [string]$ImageNameMap[$key]
          }
        }
        $fallbackStem = if ([string]::IsNullOrWhiteSpace($sourceStem)) { "image" } else { $sourceStem.ToLowerInvariant() }
        $preferredStem = Convert-ToMediaFileStem -Value $nameHint -Fallback $fallbackStem
        $preferredFileName = "$preferredStem$sourceExt"
      }

      $targetPath = Resolve-TargetMediaPath `
        -SourcePath $sourcePath `
        -TargetMediaDirectory $TargetMediaDirectory `
        -PreferredFileName $preferredFileName

      Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
      $existingByHash[$sourceHash] = [System.IO.Path]::GetFileName($targetPath)

      $relativeTarget = Get-RelativeMarkdownPath -FromDirectory $outputDir -ToPath $targetPath
      $copiedPaths[$cacheKey] = @{
        Path = $relativeTarget
      }
    }

    $mapped = $copiedPaths[$cacheKey]
    $newTarget = $mapped.Path
    if ($newTarget -match '\s') {
      $newTarget = "<$newTarget>"
    }

    $marker = ""
    if (-not [string]::IsNullOrWhiteSpace($altText)) {
      $markerCounter++
      $altMarkers[$markerCounter] = $altText.Trim()
      $marker = "<!--codex-img-alt-id:$markerCounter-->"
    }

    $titleSuffix = if ([string]::IsNullOrWhiteSpace($titlePart)) { "" } else { " $titlePart" }
    return "![]($newTarget$titleSuffix)$attrs$marker"
  })

  if ($altMarkers.Count -eq 0) {
    return $rewritten
  }

  $sep = if ($rewritten.Contains("`r`n")) { "`r`n" } else { "`n" }
  $lines = $rewritten -split $sep, -1
  $outLines = New-Object 'System.Collections.Generic.List[string]'
  $markerPattern = [regex]'<!--codex-img-alt-id:(?<id>\d+)-->'

  $i = 0
  while ($i -lt $lines.Length) {
    $line = $lines[$i]
    $markerMatch = $markerPattern.Match($line)

    if ($markerMatch.Success) {
      $markerId = [int]$markerMatch.Groups['id'].Value
      $altText = if ($altMarkers.ContainsKey($markerId)) {
        [string]$altMarkers[$markerId]
      } else {
        ""
      }

      $line = $markerPattern.Replace($line, "")
      $outLines.Add($line)

      $j = $i + 1
      $hasLeadingBlank = $false
      if ($j -lt $lines.Length -and [string]::IsNullOrWhiteSpace($lines[$j])) {
        $hasLeadingBlank = $true
        $j++
      }

      if (-not [string]::IsNullOrWhiteSpace($altText) -and
          $j -lt $lines.Length -and
          $lines[$j].Trim() -eq $altText) {
        $i = $j + 1
        if ($i -lt $lines.Length -and [string]::IsNullOrWhiteSpace($lines[$i])) {
          $i++
        }
        if ($hasLeadingBlank) {
          $outLines.Add("")
        }
        continue
      }

      $i++
      continue
    }

    $outLines.Add($line)
    $i++
  }

  return ($outLines -join $sep)
}

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension
$outputDirectory = Split-Path -Parent $resolvedOutput
$targetMediaDirectory = Join-Path $outputDirectory "_media"

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("d2m_" + [System.Guid]::NewGuid().ToString("N"))
$tempOutput = Join-Path $tempRoot "output.md"
$tempExtract = Join-Path $tempRoot "media_extract"

try {
  New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

  $tempInput = Join-Path $tempRoot "input.docx"
  Copy-Item -LiteralPath $resolvedInput -Destination $tempInput -Force

  $xrefData = Extract-CrossReferences -DocxPath $tempInput
  $imageNameMap = Get-WordImageNameMap -DocxPath $tempInput

  & pandoc `
    -f "docx+styles" `
    -t "markdown+fenced_divs-smart-simple_tables-multiline_tables" `
    --wrap=none `
    --extract-media="$tempExtract" `
    --lua-filter="$filterPath" `
    "$tempInput" `
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
    -ImageNameMap $imageNameMap `
    -OutputPath $resolvedOutput
  $rendered = Remove-MarkdownEscapes -Markdown $rendered
  if (-not ($rendered.EndsWith("`n") -or $rendered.EndsWith("`r"))) {
    $newline = if ($rendered.Contains("`r`n")) { "`r`n" } else { "`n" }
    $rendered += $newline
  }
  [System.IO.File]::WriteAllText($resolvedOutput, $rendered, $utf8NoBom)
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Force -Recurse
  }
}

Write-Output "Created: $resolvedOutput"
