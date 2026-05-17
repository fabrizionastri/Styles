[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "..\filters\compact_to_docx_v1.lua"
$referenceDoc = Join-Path $scriptRoot "..\styles\flexup_template.docx"
$defaultInputExtension = ".md"
$defaultOutputExtension = ".docx"

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

  for ($i = 0; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]

    if ($line -match '^\s*(```+|~~~+)') {
      $token = $Matches[1]
      $char = $token.Substring(0, 1)
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

  # Convert ::: LargeTable blocks back to standard GFM pipe tables so that
  # pandoc + compact_to_docx.lua can process them normally.
  #
  # Cell format  : content line(s) ending with " |"  (trailing space + pipe)
  # Row format   : last cell of the row ends with " ||" (trailing space + double-pipe)
  # First row    : treated as the table header.
  # Multi-para   : blank lines within a cell are paragraph separators; each
  #               non-blank group is joined with <br> before entering the cell.

  $lines = $Markdown -replace "`r`n", "`n" -split "`n", -1
  $outLines = [System.Collections.Generic.List[string]]::new()
  $i = 0

  while ($i -lt $lines.Length) {

    if ($lines[$i] -match '^:::\s*LargeTable\s*$') {
      $i++  # skip fence-open line

      # Collect rows: list of (list of string[])  row -> cells -> content lines
      $allRows    = [System.Collections.Generic.List[object]]::new()
      $cellLines  = [System.Collections.Generic.List[string]]::new()
      $rowCells   = [System.Collections.Generic.List[object]]::new()

      while ($i -lt $lines.Length -and $lines[$i] -notmatch '^:::\s*$') {
        $l = $lines[$i]; $i++

        if ($l -match '^(.*) \|\|$') {
          # End of cell AND row
          $cellLines.Add($Matches[1])
          $rowCells.Add($cellLines.ToArray())
          $allRows.Add($rowCells.ToArray())
          $cellLines = [System.Collections.Generic.List[string]]::new()
          $rowCells  = [System.Collections.Generic.List[object]]::new()
        }
        elseif ($l -match '^(.*) \|$') {
          # End of cell (row continues)
          $cellLines.Add($Matches[1])
          $rowCells.Add($cellLines.ToArray())
          $cellLines = [System.Collections.Generic.List[string]]::new()
        }
        else {
          $cellLines.Add($l)
        }
      }
      if ($i -lt $lines.Length) { $i++ }  # skip fence-close line

      if ($allRows.Count -eq 0) { continue }

      $colCount = ([object[]]$allRows[0]).Count

      # Flatten a cell's line array to a single GFM-safe string.
      # Blank-line-separated groups become <br>-separated segments;
      # lines within a group are also joined with <br>.
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
        $flat = $flat.Replace('|', '\|')   # escape pipes for GFM
        if ([string]::IsNullOrWhiteSpace($flat)) { return ' ' }
        return $flat.Trim()
      }

      # Build flat strings for every cell in every row.
      $flatRows = foreach ($row in $allRows) {
        $cells = foreach ($cellArr in [object[]]$row) {
          & $flattenCell ([string[]]$cellArr)
        }
        ,@($cells)
      }

      # Compute per-column widths (minimum 3).
      $colWidths = @(3) * $colCount
      foreach ($row in $flatRows) {
        $rowArr = [string[]]$row
        for ($c = 0; $c -lt [Math]::Min($colCount, $rowArr.Length); $c++) {
          $w = $rowArr[$c].Length
          if ($w -gt $colWidths[$c]) { $colWidths[$c] = $w }
        }
      }

      # Render header (first row).
      $hArr   = [string[]]$flatRows[0]
      $hParts = for ($c = 0; $c -lt $colCount; $c++) {
        $v = if ($c -lt $hArr.Length) { $hArr[$c] } else { ' ' }
        $v.PadRight($colWidths[$c])
      }
      $outLines.Add('| ' + ($hParts -join ' | ') + ' |')

      # Separator row.
      $sepParts = $colWidths | ForEach-Object { '-' * $_ }
      $outLines.Add('| ' + ($sepParts -join ' | ') + ' |')

      # Body rows.
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

function Get-PreviousElementSibling {
  param([System.Xml.XmlNode]$Node)
  $cursor = $Node.PreviousSibling
  while ($cursor -and $cursor.NodeType -ne [System.Xml.XmlNodeType]::Element) {
    $cursor = $cursor.PreviousSibling
  }
  return $cursor
}

function Get-NextElementSibling {
  param([System.Xml.XmlNode]$Node)
  $cursor = $Node.NextSibling
  while ($cursor -and $cursor.NodeType -ne [System.Xml.XmlNodeType]::Element) {
    $cursor = $cursor.NextSibling
  }
  return $cursor
}

function Test-IsBlankParagraph {
  param(
    [System.Xml.XmlNode]$Node,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if (-not $Node -or $Node.LocalName -ne 'p') {
    return $false
  }

  if ($Node.SelectSingleNode('.//w:drawing', $NamespaceManager)) {
    return $false
  }

  $texts = $Node.SelectNodes('.//w:t', $NamespaceManager)
  if ($texts.Count -eq 0) {
    return $true
  }

  foreach ($t in $texts) {
    if (-not [string]::IsNullOrWhiteSpace($t.InnerText)) {
      return $false
    }
  }

  return $true
}

function New-EmptyWordParagraph {
  param(
    [xml]$XmlDocument,
    [string]$NamespaceUri
  )

  return $XmlDocument.CreateElement('w', 'p', $NamespaceUri)
}

function Ensure-BlankParagraphAroundNode {
  param(
    [System.Xml.XmlNode]$Node,
    [xml]$XmlDocument,
    [System.Xml.XmlNamespaceManager]$NamespaceManager,
    [string]$NamespaceUri,
    [ref]$Modified
  )

  $prev = Get-PreviousElementSibling -Node $Node
  if (-not (Test-IsBlankParagraph -Node $prev -NamespaceManager $NamespaceManager)) {
    $blankBefore = New-EmptyWordParagraph -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri
    $Node.ParentNode.InsertBefore($blankBefore, $Node) | Out-Null
    $Modified.Value = $true
  }

  $next = Get-NextElementSibling -Node $Node
  if (-not (Test-IsBlankParagraph -Node $next -NamespaceManager $NamespaceManager)) {
    $blankAfter = New-EmptyWordParagraph -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri
    if ($Node.NextSibling) {
      $Node.ParentNode.InsertAfter($blankAfter, $Node) | Out-Null
    } else {
      $Node.ParentNode.AppendChild($blankAfter) | Out-Null
    }
    $Modified.Value = $true
  }
}

function Test-IsWordBookmarkName {
  param([string]$Name)

  if ([string]::IsNullOrWhiteSpace($Name)) {
    return $false
  }

  return $Name -match '^[A-Za-z_][A-Za-z0-9_]{0,39}$'
}

function Get-PandocBookmarkName {
  param([string]$Identifier)

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Identifier)
  $sha1 = [System.Security.Cryptography.SHA1]::Create()
  try {
    $hash = $sha1.ComputeHash($bytes)
  }
  finally {
    $sha1.Dispose()
  }

  $hex = (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
  if ($hex.Length -gt 1) {
    $hex = $hex.Substring(1)
  }
  return "X" + $hex
}

function Get-MarkdownBookmarkRenameMap {
  param([string]$Markdown)

  $map = @{}
  $ids = New-Object 'System.Collections.Generic.HashSet[string]'

  foreach ($m in [regex]::Matches($Markdown, '\{#([^\s\}]+)\}')) {
    $null = $ids.Add($m.Groups[1].Value)
  }

  foreach ($m in [regex]::Matches($Markdown, '\]\(#([^\s\)]+)')) {
    $id = $m.Groups[1].Value.Trim()
    if ($id.StartsWith("<") -and $id.EndsWith(">")) {
      $id = $id.Substring(1, $id.Length - 2)
    }
    if ($id -match '^(?i)(https?|ftp|ftps|mailto|tel|data|file):') {
      continue
    }
    if ($id.StartsWith("/")) {
      continue
    }
    $null = $ids.Add($id)
  }

  foreach ($id in $ids) {
    if (-not (Test-IsWordBookmarkName -Name $id)) {
      continue
    }
    $hashed = Get-PandocBookmarkName -Identifier $id
    $map[$hashed] = $id
  }

  return $map
}

function Normalize-RefFieldSwitches {
  param([string]$RawSwitches)

  $tokens = [regex]::Matches($RawSwitches, '\\[A-Za-z]+') | ForEach-Object { $_.Value }
  if ($tokens.Count -eq 0) {
    return '\h \n'
  }

  return ($tokens -join ' ')
}

function New-WordFieldControlRun {
  param(
    [xml]$XmlDocument,
    [string]$NamespaceUri,
    [string]$FieldType
  )

  $run = $XmlDocument.CreateElement('w', 'r', $NamespaceUri)
  $fldChar = $XmlDocument.CreateElement('w', 'fldChar', $NamespaceUri)
  $null = $fldChar.SetAttribute('fldCharType', $NamespaceUri, $FieldType)
  $run.AppendChild($fldChar) | Out-Null
  return $run
}

function New-WordInstructionRun {
  param(
    [xml]$XmlDocument,
    [string]$NamespaceUri,
    [string]$Instruction
  )

  $run = $XmlDocument.CreateElement('w', 'r', $NamespaceUri)
  $instr = $XmlDocument.CreateElement('w', 'instrText', $NamespaceUri)
  $null = $instr.SetAttribute('xml:space', 'preserve')
  $instr.InnerText = $Instruction
  $run.AppendChild($instr) | Out-Null
  return $run
}

function New-WordTextRun {
  param(
    [xml]$XmlDocument,
    [string]$NamespaceUri,
    [string]$Text
  )

  $run = $XmlDocument.CreateElement('w', 'r', $NamespaceUri)
  $textNode = $XmlDocument.CreateElement('w', 't', $NamespaceUri)
  if ($Text.StartsWith(" ") -or $Text.EndsWith(" ") -or $Text.Contains("  ")) {
    $null = $textNode.SetAttribute('xml:space', 'preserve')
  }
  $textNode.InnerText = $Text
  $run.AppendChild($textNode) | Out-Null
  return $run
}

function Test-RunHasVisibleContent {
  param(
    [System.Xml.XmlNode]$RunNode,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  if ($RunNode.SelectSingleNode('.//w:drawing', $NamespaceManager)) {
    return $true
  }
  if ($RunNode.SelectSingleNode('.//w:tab', $NamespaceManager)) {
    return $true
  }
  if ($RunNode.SelectSingleNode('.//w:br', $NamespaceManager)) {
    return $true
  }

  foreach ($t in $RunNode.SelectNodes('.//w:t', $NamespaceManager)) {
    if (-not [string]::IsNullOrEmpty($t.InnerText)) {
      return $true
    }
  }

  return $false
}

function Get-TextFromRuns {
  param(
    [System.Collections.Generic.List[System.Xml.XmlNode]]$Runs,
    [System.Xml.XmlNamespaceManager]$NamespaceManager
  )

  $parts = New-Object 'System.Collections.Generic.List[string]'
  foreach ($run in $Runs) {
    foreach ($t in @($run.SelectNodes('.//w:t', $NamespaceManager))) {
      $parts.Add($t.InnerText)
    }
  }
  return ($parts -join '')
}

function Try-ExtractMarkdownLinkLabel {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $null
  }

  $trimmed = $Text.Trim()
  if ($trimmed -match '^\[(?<label>.+)\]\(#.+\)$') {
    return $Matches['label']
  }

  return $null
}

function Rename-BookmarksAndConvertXrefLinks {
  param(
    [xml]$XmlDocument,
    [System.Xml.XmlNamespaceManager]$NamespaceManager,
    [string]$NamespaceUri,
    [hashtable]$BookmarkRenameMap
  )

  $modified = $false

  foreach ($bm in $XmlDocument.SelectNodes('//w:bookmarkStart[@w:name]', $NamespaceManager)) {
    $name = $bm.GetAttribute('name', $NamespaceUri)
    if ($BookmarkRenameMap.ContainsKey($name)) {
      $null = $bm.SetAttribute('name', $NamespaceUri, [string]$BookmarkRenameMap[$name])
      $modified = $true
    }
  }

  foreach ($link in $XmlDocument.SelectNodes('//w:hyperlink[@w:anchor]', $NamespaceManager)) {
    $anchor = $link.GetAttribute('anchor', $NamespaceUri)
    if ($BookmarkRenameMap.ContainsKey($anchor)) {
      $null = $link.SetAttribute('anchor', $NamespaceUri, [string]$BookmarkRenameMap[$anchor])
      $modified = $true
    }
  }

  $bookmarkNames = @{}
  foreach ($bm in $XmlDocument.SelectNodes('//w:bookmarkStart[@w:name]', $NamespaceManager)) {
    $bookmarkNames[$bm.GetAttribute('name', $NamespaceUri)] = $true
  }

  $hyperlinks = @($XmlDocument.SelectNodes('//w:hyperlink[@w:anchor]', $NamespaceManager))
  foreach ($hyperlink in $hyperlinks) {
    $anchor = $hyperlink.GetAttribute('anchor', $NamespaceUri)
    if ([string]::IsNullOrWhiteSpace($anchor)) {
      continue
    }

    $textNodes = @($hyperlink.SelectNodes('.//w:t', $NamespaceManager))
    if ($textNodes.Count -eq 0) {
      continue
    }

    $fullText = ($textNodes | ForEach-Object { $_.InnerText }) -join ''
    if ($fullText -notmatch '^\[\[XRF:(?<switches>[^\]]*)\]\]') {
      continue
    }

    $switchesRaw = $Matches['switches']
    $markerText = "[[XRF:$switchesRaw]]"

    $displayRuns = New-Object 'System.Collections.Generic.List[System.Xml.XmlNode]'
    foreach ($child in @($hyperlink.ChildNodes)) {
      if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element -and $child.LocalName -eq 'r') {
        $displayRuns.Add($child.CloneNode($true))
      }
    }

    $remaining = $markerText.Length
    foreach ($run in $displayRuns) {
      foreach ($t in @($run.SelectNodes('.//w:t', $NamespaceManager))) {
        if ($remaining -le 0) { break }
        $text = $t.InnerText
        if ($text.Length -le $remaining) {
          $t.InnerText = ''
          $remaining -= $text.Length
        } else {
          $t.InnerText = $text.Substring($remaining)
          $remaining = 0
        }
      }
      if ($remaining -le 0) { break }
    }

    $cleanRuns = New-Object 'System.Collections.Generic.List[System.Xml.XmlNode]'
    foreach ($run in $displayRuns) {
      if (Test-RunHasVisibleContent -RunNode $run -NamespaceManager $NamespaceManager) {
        $cleanRuns.Add($run)
      }
    }

    if ($cleanRuns.Count -eq 0) {
      $fallback = $fullText.Substring([Math]::Min($markerText.Length, $fullText.Length))
      if ([string]::IsNullOrWhiteSpace($fallback)) {
        $fallback = $anchor
      }
      $cleanRuns.Add((New-WordTextRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -Text $fallback))
    }

    $displayText = Get-TextFromRuns -Runs $cleanRuns -NamespaceManager $NamespaceManager
    $label = Try-ExtractMarkdownLinkLabel -Text $displayText
    if (-not [string]::IsNullOrWhiteSpace($label)) {
      $cleanRuns.Clear()
      $cleanRuns.Add((New-WordTextRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -Text $label))
    }

    $replacementNodes = New-Object 'System.Collections.Generic.List[System.Xml.XmlNode]'
    if ($bookmarkNames.ContainsKey($anchor)) {
      $switches = Normalize-RefFieldSwitches -RawSwitches $switchesRaw
      $instruction = " REF $anchor"
      if (-not [string]::IsNullOrWhiteSpace($switches)) {
        $instruction += " $switches"
      }
      $instruction += " "

      $replacementNodes.Add((New-WordFieldControlRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -FieldType 'begin'))
      $replacementNodes.Add((New-WordInstructionRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -Instruction $instruction))
      $replacementNodes.Add((New-WordFieldControlRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -FieldType 'separate'))
      foreach ($run in $cleanRuns) {
        $replacementNodes.Add($run)
      }
      $replacementNodes.Add((New-WordFieldControlRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -FieldType 'end'))
    } else {
      $replacementNodes.Add((New-WordTextRun -XmlDocument $XmlDocument -NamespaceUri $NamespaceUri -Text "[BROKEN_XREF:#$anchor]"))
    }

    $parent = $hyperlink.ParentNode
    foreach ($node in $replacementNodes) {
      $parent.InsertBefore($node, $hyperlink) | Out-Null
    }
    $parent.RemoveChild($hyperlink) | Out-Null
    $modified = $true
  }

  return $modified
}

function Format-DocxLayout {
  param(
    [string]$DocxPath,
    [string]$SourceMarkdown
  )

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $nsUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

  $zip = [System.IO.Compression.ZipFile]::Open(
    $DocxPath, [System.IO.Compression.ZipArchiveMode]::Update)

  try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $content = $reader.ReadToEnd()
    $reader.Dispose()

    $xml = [xml]$content
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $nsUri)

    $modified = $false
    $bookmarkRenameMap = Get-MarkdownBookmarkRenameMap -Markdown $SourceMarkdown
    if (Rename-BookmarksAndConvertXrefLinks -XmlDocument $xml -NamespaceManager $ns -NamespaceUri $nsUri -BookmarkRenameMap $bookmarkRenameMap) {
      $modified = $true
    }

    foreach ($tbl in $xml.SelectNodes('//w:tbl', $ns)) {
      # --- Table borders ---
      $tblPr = $tbl.SelectSingleNode('w:tblPr', $ns)
      if (-not $tblPr) {
        $tblPr = $xml.CreateElement('w', 'tblPr', $nsUri)
        $tbl.PrependChild($tblPr) | Out-Null
        $modified = $true
      }

      $old = $tblPr.SelectSingleNode('w:tblBorders', $ns)
      if ($old) {
        $tblPr.RemoveChild($old) | Out-Null
        $modified = $true
      }

      $borders = $xml.CreateElement('w', 'tblBorders', $nsUri)
      foreach ($side in @('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) {
        $b = $xml.CreateElement('w', $side, $nsUri)
        $null = $b.SetAttribute('val', $nsUri, 'single')
        $null = $b.SetAttribute('sz', $nsUri, '4')
        $null = $b.SetAttribute('space', $nsUri, '0')
        $null = $b.SetAttribute('color', $nsUri, 'BFBFBF')
        $borders.AppendChild($b) | Out-Null
      }
      $tblPr.AppendChild($borders) | Out-Null
      $modified = $true

      # --- Center table ---
      $tblJc = $tblPr.SelectSingleNode('w:jc', $ns)
      if (-not $tblJc) {
        $tblJc = $xml.CreateElement('w', 'jc', $nsUri)
        $tblPr.AppendChild($tblJc) | Out-Null
        $modified = $true
      }
      if ($tblJc.GetAttribute('val', $nsUri) -ne 'center') {
        $null = $tblJc.SetAttribute('val', $nsUri, 'center')
        $modified = $true
      }

      # --- Header row shading ---
      $firstRow = $tbl.SelectSingleNode('w:tr', $ns)
      if (-not $firstRow) { continue }

      foreach ($tc in $firstRow.SelectNodes('w:tc', $ns)) {
        $tcPr = $tc.SelectSingleNode('w:tcPr', $ns)
        if (-not $tcPr) {
          $tcPr = $xml.CreateElement('w', 'tcPr', $nsUri)
          $tc.PrependChild($tcPr) | Out-Null
          $modified = $true
        }

        $oldShd = $tcPr.SelectSingleNode('w:shd', $ns)
        if ($oldShd) {
          $tcPr.RemoveChild($oldShd) | Out-Null
          $modified = $true
        }

        $shd = $xml.CreateElement('w', 'shd', $nsUri)
        $null = $shd.SetAttribute('val', $nsUri, 'clear')
        $null = $shd.SetAttribute('color', $nsUri, 'auto')
        $null = $shd.SetAttribute('fill', $nsUri, 'F2F2F2')
        $tcPr.AppendChild($shd) | Out-Null
        $modified = $true
      }
    }

    # Center all image paragraphs.
    foreach ($imgPara in $xml.SelectNodes('//w:p[.//w:drawing]', $ns)) {
      $pPr = $imgPara.SelectSingleNode('w:pPr', $ns)
      if (-not $pPr) {
        $pPr = $xml.CreateElement('w', 'pPr', $nsUri)
        $imgPara.PrependChild($pPr) | Out-Null
        $modified = $true
      }

      $jc = $pPr.SelectSingleNode('w:jc', $ns)
      if (-not $jc) {
        $jc = $xml.CreateElement('w', 'jc', $nsUri)
        $pPr.AppendChild($jc) | Out-Null
        $modified = $true
      }

      if ($jc.GetAttribute('val', $nsUri) -ne 'center') {
        $null = $jc.SetAttribute('val', $nsUri, 'center')
        $modified = $true
      }
    }

    # Ensure visible blank lines before and after top-level tables and illustrations.
    $bodyTargets = @(
      $xml.SelectNodes('/w:document/w:body/*[self::w:tbl or (self::w:p and .//w:drawing)]', $ns)
    )
    foreach ($target in $bodyTargets) {
      Ensure-BlankParagraphAroundNode `
        -Node $target `
        -XmlDocument $xml `
        -NamespaceManager $ns `
        -NamespaceUri $nsUri `
        -Modified ([ref]$modified)
    }

    if ($modified) {
      $entry.Delete()
      $newEntry = $zip.CreateEntry('word/document.xml',
        [System.IO.Compression.CompressionLevel]::Optimal)
      $ws = $newEntry.Open()
      $xml.Save($ws)
      $ws.Close()
    }
  }
  finally {
    $zip.Dispose()
  }
}

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

if (-not (Test-Path -LiteralPath $referenceDoc)) {
  throw "Missing reference doc: $referenceDoc"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension

$preparedMarkdown = Convert-OffsetPrefixesToStyleSuffix -InputPath $resolvedInput
$preparedMarkdown = Convert-LargeTableToGfm -Markdown $preparedMarkdown
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("m2d_" + [System.Guid]::NewGuid().ToString("N"))
$tempInput = Join-Path $tempRoot "input.md"
$preparedMarkdown = Stage-MarkdownImages -Markdown $preparedMarkdown -MarkdownPath $resolvedInput -TempRoot $tempRoot
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
[System.IO.File]::WriteAllText($tempInput, $preparedMarkdown, $utf8NoBom)

try {
  & pandoc `
    -f "markdown+fancy_lists+lists_without_preceding_blankline+fenced_divs-smart" `
    -t "docx" `
    --no-highlight `
    --resource-path="$tempRoot" `
    --reference-doc="$referenceDoc" `
    --lua-filter="$filterPath" `
    "$tempInput" `
    -o "$resolvedOutput"
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Force -Recurse
  }
}

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

Format-DocxLayout -DocxPath $resolvedOutput -SourceMarkdown $preparedMarkdown

Write-Output "Created: $resolvedOutput"
