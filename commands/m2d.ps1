[CmdletBinding()]
param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$InputFile,

  [Parameter(Position = 1)]
  [string]$OutputFile
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$filterPath = Join-Path $scriptRoot "..\filters\compact_to_docx.lua"
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

function Format-DocxLayout {
  param([string]$DocxPath)

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

Format-DocxLayout -DocxPath $resolvedOutput

Write-Output "Created: $resolvedOutput"
