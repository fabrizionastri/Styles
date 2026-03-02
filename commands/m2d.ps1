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

function Format-DocxTables {
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
      $modified = $true

      # --- Table borders ---
      $tblPr = $tbl.SelectSingleNode('w:tblPr', $ns)
      if (-not $tblPr) {
        $tblPr = $xml.CreateElement('w', 'tblPr', $nsUri)
        $tbl.PrependChild($tblPr) | Out-Null
      }

      $old = $tblPr.SelectSingleNode('w:tblBorders', $ns)
      if ($old) { $tblPr.RemoveChild($old) | Out-Null }

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

      # --- Header row shading ---
      $firstRow = $tbl.SelectSingleNode('w:tr', $ns)
      if (-not $firstRow) { continue }

      foreach ($tc in $firstRow.SelectNodes('w:tc', $ns)) {
        $tcPr = $tc.SelectSingleNode('w:tcPr', $ns)
        if (-not $tcPr) {
          $tcPr = $xml.CreateElement('w', 'tcPr', $nsUri)
          $tc.PrependChild($tcPr) | Out-Null
        }

        $oldShd = $tcPr.SelectSingleNode('w:shd', $ns)
        if ($oldShd) { $tcPr.RemoveChild($oldShd) | Out-Null }

        $shd = $xml.CreateElement('w', 'shd', $nsUri)
        $null = $shd.SetAttribute('val', $nsUri, 'clear')
        $null = $shd.SetAttribute('color', $nsUri, 'auto')
        $null = $shd.SetAttribute('fill', $nsUri, 'F2F2F2')
        $tcPr.AppendChild($shd) | Out-Null
      }
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
$tempInput = Join-Path ([System.IO.Path]::GetTempPath()) ("m2d_" + [System.Guid]::NewGuid().ToString("N") + ".md")
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($tempInput, $preparedMarkdown, $utf8NoBom)

try {
  & pandoc `
    -f "markdown+fancy_lists+lists_without_preceding_blankline+fenced_divs" `
    -t "docx" `
    --no-highlight `
    --reference-doc="$referenceDoc" `
    --lua-filter="$filterPath" `
    "$tempInput" `
    -o "$resolvedOutput"
}
finally {
  if (Test-Path -LiteralPath $tempInput) {
    Remove-Item -LiteralPath $tempInput -Force
  }
}

if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

Format-DocxTables -DocxPath $resolvedOutput

Write-Output "Created: $resolvedOutput"
