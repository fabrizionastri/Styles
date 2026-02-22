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

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputDir = Split-Path -Parent $InputPath
    $inputBase = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    return (Join-Path $inputDir ($inputBase + $DefaultExtension))
  }

  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($OutputPath))) {
    return ($OutputPath + $DefaultExtension)
  }

  return $OutputPath
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

if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  throw "Pandoc is not installed or not in PATH."
}

if (-not (Test-Path -LiteralPath $filterPath)) {
  throw "Missing filter: $filterPath"
}

$resolvedInput = Resolve-InputPath -PathValue $InputFile -DefaultExtension $defaultInputExtension
$resolvedOutput = Resolve-OutputPath -InputPath $resolvedInput -OutputPath $OutputFile -DefaultExtension $defaultOutputExtension

$tempOutput = Join-Path ([System.IO.Path]::GetTempPath()) ("d2m_" + [System.Guid]::NewGuid().ToString("N") + ".md")

try {
  $xrefData = Extract-CrossReferences -DocxPath $resolvedInput

  & pandoc `
    -f "docx+styles" `
    -t "markdown+fenced_divs" `
    --wrap=none `
    --lua-filter="$filterPath" `
    "$resolvedInput" `
    -o "$tempOutput"

  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  $rendered = [System.IO.File]::ReadAllText($tempOutput, $utf8NoBom)
  $rendered = Add-CrossReferences -Markdown $rendered -XRefData $xrefData
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($resolvedOutput, $rendered, $utf8Bom)
}
finally {
  if (Test-Path -LiteralPath $tempOutput) {
    Remove-Item -LiteralPath $tempOutput -Force
  }
}

Write-Output "Created: $resolvedOutput"
