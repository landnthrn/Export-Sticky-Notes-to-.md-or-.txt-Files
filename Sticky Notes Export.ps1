# Close Sticky Notes so DB isn't locked
taskkill /IM Microsoft.Notes.exe /F 2>$null | Out-Null

# --- Colors & helpers ---
$ColorQuestion = "Green"   # bright green
$ColorOption   = "White"   # bright white
$ColorWarn     = "Yellow"

function Show-Question($lines) { foreach ($ln in $lines) { Write-Host $ln -ForegroundColor $ColorQuestion } }
function Show-Option($lines)   { foreach ($ln in $lines) { Write-Host $ln -ForegroundColor $ColorOption } }
function Show-Warn($text)      { Write-Host $text -ForegroundColor $ColorWarn }

# --- Paths ---
$plum = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
if (!(Test-Path $plum)) { throw "Couldn't find plum.sqlite at: $plum" }

Show-Question @("Launching The Powershell Feature in CMD...")
Show-Option @("")

# Output folder (loop until valid)
do {
  Write-Host -ForegroundColor $ColorQuestion -NoNewline "Paste a path where you want the folder with the exports to be created: "
  $vault = Read-Host
  if ([string]::IsNullOrWhiteSpace($vault)) { Show-Warn "Please enter a path."; continue }
  if (!(Test-Path $vault)) { Show-Warn "Vault path not found. Try again."; continue }
  Show-Option @("")
  break
} while ($true)
$outDir = Join-Path $vault "Sticky Notes Export"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# --- User preferences ---
$includeDate = $false
$dateInFilename = $false
$fileExtension = ".md"
$includeMarkdown = $false

# Ask about including last edited date/time (loop)
do {
  Show-Option @("")
Show-Question @("Do you want to add time and date of when the notes were last edited?")
Show-Option @("Y - Yes","S - Skip","")
Write-Host -ForegroundColor $ColorOption "Enter Command: " -NoNewline
$dateChoice = Read-Host
  if ($dateChoice -eq "Y" -or $dateChoice -eq "y") {
    $includeDate = $true
    break
  } elseif ($dateChoice -eq "S" -or $dateChoice -eq "s") {
    $includeDate = $false
    break
  } else {
    Show-Warn "Invalid entry. Please choose Y or S."
  }
  Show-Option @("")
} while ($true)
$null = Show-Option @("")

# Ask where to place date if enabled (loop)
if ($includeDate) {
  do {
    Show-Option @("")
    Show-Question @("Do you want the last edited time and date in the filename or at the top of the file?")
    Show-Option @("1 - Filename","2 - Top of File Content","")
    Write-Host -ForegroundColor $ColorOption "Enter Command: " -NoNewline
    $dateLocation = Read-Host
    if ($dateLocation -eq "1") {
      $dateInFilename = $true
      break
    } elseif ($dateLocation -eq "2") {
      $dateInFilename = $false
      break
    } else {
      Show-Warn "Invalid entry. Please choose 1 or 2."
    }
    Show-Option @("")
  } while ($true)
  $null = Show-Option @("")
}

# Ask about file format (loop)
do {
  Show-Option @("")
  Show-Question @("Do you want files to export in .md or .txt?")
  Show-Option @("1 - .md","2 - .txt","")
  Write-Host -ForegroundColor $ColorOption "Enter Command: " -NoNewline
  $formatChoice = Read-Host
  if ($formatChoice -eq "1") {
    $fileExtension = ".md"
    break
  } elseif ($formatChoice -eq "2") {
    $fileExtension = ".txt"
    break
  } else {
    Show-Warn "Invalid entry. Please choose 1 or 2."
  }
  Show-Option @("")
} while ($true)
$null = Show-Option @("")

# Ask about markdown formatting (loop)
do {
  Show-Option @("")
  Show-Question @("Include markdown format like **bold**, *italic*, _underline_, ~~strikethrough~~, and so on?")
  Show-Option @("Y - Yes","N - No","")
  Write-Host -ForegroundColor $ColorOption "Enter Command: " -NoNewline
  $markdownChoice = Read-Host
  if ($markdownChoice -eq "Y" -or $markdownChoice -eq "y") {
    $includeMarkdown = $true
    break
  } elseif ($markdownChoice -eq "N" -or $markdownChoice -eq "n") {
    $includeMarkdown = $false
    break
  } else {
    Show-Warn "Invalid entry. Please choose Y or N."
  }
  Show-Option @("")
} while ($true)
$null = Show-Option @("")

# sqlite3 path
$sqlite = "C:\CleanPaths\sqlite\sqlite3.exe"
if (!(Test-Path $sqlite)) { throw "sqlite3 not found at $sqlite" }

# temp CSV paths
$tmpCsvNoBom   = "C:\CleanPaths\sqlite\sticky_export_utf8.csv"
$tmpCsvWithBom = "C:\CleanPaths\sqlite\sticky_export_utf8_bom.csv"

function Run-Sqlite([string]$sql) { & $sqlite "`"$plum`"" $sql }
function Run-SqliteScalar([string]$sql) { (Run-Sqlite $sql) -join "`n" }
function Get-Col($row,[string]$name){$p=$row.PSObject.Properties[$name];if($p){[string]$p.Value}else{""}}

# Detect table + columns
$table = Run-SqliteScalar "SELECT name FROM sqlite_master WHERE type='table' AND LOWER(name) IN ('note','notes');"
if (-not $table) { throw "Couldn't locate Sticky Notes table (tried 'Note' and 'Notes')." }

$colsCsv = Run-SqliteScalar "SELECT group_concat(name, ',') FROM pragma_table_info('$table');"
if (-not $colsCsv) { throw "Couldn't read columns from table $table." }
$cols = $colsCsv.Split(',') | ForEach-Object { $_.Trim() }

# Check for formatting/annotation tables that might contain UUID-to-format mappings
# This is exploratory - structure may vary by Sticky Notes version
$formatTable = $null
$allTables = Run-SqliteScalar "SELECT name FROM sqlite_master WHERE type='table';"
$allTablesArray = $allTables -split "`n" | Where-Object { $_ -and $_.Trim() }
foreach ($tbl in $allTablesArray) {
  $tblLower = $tbl.ToLower()
  if ($tblLower -like '*format*' -or $tblLower -like '*style*' -or $tblLower -like '*annotation*' -or $tblLower -like '*run*' -or $tblLower -like '*segment*') {
    $formatTable = $tbl.Trim()
    break
  }
}

# Export these if present (order matters)
$desired = @('Id','Title','Text','Content','RichText','Body','DateCreated','DateModified','CreatedAt','UpdatedAt')
$have = @(); foreach ($c in $desired) { if ($cols -contains $c) { $have += $c } }
if ($have.Count -eq 0) { throw "Table $table has none of the expected columns." }

# --- Ask sqlite to write the CSV (UTF-8, no BOM) directly to file ---
$selectList = ($have -join ', ')
$script = @"
.headers on
.mode csv
.output $tmpCsvNoBom
SELECT $selectList FROM $table;
.output stdout
"@

$scriptPath = "C:\CleanPaths\sqlite\sticky_export.sql"
[System.IO.File]::WriteAllText($scriptPath, $script, (New-Object System.Text.UTF8Encoding($false)))

& $sqlite "`"$plum`"" ".read `"$scriptPath`""

if (!(Test-Path $tmpCsvNoBom)) { throw "Export failed: CSV not created at $tmpCsvNoBom" }

# --- Re-save with UTF-8 BOM so Import-Csv (PS 5.1) reads as UTF-8 ---
[byte[]]$bom = 0xEF,0xBB,0xBF
$bytes = [System.IO.File]::ReadAllBytes($tmpCsvNoBom)
$fs = [System.IO.File]::Open($tmpCsvWithBom, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
$fs.Write($bom, 0, $bom.Length)
$fs.Write($bytes, 0, $bytes.Length)
$fs.Close()

# --- Helpers ---
function Sanitize-Name([string]$s, [bool]$allowLonger = $false){
  if ($null -eq $s) { $s = "" }
  $s = $s -replace '(?<!\\)\\b0?(?=\b|[\s\p{P}]|$)', ''     # strip \b / \b0
  $t = ($s -replace '[\\/:*?"<>|]', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($t)) { $t = "Note" }
  # Only apply length limit if not allowing longer (e.g., when date is included)
  if (-not $allowLonger -and $t.Length -gt 60) { $t = $t.Substring(0,60) }
  return $t
}

function Convert-NumericTimestampToLocal([string]$value) {
  if ([string]::IsNullOrEmpty($value)) { return $null }
  try {
    $num = [double]::Parse($value)
  } catch {
    return $null
  }

  # Heuristics:
  # > 1e14 : treat as .NET ticks since 0001-01-01 (100ns)
  # > 1e11 : treat as Unix ms
  # > 1e9  : treat as Unix seconds
  if ($num -gt 1e14) {
    try {
      $dt = [DateTime]::SpecifyKind([DateTime]::MinValue.AddTicks([long]$num), [DateTimeKind]::Utc)
      return $dt.ToLocalTime()
    } catch {}
  }
  if ($num -gt 1e11) {
    try { return [DateTimeOffset]::FromUnixTimeMilliseconds([long]$num).ToLocalTime().DateTime } catch {}
  }
  if ($num -gt 1e9) {
    try { return [DateTimeOffset]::FromUnixTimeSeconds([long]$num).ToLocalTime().DateTime } catch {}
  }

  # Fallback: try parse as DateTime
  try { return [DateTime]::Parse($value).ToLocalTime() } catch {}
  return $null
}

function Format-DateForFile([DateTime]$dt, [bool]$forFilename) {
  # Filename: LE H'mmAM MM.dd.yyyy
  # Content:  Last Edited H'mmAM MM.dd.yyyy
  $hour = $dt.Hour
  $minute = $dt.Minute.ToString("00")
  $ampm = "AM"
  if ($hour -eq 0) {
    $hour = 12
  } elseif ($hour -eq 12) {
    $ampm = "PM"
  } elseif ($hour -gt 12) {
    $hour = $hour - 12
    $ampm = "PM"
  }
  $hourStr = $hour.ToString()
  
  $month = $dt.Month.ToString("00")
  $day = $dt.Day.ToString("00")
  $year = $dt.Year.ToString()
  
  if ($forFilename) {
    return "LE $hourStr'$minute$ampm $month.$day.$year"
  } else {
    return "Last Edited $hourStr'$minute$ampm $month.$day.$year"
  }
}

function Convert-StickyNotesUuidFormat([string]$rawText) {
  if ([string]::IsNullOrEmpty($rawText)) { return $rawText }

  $text = $rawText

  # Check if this uses UUID markers (Sticky Notes format) instead of RTF
  $uuidPattern = '=[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}'
  $hasUuidMarkers = $text -match $uuidPattern

  if ($hasUuidMarkers) {
    # Simple approach: split by = and keep only non-UUID parts
    $parts = $text -split '='

    $result = ""
    for ($i = 0; $i -lt $parts.Length; $i++) {
      $part = $parts[$i].Trim()

      # If this part is not empty and not a UUID, keep it
      if (-not [string]::IsNullOrEmpty($part)) {
        # Check if it's a UUID (36 chars with hyphens in specific format)
        $isUuid = $part -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
        if (-not $isUuid) {
          if ($result.Length -gt 0) {
            $result += "`n"  # Add newline between text segments
          }
          $result += $part
        }
      }
    }

    if (-not [string]::IsNullOrEmpty($result)) {
      $text = $result
    } else {
      # Fallback: strip all = characters and UUIDs
      $text = $text -replace $uuidPattern, '' -replace '=', ''
    }
  }

  return $text
}

function Convert-RtfToPlain([string]$rtfText) {
  if ([string]::IsNullOrEmpty($rtfText)) { return $rtfText }

  $text = $rtfText

  # Normalize line markers first
  $text = $text -replace '\\par', "`r`n"
  $text = $text -replace '\\line', "`r`n"

  # Remove ids and common format toggles
  $text = $text -replace '\\id=[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', ''
  $text = $text -replace '\\b0', ''
  $text = $text -replace '\\b\b?', ''
  $text = $text -replace '\\i0', ''
  $text = $text -replace '\\i\b?', ''
  $text = $text -replace '\\ul0', ''
  $text = $text -replace '\\ul\b?', ''
  $text = $text -replace '\\strike0', ''
  $text = $text -replace '\\strike\b?', ''

  # Strip any remaining RTF control words and braces
  $text = $text -replace '\\[a-z]+\d*', ''
  $text = $text -replace '[{}]', ''

  # Trim line ends but preserve blank lines
  $lines = [regex]::Split($text, "\r?\n")
  $lines = $lines | ForEach-Object { $_.TrimEnd() }
  return ($lines -join "`r`n")
}

function Convert-RtfToMarkdown([string]$rtfText) {
  if ([string]::IsNullOrEmpty($rtfText)) { return $rtfText }

  $text = $rtfText

  # Check for RTF control codes
  $isRtf = $text -match '\\[a-z]+\d*|\{\s*\\[a-z]+|\\id=' -or $text.Contains('{\rtf') -or $text.Contains('\b') -or $text.Contains('\i')

  if ($isRtf) {
    # Remove RTF header if present
    $text = $text -replace '\\rtf1[^\\]*', ''

    # Split into lines (preserve empty lines) and process each line
    $lines = [regex]::Split($text, "\r?\n")
    $processedLines = @()
    $boldOpen = $false
    $italicOpen = $false
    $underlineOpen = $false
    $strikeOpen = $false

    foreach ($lineRaw in $lines) {
      $line = $lineRaw  # keep as-is to preserve blank lines

      # Remove \id=UUID parts
      $line = $line -replace '\\id=[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', ''

      # Replace RTF toggles even if closing marker is missing
      $line = $line -replace '\\b0', '**'
      $line = $line -replace '\\b\s*', '**'
      $line = $line -replace '\\i0', '*'
      $line = $line -replace '\\i\s*', '*'
      $line = $line -replace '\\ul0', '_'
      $line = $line -replace '\\ul\s*', '_'
      $line = $line -replace '\\strike0', '~~'
      $line = $line -replace '\\strike\s*', '~~'

      # Handle any remaining RTF control codes
      $line = $line -replace '\\[a-z]+\d*\s*', ''

      # Collapse multiple markdown delimiters (e.g., ****TEXT**** -> **TEXT**)
      $line = $line -replace '\*{3,}', '**'
      $line = $line -replace '~{3,}', '~~'
      $line = $line -replace '_{3,}', '_'

       # Prepend any open spans from previous line (only if line has content)
       if (-not [string]::IsNullOrEmpty($line)) {
         if ($boldOpen) { $line = '**' + $line }
         if ($italicOpen) { $line = '*' + $line }
         if ($underlineOpen) { $line = '_' + $line }
         if ($strikeOpen) { $line = '~~' + $line }
       }

       # Track opens/closes across lines; toggle if odd count on this line
       $boldCount = [regex]::Matches($line, '\*\*').Count
       if ($boldCount % 2 -eq 1) { $boldOpen = -not $boldOpen }
       $italicCount = [regex]::Matches($line, '\*').Count
       if ($italicCount % 2 -eq 1) { $italicOpen = -not $italicOpen }
       $underlineCount = [regex]::Matches($line, '_').Count
       if ($underlineCount % 2 -eq 1) { $underlineOpen = -not $underlineOpen }
       $strikeCount = [regex]::Matches($line, '~~').Count
       if ($strikeCount % 2 -eq 1) { $strikeOpen = -not $strikeOpen }

      # Preserve blank lines; add the line as-is
      $processedLines += $line
    }

    # Close any still-open spans at end
    if ($processedLines.Count -gt 0) {
      if ($strikeOpen)    { $processedLines[-1] = $processedLines[-1] + '~~' }
      if ($underlineOpen) { $processedLines[-1] = $processedLines[-1] + '_' }
      if ($italicOpen)    { $processedLines[-1] = $processedLines[-1] + '*' }
      if ($boldOpen)      { $processedLines[-1] = $processedLines[-1] + '**' }
    }

    $text = $processedLines -join "`r`n"
  }

  return $text
}

# --- Create one file per sticky, preserve all Unicode and formatting ---
$made = 0
Import-Csv -Path $tmpCsvWithBom | ForEach-Object {
  $id           = Get-Col $_ 'Id'
  $dateModified = Get-Col $_ 'DateModified'
  $dateCreated  = Get-Col $_ 'DateCreated'
  $createdAt    = Get-Col $_ 'CreatedAt'
  $updatedAt    = Get-Col $_ 'UpdatedAt'
  
  # Get both RichText and Text fields
  $richText = Get-Col $_ 'RichText'
  $plainText = Get-Col $_ 'Text'
  if ([string]::IsNullOrEmpty($richText) -and [string]::IsNullOrEmpty($plainText)) {
    $richText = Get-Col $_ 'Content'
    if ([string]::IsNullOrEmpty($richText)) { $richText = Get-Col $_ 'Body' }
  }
  if ([string]::IsNullOrEmpty($richText) -and [string]::IsNullOrEmpty($plainText)) { return }


  # Determine which field to use
  $raw = if (-not [string]::IsNullOrEmpty($richText)) { $richText } else { $plainText }

  if ($includeMarkdown) {
    try {
      $body = Convert-RtfToMarkdown $raw
    } catch {
      Write-Warning "Could not convert formatting for note $id, using plain text."
      $body = Convert-RtfToPlain $raw
    }
  } else {
    # Plain text: strip RTF/Sticky codes and leave unformatted text
    $body = Convert-RtfToPlain $raw
  }

  # Replace common RTF line markers with newlines before stripping other codes
  $body = $body -replace '\\par', "`r`n"
  $body = $body -replace '\\line', "`r`n"

  # Clean Sticky artifacts, but preserve blank lines
  $body = $body -replace '\\id=[0-9a-fA-F-]+',''                     # remove \id=...
  $body = $body -replace '(?<!\\)\\b0?(?=\b|[\s\p{P}]|$)', ''         # remove \b / \b0
  $body = [regex]::Replace($body, '(^|\r?\n)[ \t]+(?=\S)', '$1')      # remove leading spaces ONLY on non-empty lines
  # (no collapsing of blank lines, no Trim())

  # Format date/time if requested using UpdatedAt/CreatedAt (bigint) when present
  $dateStrForFilename = ""
  $dateStrForContent = ""
  if ($includeDate) {
    $dt = $null
    if (-not $dt) { $dt = Convert-NumericTimestampToLocal $updatedAt }
    if (-not $dt) { $dt = Convert-NumericTimestampToLocal $createdAt }
    if (-not $dt) { $dt = Convert-NumericTimestampToLocal $dateModified }
    if (-not $dt) { $dt = Convert-NumericTimestampToLocal $dateCreated }
    if (-not $dt) { $dt = [DateTime]::UtcNow }

    if ($dt) {
      $dateStrForFilename = Format-DateForFile $dt $true
      $dateStrForContent  = Format-DateForFile $dt $false
    }
  }

  # Title from first non-empty line of cleaned body; fallback to raw if needed
  $first = ($body -split '\r?\n' | Where-Object { $_.Trim() -ne "" } | Select-Object -First 1)
  if ([string]::IsNullOrWhiteSpace($first)) {
    $rawForTitle = if (-not [string]::IsNullOrEmpty($richText)) { $richText } else { $plainText }
    $first = ($rawForTitle -split '\r?\n' | Where-Object { $_.Trim() -ne "" } | Select-Object -First 1)
  }
  # Remove RTF codes and \id=UUID from title for filename
  $firstClean = $first -replace '\\id=[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', ''
  $firstClean = $firstClean -replace '\\[a-z]+\d*\s*', ''
  $firstClean = $firstClean.Trim()
  if ([string]::IsNullOrWhiteSpace($firstClean)) { $firstClean = "Note" }
  $name = Sanitize-Name $firstClean

  # Add date to filename if requested
  if ($includeDate -and $dateInFilename -and -not [string]::IsNullOrEmpty($dateStrForFilename)) {
    $name = "$name - $dateStrForFilename"
    $name = Sanitize-Name $name $true  # Re-sanitize but allow longer for date
  }

  # ensure unique filename
  $file = Join-Path $outDir ($name + $fileExtension)
  $i = 2
  while (Test-Path $file) { $file = Join-Path $outDir ("$name ($i)$fileExtension"); $i++ }

  # Add date at top of file if requested
  $fileContent = $body
  if ($includeDate -and -not $dateInFilename -and -not [string]::IsNullOrEmpty($dateStrForContent)) {
    $fileContent = "$dateStrForContent`r`n`r`n" + $fileContent
  }

  # Write as UTF-8 (no BOM)
  [System.IO.File]::WriteAllText($file, $fileContent + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
  $made++
}

# Cleanup temps
Remove-Item $tmpCsvNoBom,$tmpCsvWithBom,$scriptPath -ErrorAction SilentlyContinue

$extDisplay = if ($fileExtension -eq ".txt") { ".txt" } else { ".md" }
Write-Host "Done! All Sticky Notes are now exported to $extDisplay files in:" -ForegroundColor $ColorQuestion
Write-Host " $outDir" -ForegroundColor $ColorOption
Write-Host "" -ForegroundColor $ColorQuestion
Write-Host "Found this useful? Please give the GitHub post a star :)" -ForegroundColor $ColorQuestion
Write-Host "" -ForegroundColor $ColorQuestion
Write-Host "Check out more of my creations on GitHub:" -ForegroundColor $ColorQuestion
Write-Host "https://github.com/landnthrn?tab=repositories" -ForegroundColor $ColorQuestion
Write-Host "" -ForegroundColor $ColorQuestion
Write-Host "Press any key to continue . . ." -ForegroundColor $ColorQuestion
[void][System.Console]::ReadKey($true)

