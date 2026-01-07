<#
EIP/ZIP Extraction GUI
Final Version
- ALIGNMENT: Precise vertical alignment between "Delete older than" and the "Clean" button.
- AUTO-FILL: Automatically populates Cleanup and Search folder paths after extraction.
- LAYOUT: Optimized window height (960px) for Search Results visibility.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------- LOGIC FUNCTIONS --------------------
function Expand-ZipRecursive {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [switch]$DeleteZip
    )
    if (-not (Test-Path -LiteralPath $DestinationPath)) { New-Item -ItemType Directory -Path $DestinationPath | Out-Null }
    Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force
    if ($DeleteZip) { Remove-Item -LiteralPath $ZipPath -Force }
    while ($true) {
        $zips = Get-ChildItem -Path $DestinationPath -Recurse -Filter *.zip -File -ErrorAction SilentlyContinue
        if (-not $zips) { break }
        foreach ($zip in $zips) {
            $dest = Join-Path -Path $zip.DirectoryName -ChildPath ([IO.Path]::GetFileNameWithoutExtension($zip.Name))
            if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
            Expand-Archive -Path $zip.FullName -DestinationPath $dest -Force
            Remove-Item -LiteralPath $zip.FullName -Force
        }
    }
}

function Get-EffectiveWorkingFolder {
    param([Parameter(Mandatory = $true)][string]$DestRoot)
    if (-not (Test-Path -LiteralPath $DestRoot)) { return $DestRoot }
    # Check if the folder contains a single subfolder (common in zips)
    $rootFiles = @(Get-ChildItem -LiteralPath $DestRoot -File -ErrorAction SilentlyContinue)
    $rootDirs  = @(Get-ChildItem -LiteralPath $DestRoot -Directory -ErrorAction SilentlyContinue)
    if ($rootFiles.Count -eq 0 -and $rootDirs.Count -eq 1) {
        return $rootDirs[0].FullName
    }
    return $DestRoot
}

function Open-FileWithEditor {
    param([Parameter(Mandatory = $true)][string]$Path)
    $npp64 = 'C:\Program Files\Notepad++\notepad++.exe'; $npp32 = 'C:\Program Files (x86)\Notepad++\notepad++.exe'; $editor = $null
    if (Test-Path -LiteralPath $npp64) { $editor = $npp64 } elseif (Test-Path -LiteralPath $npp32) { $editor = $npp32 }
    if ($editor) { Start-Process -FilePath $editor -ArgumentList @("`"$Path`"") | Out-Null } else { Start-Process -FilePath $Path | Out-Null }
}

function Clean-FolderKeepLogsTxt {
    param([string]$RootPath, [switch]$KeepLogsAndTxt, [switch]$PurgeByAge, [int]$MaxAgeDays)
    if (-not (Test-Path -LiteralPath $RootPath)) { throw "Folder not found: $RootPath" }
    $cutoff = $null; if ($PurgeByAge -and $MaxAgeDays -gt 0) { $cutoff = (Get-Date).AddDays(-1 * $MaxAgeDays) }
    if ($KeepLogsAndTxt -or $cutoff) {
        $files = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $remove = $false
            if ($KeepLogsAndTxt) { $ext = $file.Extension.ToLowerInvariant(); if ($ext -ne '.log' -and $ext -ne '.txt') { $remove = $true } }
            if (-not $remove -and $cutoff) { if ($file.LastWriteTime -lt $cutoff) { $remove = $true } }
            if ($remove) { Remove-Item -LiteralPath $file.FullName -Force }
        }
    }
    $dirs = Get-ChildItem -Path $RootPath -Recurse -Directory -ErrorAction SilentlyContinue | Sort-Object FullName -Descending
    foreach ($dir in $dirs) { $items = Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue; if (-not $items) { Remove-Item -LiteralPath $dir.FullName -Force } }
}

function Search-FilesForKeywords {
    param([string]$RootPath, [string]$Keyword1, [string]$Keyword2, [System.Windows.Forms.ListView]$ResultsList)
    $ResultsList.Items.Clear(); $ResultsList.BeginUpdate(); $filesMatched = 0; $totalMatches = 0
    try {
        $comparison = [System.StringComparison]::OrdinalIgnoreCase; $hasKeyword1 = -not [string]::IsNullOrWhiteSpace($Keyword1); $hasKeyword2 = -not [string]::IsNullOrWhiteSpace($Keyword2)
        $files = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try { $lines = Get-Content -Path $file.FullName -ReadCount 0 -ErrorAction Stop; if ($lines -is [string]) { $lines = @($lines) } } catch { continue }
            $matchCount = 0; $matchedLines = @()
            for ($i = 0; $i -lt $lines.Count; $i++) {
                $line = $lines[$i]; if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $match = $false
                if ($hasKeyword1 -and $hasKeyword2) { $match = ($line.IndexOf($Keyword1, $comparison) -ge 0 -and $line.IndexOf($Keyword2, $comparison) -ge 0) }
                elseif ($hasKeyword1) { $match = ($line.IndexOf($Keyword1, $comparison) -ge 0) }
                elseif ($hasKeyword2) { $match = ($line.IndexOf($Keyword2, $comparison) -ge 0) }
                if ($match) { $matchCount++; $matchedLines += ("{0}`t{1}" -f ($i + 1), $line) }
            }
            if ($matchCount -gt 0) {
                $item = New-Object System.Windows.Forms.ListViewItem($file.Name); [void]$item.SubItems.Add($matchCount.ToString())
                $item.Tag = [PSCustomObject]@{ Path = $file.FullName; Lines = $matchedLines }; $item.ToolTipText = $file.FullName
                [void]$ResultsList.Items.Add($item); $filesMatched++; $totalMatches += $matchCount
            }
        }
    } finally { $ResultsList.EndUpdate() }
    return @{ Files = $filesMatched; Matches = $totalMatches }
}

function Get-UniqueFolderPath {
    param([string]$BasePath)
    if (-not (Test-Path -LiteralPath $BasePath)) { return $BasePath }
    $parent = Split-Path -Path $BasePath -Parent; $leaf = Split-Path -Path $BasePath -Leaf
    for ($i = 1; $i -le 999; $i++) {
        $suffix = ('_{0:D2}' -f $i); $candidate = Join-Path -Path $parent -ChildPath ($leaf + $suffix)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw "Unable to find an available folder name for: $BasePath"
}

function Get-ProposedOutputFolder {
    $zipPath = $txtPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($zipPath)) { return '' }
    $root = $txtExtractTo.Text.Trim(); if ([string]::IsNullOrWhiteSpace($root)) { $root = 'C:\temp' }
    $name = [IO.Path]::GetFileNameWithoutExtension($zipPath)
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    return (Join-Path -Path $root -ChildPath $name)
}

function Test-FolderHasContent {
    param([string]$Path)
    try { $item = Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | Select-Object -First 1; return ($null -ne $item) } catch { return $false }
}

# -------------------- UI SETUP --------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'EIP/ZIP Extraction'
$form.Size = New-Object System.Drawing.Size(940, 960) 
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Bahnschrift', 9.5)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.AllowDrop = $true

$fontGroupTitle = New-Object System.Drawing.Font('Bahnschrift', 11, [System.Drawing.FontStyle]::Bold)
$toolTip = New-Object System.Windows.Forms.ToolTip

# Main Layout
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$mainLayout.RowCount = 3
$mainLayout.ColumnCount = 1
[void]$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 70)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)))

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.BackColor = [System.Drawing.Color]::FromArgb(0, 121, 107)
$header.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'EIP/ZIP Extraction'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Bahnschrift', 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(8, 6)
$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = 'Extract and clean folders'
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(230, 255, 255, 255)
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(10, 36)
$header.Controls.AddRange(@($lblTitle, $lblSub))

# Content Layout
$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.Dock = 'Fill'
$contentLayout.ColumnCount = 2
$contentLayout.RowCount = 2
$contentLayout.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 4)
[void]$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 290)))
[void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# Group Boxes
$grpExtract = New-Object System.Windows.Forms.GroupBox
$grpExtract.Text = '1) Extract'
$grpExtract.Font = $fontGroupTitle
$grpExtract.Dock = 'Fill'
$grpExtract.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)
$grpExtract.BackColor = [System.Drawing.Color]::White
$grpExtract.Margin = New-Object System.Windows.Forms.Padding(0, 0, 5, 0)

$grpClean = New-Object System.Windows.Forms.GroupBox
$grpClean.Text = '2) Cleanup'
$grpClean.Font = $fontGroupTitle
$grpClean.Dock = 'Fill'
$grpClean.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)
$grpClean.BackColor = [System.Drawing.Color]::White
$grpClean.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)

$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = '3) Search'
$grpSearch.Font = $fontGroupTitle
$grpSearch.Dock = 'Fill'
$grpSearch.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)
$grpSearch.BackColor = [System.Drawing.Color]::White
$grpSearch.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)

# ==================== 1. EXTRACT UI ====================
$extractContainer = New-Object System.Windows.Forms.TableLayoutPanel
$extractContainer.Dock = 'Fill'
$extractContainer.RowCount = 2
$extractContainer.ColumnCount = 1
$extractContainer.Font = $form.Font 
[void]$extractContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$extractContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$extractContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblExtractHelp = New-Object System.Windows.Forms.Label
$lblExtractHelp.Text = 'Select an archive, choose a destination, then Extract.'
$lblExtractHelp.AutoSize = $true
$lblExtractHelp.ForeColor = [System.Drawing.Color]::Gray

$extractLayout = New-Object System.Windows.Forms.TableLayoutPanel
$extractLayout.Dock = 'Fill'
$extractLayout.RowCount = 8
$extractLayout.ColumnCount = 1
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45)))
[void]$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) 
[void]$extractLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblArchive = New-Object System.Windows.Forms.Label
$lblArchive.Text = 'Archive file (.eip / .zip) (required)'
$lblArchive.AutoSize = $true

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Dock = 'Fill'
$txtPath.BorderStyle = 'FixedSingle'
$txtPath.Anchor = 'Left, Right'
$txtPath.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'

$pathLayout = New-Object System.Windows.Forms.TableLayoutPanel
$pathLayout.Dock = 'Fill'
$pathLayout.ColumnCount = 2
[void]$pathLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$pathLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$pathLayout.Controls.Add($txtPath, 0, 0)
[void]$pathLayout.Controls.Add($btnBrowse, 1, 0)

$lblArchiveTip = New-Object System.Windows.Forms.Label
$lblArchiveTip.Text = 'Tip: you can drag & drop a .zip/.eip file onto this window.'
$lblArchiveTip.AutoSize = $true
$lblArchiveTip.ForeColor = [System.Drawing.Color]::Gray

$lblExtractTo = New-Object System.Windows.Forms.Label
$lblExtractTo.Text = 'Destination folder (required)'
$lblExtractTo.AutoSize = $true

$txtExtractTo = New-Object System.Windows.Forms.TextBox
$txtExtractTo.Dock = 'Fill'
$txtExtractTo.BorderStyle = 'FixedSingle'
$txtExtractTo.Text = 'C:\temp'
$txtExtractTo.Anchor = 'Left, Right'
$txtExtractTo.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$btnBrowseExtractTo = New-Object System.Windows.Forms.Button
$btnBrowseExtractTo.Text = 'Browse...'

$extractToLayout = New-Object System.Windows.Forms.TableLayoutPanel
$extractToLayout.Dock = 'Fill'
$extractToLayout.ColumnCount = 2
[void]$extractToLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$extractToLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$extractToLayout.Controls.Add($txtExtractTo, 0, 0)
[void]$extractToLayout.Controls.Add($btnBrowseExtractTo, 1, 0)

$lblOutputFolder = New-Object System.Windows.Forms.Label
$lblOutputFolder.Dock = 'Fill'
$lblOutputFolder.AutoEllipsis = $true
$lblOutputFolder.ForeColor = [System.Drawing.Color]::Gray
$lblOutputFolder.Text = ''
$lblOutputFolder.Visible = $false

$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Text = 'Extract'
$btnExtract.Enabled = $false

$btnOpenOutput = New-Object System.Windows.Forms.Button
$btnOpenOutput.Text = 'Open output'
$btnOpenOutput.Enabled = $false

$btnCopyOutput = New-Object System.Windows.Forms.Button
$btnCopyOutput.Text = 'Copy path'
$btnCopyOutput.Enabled = $false

$extractActionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$extractActionLayout.Dock = 'Fill'
$extractActionLayout.ColumnCount = 7
$extractActionLayout.RowCount = 1
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$btnExtract.Anchor = 'None'
$btnOpenOutput.Anchor = 'None'
$btnCopyOutput.Anchor = 'None'
[void]$extractActionLayout.Controls.Add($btnExtract, 1, 0)
[void]$extractActionLayout.Controls.Add($btnOpenOutput, 3, 0)
[void]$extractActionLayout.Controls.Add($btnCopyOutput, 5, 0)

[void]$extractLayout.Controls.Add($lblArchive, 0, 0)
[void]$extractLayout.Controls.Add($pathLayout, 0, 1)
[void]$extractLayout.Controls.Add($lblArchiveTip, 0, 2)
[void]$extractLayout.Controls.Add($lblExtractTo, 0, 3)
[void]$extractLayout.Controls.Add($extractToLayout, 0, 4)
[void]$extractLayout.Controls.Add($lblOutputFolder, 0, 5)
[void]$extractLayout.Controls.Add($extractActionLayout, 0, 6)

[void]$extractContainer.Controls.Add($lblExtractHelp, 0, 0)
[void]$extractContainer.Controls.Add($extractLayout, 0, 1)
$grpExtract.Controls.Add($extractContainer)

# ==================== 2. CLEANUP UI ====================
$cleanContainer = New-Object System.Windows.Forms.TableLayoutPanel
$cleanContainer.Dock = 'Fill'
$cleanContainer.RowCount = 2
$cleanContainer.ColumnCount = 1
$cleanContainer.Font = $form.Font 
[void]$cleanContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$cleanContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$cleanContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblCleanHelp = New-Object System.Windows.Forms.Label
$lblCleanHelp.Text = 'Optional: keep only logs/text and/or purge by age (recursive).'
$lblCleanHelp.AutoSize = $true
$lblCleanHelp.ForeColor = [System.Drawing.Color]::Gray

$cleanLayout = New-Object System.Windows.Forms.TableLayoutPanel
$cleanLayout.Dock = 'Fill'
$cleanLayout.RowCount = 4
$cleanLayout.ColumnCount = 1
[void]$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
[void]$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))
[void]$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
[void]$cleanLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = 'Folder to clean (required)'
$lblFolder.AutoSize = $true

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Dock = 'Fill'
$txtFolder.BorderStyle = 'FixedSingle'
$txtFolder.Anchor = 'Left, Right'
$txtFolder.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$btnBrowseFolder = New-Object System.Windows.Forms.Button
$btnBrowseFolder.Text = 'Browse...'

$folderLayout = New-Object System.Windows.Forms.TableLayoutPanel
$folderLayout.Dock = 'Fill'
$folderLayout.ColumnCount = 2
[void]$folderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$folderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$folderLayout.Controls.Add($txtFolder, 0, 0)
[void]$folderLayout.Controls.Add($btnBrowseFolder, 1, 0)

$chkKeepLogs = New-Object System.Windows.Forms.CheckBox
$chkKeepLogs.Text = 'Keep only .log and .txt'
$chkKeepLogs.AutoSize = $true
$chkKeepLogs.Checked = $true

$chkPurgeAge = New-Object System.Windows.Forms.CheckBox
$chkPurgeAge.Text = 'Delete older than'
$chkPurgeAge.AutoSize = $true
$chkPurgeAge.Checked = $false
$chkPurgeAge.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0) # Margin handled by container

$nudDays = New-Object System.Windows.Forms.NumericUpDown
$nudDays.Minimum = 1
$nudDays.Maximum = 3650
$nudDays.Value = 30
$nudDays.Width = 60
$nudDays.Enabled = $false
$nudDays.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 0)

$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Text = 'days (opt.)'
$lblDays.AutoSize = $true
$lblDays.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 0)

$btnClean = New-Object System.Windows.Forms.Button
$btnClean.Text = 'Clean'
$btnClean.Anchor = 'Right'

# --- FIXED: ALIGNMENT ---
$ageRowLayout = New-Object System.Windows.Forms.TableLayoutPanel
$ageRowLayout.Dock = 'Fill'
$ageRowLayout.RowCount = 1
$ageRowLayout.ColumnCount = 2
[void]$ageRowLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$ageRowLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 135)))

$leftGroup = New-Object System.Windows.Forms.FlowLayoutPanel
$leftGroup.Dock = 'Fill'
$leftGroup.FlowDirection = 'LeftToRight'
$leftGroup.Controls.Add($chkPurgeAge)
$leftGroup.Controls.Add($nudDays)
$leftGroup.Controls.Add($lblDays)
# Added 7px top padding to perfectly align 20px text with 32px button
$leftGroup.Padding = New-Object System.Windows.Forms.Padding(0, 7, 0, 0) 
$leftGroup.Margin = New-Object System.Windows.Forms.Padding(0)

[void]$ageRowLayout.Controls.Add($leftGroup, 0, 0)
[void]$ageRowLayout.Controls.Add($btnClean, 1, 0)

[void]$cleanLayout.Controls.Add($lblFolder, 0, 0)
[void]$cleanLayout.Controls.Add($folderLayout, 0, 1)
[void]$cleanLayout.Controls.Add($chkKeepLogs, 0, 2)
[void]$cleanLayout.Controls.Add($ageRowLayout, 0, 3)

[void]$cleanContainer.Controls.Add($lblCleanHelp, 0, 0)
[void]$cleanContainer.Controls.Add($cleanLayout, 0, 1)
$grpClean.Controls.Add($cleanContainer)

# ==================== 3. SEARCH UI ====================
$searchContainer = New-Object System.Windows.Forms.TableLayoutPanel
$searchContainer.Dock = 'Fill'
$searchContainer.RowCount = 2
$searchContainer.ColumnCount = 1
$searchContainer.Font = $form.Font # Reset font
[void]$searchContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
[void]$searchContainer.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$searchContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblSearchHelp = New-Object System.Windows.Forms.Label
$lblSearchHelp.Text = 'Search is recursive in all files under the folder.'
$lblSearchHelp.AutoSize = $true
$lblSearchHelp.ForeColor = [System.Drawing.Color]::Gray

$searchLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchLayout.Dock = 'Fill'
$searchLayout.RowCount = 7
$searchLayout.ColumnCount = 1
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) 
[void]$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 45))) 
[void]$searchLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblSearchFolder = New-Object System.Windows.Forms.Label
$lblSearchFolder.Text = 'Folder to search (required)'
$lblSearchFolder.AutoSize = $true

$txtSearchFolder = New-Object System.Windows.Forms.TextBox
$txtSearchFolder.Dock = 'Fill'
$txtSearchFolder.BorderStyle = 'FixedSingle'
$txtSearchFolder.Anchor = 'Left, Right'
$txtSearchFolder.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$btnBrowseSearchFolder = New-Object System.Windows.Forms.Button
$btnBrowseSearchFolder.Text = 'Browse...'

$searchFolderLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchFolderLayout.Dock = 'Fill'
$searchFolderLayout.ColumnCount = 2
[void]$searchFolderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$searchFolderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchFolderLayout.Controls.Add($txtSearchFolder, 0, 0)
[void]$searchFolderLayout.Controls.Add($btnBrowseSearchFolder, 1, 0)

$lblKeyword = New-Object System.Windows.Forms.Label
$lblKeyword.Text = 'Keyword 1'
$lblKeyword.AutoSize = $true
$lblKeyword.TextAlign = 'MiddleLeft'
$txtKeyword = New-Object System.Windows.Forms.TextBox
$txtKeyword.Dock = 'Fill'
$txtKeyword.BorderStyle = 'FixedSingle'
$txtKeyword.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$lblKeyword2 = New-Object System.Windows.Forms.Label
$lblKeyword2.Text = 'Keyword 2 (opt.)'
$lblKeyword2.AutoSize = $true
$lblKeyword2.TextAlign = 'MiddleLeft'
$txtKeyword2 = New-Object System.Windows.Forms.TextBox
$txtKeyword2.Dock = 'Fill'
$txtKeyword2.BorderStyle = 'FixedSingle'
$txtKeyword2.Font = New-Object System.Drawing.Font('Bahnschrift', 10)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'Search'

$searchInputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchInputLayout.Dock = 'Fill'
$searchInputLayout.ColumnCount = 3
[void]$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchInputLayout.Controls.Add($lblKeyword, 0, 0)
[void]$searchInputLayout.Controls.Add($txtKeyword, 1, 0)
[void]$searchInputLayout.Controls.Add($btnSearch, 2, 0)

$searchInputLayout2 = New-Object System.Windows.Forms.TableLayoutPanel
$searchInputLayout2.Dock = 'Fill'
$searchInputLayout2.ColumnCount = 3
[void]$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchInputLayout2.Controls.Add($lblKeyword2, 0, 0)
[void]$searchInputLayout2.Controls.Add($txtKeyword2, 1, 0)
[void]$searchInputLayout2.Controls.Add((New-Object System.Windows.Forms.Panel), 2, 0)

$lblAndHelp = New-Object System.Windows.Forms.Label
$lblAndHelp.Text = 'If both keywords are set, a line must contain BOTH.'
$lblAndHelp.AutoSize = $true
$lblAndHelp.ForeColor = [System.Drawing.Color]::Gray

$listResults = New-Object System.Windows.Forms.ListView
$listResults.Dock = 'Fill'
$listResults.View = 'Details'
$listResults.FullRowSelect = $true
$listResults.GridLines = $true
$listResults.HideSelection = $false
$listResults.MultiSelect = $false
$listResults.ShowItemToolTips = $true
$listResults.Columns.Add('File', 620) | Out-Null
$listResults.Columns.Add('Matches', 80) | Out-Null
$listResults.Font = New-Object System.Drawing.Font('Bahnschrift', 9.5)

$btnOpenFile = New-Object System.Windows.Forms.Button
$btnOpenFile.Text = 'Open in editor'
$btnOpenFile.Enabled = $false

$btnOpenFolder = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text = 'Show in Explorer'
$btnOpenFolder.Enabled = $false

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export report'
$btnExport.Enabled = $false

$searchActionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchActionLayout.Dock = 'Fill'
$searchActionLayout.ColumnCount = 7
$searchActionLayout.RowCount = 1
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
[void]$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$btnOpenFile.Anchor = 'None'
$btnOpenFolder.Anchor = 'None'
$btnExport.Anchor = 'None'
[void]$searchActionLayout.Controls.Add($btnOpenFile, 1, 0)
[void]$searchActionLayout.Controls.Add($btnOpenFolder, 3, 0)
[void]$searchActionLayout.Controls.Add($btnExport, 5, 0)

[void]$searchLayout.Controls.Add($lblSearchFolder, 0, 0)
[void]$searchLayout.Controls.Add($searchFolderLayout, 0, 1)
[void]$searchLayout.Controls.Add($searchInputLayout, 0, 2)
[void]$searchLayout.Controls.Add($searchInputLayout2, 0, 3)
[void]$searchLayout.Controls.Add($lblAndHelp, 0, 4)
[void]$searchLayout.Controls.Add($listResults, 0, 5)
[void]$searchLayout.Controls.Add($searchActionLayout, 0, 6)

[void]$searchContainer.Controls.Add($lblSearchHelp, 0, 0)
[void]$searchContainer.Controls.Add($searchLayout, 0, 1)
$grpSearch.Controls.Add($searchContainer)

# Status
$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Status'
$grpLog.Dock = 'Fill'
$grpLog.Padding = New-Object System.Windows.Forms.Padding(10, 18, 10, 10)
$grpLog.BackColor = [System.Drawing.Color]::White
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Dock = 'Fill'
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$txtLog.Font = New-Object System.Drawing.Font('Bahnschrift', 9.5)
$grpLog.Controls.Add($txtLog)

# Final Form Assembly
[void]$contentLayout.Controls.Add($grpExtract, 0, 0)
[void]$contentLayout.Controls.Add($grpClean, 1, 0)
[void]$contentLayout.Controls.Add($grpSearch, 0, 1)
$contentLayout.SetColumnSpan($grpSearch, 2)
[void]$mainLayout.Controls.Add($header, 0, 0)
[void]$mainLayout.Controls.Add($contentLayout, 0, 1)
[void]$mainLayout.Controls.Add($grpLog, 0, 2)
$form.Controls.Add($mainLayout)

# -------------------- STYLING FUNCTIONS --------------------
function Set-ActionButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.FlatStyle = 'Flat'
    $Button.BackColor = [System.Drawing.Color]::FromArgb(0, 121, 107) # Green
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.FlatAppearance.BorderSize = 0
    $Button.Size = New-Object System.Drawing.Size(120, 32)
    $Button.Dock = 'None'
    $Button.Font = New-Object System.Drawing.Font('Bahnschrift', 10)
}
function Set-BrowseButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.FlatStyle = 'Flat'
    $Button.BackColor = [System.Drawing.Color]::FromArgb(236, 239, 241) # Grey
    $Button.ForeColor = [System.Drawing.Color]::FromArgb(33, 33, 33)
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $Button.Size = New-Object System.Drawing.Size(120, 32)
    $Button.Dock = 'None'
    $Button.Anchor = 'None'
    $Button.Font = New-Object System.Drawing.Font('Bahnschrift', 10)
}

Set-ActionButtonStyle -Button $btnExtract
Set-ActionButtonStyle -Button $btnClean
Set-ActionButtonStyle -Button $btnSearch
Set-ActionButtonStyle -Button $btnOpenOutput
Set-ActionButtonStyle -Button $btnCopyOutput
Set-ActionButtonStyle -Button $btnOpenFile
Set-ActionButtonStyle -Button $btnOpenFolder
Set-ActionButtonStyle -Button $btnExport

Set-BrowseButtonStyle -Button $btnBrowse
Set-BrowseButtonStyle -Button $btnBrowseExtractTo
Set-BrowseButtonStyle -Button $btnBrowseFolder
Set-BrowseButtonStyle -Button $btnBrowseSearchFolder

# -------------------- LOGGING & EVENTS --------------------
$script:LogMaxLines = 6
function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = "{0} {1}" -f $ts, $Message
    $current = $txtLog.Text
    if ([string]::IsNullOrWhiteSpace($current)) { $txtLog.Text = $line; return }
    $all = ($current -split "(`r`n|`n|`r)") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $all += $line
    if ($all.Count -gt $script:LogMaxLines) { $all = $all[-$script:LogMaxLines..-1] }
    $txtLog.Text = ($all -join [Environment]::NewLine); $txtLog.SelectionStart = $txtLog.TextLength; $txtLog.ScrollToCaret()
}
function Update-ExtractUIState {
    $zipPath = $txtPath.Text.Trim(); $name = [IO.Path]::GetFileNameWithoutExtension($zipPath)
    if ([string]::IsNullOrWhiteSpace($name)) { $lblOutputFolder.Visible = $false }
    else { $root = $txtExtractTo.Text.Trim(); if ([string]::IsNullOrWhiteSpace($root)) { $root = 'C:\temp' }; $lblOutputFolder.Text = "Output folder: $(Join-Path $root $name)"; $lblOutputFolder.Visible = $true }
    $exists = (-not [string]::IsNullOrWhiteSpace($zipPath) -and (Test-Path -LiteralPath $zipPath)); $btnExtract.Enabled = $exists
}

$btnBrowse.Add_Click({ $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Archives|*.zip;*.eip'; if($d.ShowDialog() -eq 'OK'){ $txtPath.Text = $d.FileName; Update-ExtractUIState } })
$btnBrowseExtractTo.Add_Click({ $d = New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtExtractTo.Text = $d.SelectedPath; Update-ExtractUIState } })
$btnBrowseFolder.Add_Click({ $d = New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtFolder.Text = $d.SelectedPath } })
$btnBrowseSearchFolder.Add_Click({ $d = New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtSearchFolder.Text = $d.SelectedPath } })
$chkPurgeAge.Add_CheckedChanged({ $nudDays.Enabled = $chkPurgeAge.Checked })
$txtPath.Add_TextChanged({ Update-ExtractUIState }); $txtExtractTo.Add_TextChanged({ Update-ExtractUIState })
$form.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' } else { $_.Effect = 'None' } })
$form.Add_DragDrop({ $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop); if ($files) { $f = $files[0]; if ($f -match '\.(zip|eip)$') { $txtPath.Text = $f; Update-ExtractUIState } } })
$listResults.Add_SelectedIndexChanged({ $sel = ($listResults.SelectedItems.Count -gt 0); $btnOpenFile.Enabled = $sel; $btnOpenFolder.Enabled = $sel; $btnExport.Enabled = ($listResults.Items.Count -gt 0) })
$listResults.Add_DoubleClick({ if ($listResults.SelectedItems.Count -gt 0) { Open-FileWithEditor -Path $listResults.SelectedItems[0].Tag.Path } })
$btnOpenFile.Add_Click({ if ($listResults.SelectedItems.Count -gt 0) { Open-FileWithEditor -Path $listResults.SelectedItems[0].Tag.Path } })
$btnOpenFolder.Add_Click({ if ($listResults.SelectedItems.Count -gt 0) { Start-Process 'explorer.exe' -ArgumentList "/select,`"$($listResults.SelectedItems[0].Tag.Path)`"" } })

$btnExtract.Add_Click({
    $src = $txtPath.Text.Trim(); $destRoot = $txtExtractTo.Text.Trim()
    if (-not (Test-Path $src)) { return }; if (-not (Test-Path $destRoot)) { New-Item -Type Directory -Path $destRoot | Out-Null }
    $name = [IO.Path]::GetFileNameWithoutExtension($src); $finalDest = Join-Path $destRoot $name
    $btnExtract.Enabled = $false; $form.UseWaitCursor = $true
    try { 
        Write-Log "Extracting..."
        Expand-ZipRecursive -ZipPath $src -DestinationPath $finalDest
        
        $workingFolder = Get-EffectiveWorkingFolder -DestRoot $finalDest
        $txtFolder.Text = $workingFolder       # <--- AUTO FILL CLEANUP
        $txtSearchFolder.Text = $workingFolder # <--- AUTO FILL SEARCH
        
        $btnOpenOutput.Tag = $workingFolder; $btnCopyOutput.Tag = $workingFolder; $btnOpenOutput.Enabled = $true; $btnCopyOutput.Enabled = $true
        Write-Log "Done."
        [System.Windows.Forms.MessageBox]::Show("Extraction Complete", "Done", "OK", "Information") 
    } catch { Write-Log "Error: $($_.Exception.Message)" } finally { $form.UseWaitCursor = $false; $btnExtract.Enabled = $true }
})

$btnClean.Add_Click({
    $p = $txtFolder.Text.Trim(); if (-not (Test-Path $p)) { return }
    $btnClean.Enabled = $false; $form.UseWaitCursor = $true
    try { Write-Log "Cleaning..."; Clean-FolderKeepLogsTxt -RootPath $p -KeepLogsAndTxt:$chkKeepLogs.Checked -PurgeByAge:$chkPurgeAge.Checked -MaxAgeDays ([int]$nudDays.Value); Write-Log "Done."; [System.Windows.Forms.MessageBox]::Show("Cleanup Complete", "Done", "OK", "Information") } catch { Write-Log "Error: $($_.Exception.Message)" } finally { $form.UseWaitCursor = $false; $btnClean.Enabled = $true }
})
$btnSearch.Add_Click({
    $p = $txtSearchFolder.Text.Trim(); $k1 = $txtKeyword.Text.Trim(); $k2 = $txtKeyword2.Text.Trim(); if (-not (Test-Path $p)) { return }
    $btnSearch.Enabled = $false; $form.UseWaitCursor = $true
    try { Write-Log "Searching..."; $res = Search-FilesForKeywords -RootPath $p -Keyword1 $k1 -Keyword2 $k2 -ResultsList $listResults; Write-Log "Found $($res.Files) files with $($res.Matches) matches."; $btnExport.Enabled = ($listResults.Items.Count -gt 0) } finally { $form.UseWaitCursor = $false; $btnSearch.Enabled = $true }
})
$btnOpenOutput.Add_Click({ if(Test-Path $btnOpenOutput.Tag) { Start-Process $btnOpenOutput.Tag } }); $btnCopyOutput.Add_Click({ if($btnCopyOutput.Tag) { [Windows.Forms.Clipboard]::SetText($btnCopyOutput.Tag) } })
$btnExport.Add_Click({
    $d = New-Object System.Windows.Forms.SaveFileDialog; $d.Filter='Text|*.txt'; $d.FileName='results.txt'
    if ($d.ShowDialog() -eq 'OK') {
        $lines = @("Search Report", "Folder: $($txtSearchFolder.Text)", "Date: $(Get-Date)", ""); foreach($i in $listResults.Items) { $lines += "File: $($i.Tag.Path) (Matches: $($i.SubItems[1].Text))"; foreach($l in $i.Tag.Lines) { $lines += $l }; $lines += "" }
        Set-Content $d.FileName $lines; Open-FileWithEditor $d.FileName
    }
})

[System.Windows.Forms.Application]::EnableVisualStyles(); Write-Log "Ready."; Update-ExtractUIState; [System.Windows.Forms.Application]::Run($form)