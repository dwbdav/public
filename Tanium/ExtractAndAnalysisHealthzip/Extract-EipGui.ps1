Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Expand-ZipRecursive {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [switch]$DeleteZip
    )

    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath | Out-Null
    }

    Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force

    if ($DeleteZip) {
        Remove-Item -LiteralPath $ZipPath -Force
    }

    while ($true) {
        $zips = Get-ChildItem -Path $DestinationPath -Recurse -Filter *.zip -File -ErrorAction SilentlyContinue
        if (-not $zips) {
            break
        }

        foreach ($zip in $zips) {
            $dest = Join-Path -Path $zip.DirectoryName -ChildPath ([IO.Path]::GetFileNameWithoutExtension($zip.Name))
            if (-not (Test-Path -LiteralPath $dest)) {
                New-Item -ItemType Directory -Path $dest | Out-Null
            }
            Expand-Archive -Path $zip.FullName -DestinationPath $dest -Force
            Remove-Item -LiteralPath $zip.FullName -Force
        }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'EIP/ZIP Extraction'
$form.Size = New-Object System.Drawing.Size(780, 820)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Bahnschrift', 9.5)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(12)
$mainLayout.RowCount = 3
$mainLayout.ColumnCount = 1
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72)))
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 90)))

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.BackColor = [System.Drawing.Color]::FromArgb(0, 121, 107)
$header.Padding = New-Object System.Windows.Forms.Padding(14, 10, 14, 10)

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
$lblSub.Location = New-Object System.Drawing.Point(10, 38)

$header.Controls.AddRange(@($lblTitle, $lblSub))

$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.Dock = 'Fill'
$contentLayout.ColumnCount = 2
$contentLayout.RowCount = 2
$contentLayout.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 6)
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 35)))
$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 65)))

$grpExtract = New-Object System.Windows.Forms.GroupBox
$grpExtract.Text = 'Extraction'
$grpExtract.Dock = 'Fill'
$grpExtract.Padding = New-Object System.Windows.Forms.Padding(12, 22, 12, 12)
$grpExtract.BackColor = [System.Drawing.Color]::White
$grpExtract.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)

$grpClean = New-Object System.Windows.Forms.GroupBox
$grpClean.Text = 'Cleanup'
$grpClean.Dock = 'Fill'
$grpClean.Padding = New-Object System.Windows.Forms.Padding(12, 22, 12, 12)
$grpClean.BackColor = [System.Drawing.Color]::White
$grpClean.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)

$grpSearch = New-Object System.Windows.Forms.GroupBox
$grpSearch.Text = 'Search'
$grpSearch.Dock = 'Fill'
$grpSearch.Padding = New-Object System.Windows.Forms.Padding(12, 22, 12, 12)
$grpSearch.BackColor = [System.Drawing.Color]::White
$grpSearch.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)

$extractLayout = New-Object System.Windows.Forms.TableLayoutPanel
$extractLayout.Dock = 'Fill'
$extractLayout.RowCount = 3
$extractLayout.ColumnCount = 1
$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$extractLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$extractLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = 'EIP/ZIP File'
$lblPath.AutoSize = $true

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Dock = 'Fill'
$txtPath.BorderStyle = 'FixedSingle'

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Dock = 'Fill'

$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Text = 'Extract'
$btnExtract.Width = 110

$pathLayout = New-Object System.Windows.Forms.TableLayoutPanel
$pathLayout.Dock = 'Fill'
$pathLayout.ColumnCount = 2
$pathLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$pathLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
$pathLayout.Controls.Add($txtPath, 0, 0)
$pathLayout.Controls.Add($btnBrowse, 1, 0)

$extractActionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$extractActionLayout.Dock = 'Fill'
$extractActionLayout.ColumnCount = 3
$extractActionLayout.RowCount = 1
$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$extractActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$btnExtract.Anchor = 'None'
$extractActionLayout.Controls.Add($btnExtract, 1, 0)

$extractLayout.Controls.Add($lblPath, 0, 0)
$extractLayout.Controls.Add($pathLayout, 0, 1)
$extractLayout.Controls.Add($extractActionLayout, 0, 2)
$grpExtract.Controls.Add($extractLayout)

$cleanLayout = New-Object System.Windows.Forms.TableLayoutPanel
$cleanLayout.Dock = 'Fill'
$cleanLayout.RowCount = 5
$cleanLayout.ColumnCount = 1
$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$cleanLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$cleanLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = 'Folder to clean'
$lblFolder.AutoSize = $true

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Dock = 'Fill'
$txtFolder.BorderStyle = 'FixedSingle'

$btnBrowseFolder = New-Object System.Windows.Forms.Button
$btnBrowseFolder.Text = 'Browse...'
$btnBrowseFolder.Dock = 'Fill'

$chkKeepLogs = New-Object System.Windows.Forms.CheckBox
$chkKeepLogs.Text = 'Keep only .log and .txt'
$chkKeepLogs.AutoSize = $true
$chkKeepLogs.Checked = $true

$chkPurgeAge = New-Object System.Windows.Forms.CheckBox
$chkPurgeAge.Text = 'Delete files older than'
$chkPurgeAge.AutoSize = $true
$chkPurgeAge.Checked = $false

$nudDays = New-Object System.Windows.Forms.NumericUpDown
$nudDays.Minimum = 1
$nudDays.Maximum = 3650
$nudDays.Value = 30
$nudDays.Width = 60
$nudDays.Enabled = $false

$lblDays = New-Object System.Windows.Forms.Label
$lblDays.Text = 'days'
$lblDays.AutoSize = $true

$agePanel = New-Object System.Windows.Forms.FlowLayoutPanel
$agePanel.Dock = 'Fill'
$agePanel.WrapContents = $false
$agePanel.AutoSize = $true
$agePanel.Controls.AddRange(@($chkPurgeAge, $nudDays, $lblDays))

$btnClean = New-Object System.Windows.Forms.Button
$btnClean.Text = 'Clean'
$btnClean.Width = 110

$folderLayout = New-Object System.Windows.Forms.TableLayoutPanel
$folderLayout.Dock = 'Fill'
$folderLayout.ColumnCount = 2
$folderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$folderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
$folderLayout.Controls.Add($txtFolder, 0, 0)
$folderLayout.Controls.Add($btnBrowseFolder, 1, 0)

$cleanActionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$cleanActionLayout.Dock = 'Fill'
$cleanActionLayout.ColumnCount = 3
$cleanActionLayout.RowCount = 1
$cleanActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$cleanActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$cleanActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$btnClean.Anchor = 'None'
$cleanActionLayout.Controls.Add($btnClean, 1, 0)

$cleanLayout.Controls.Add($lblFolder, 0, 0)
$cleanLayout.Controls.Add($folderLayout, 0, 1)
$cleanLayout.Controls.Add($chkKeepLogs, 0, 2)
$cleanLayout.Controls.Add($agePanel, 0, 3)
$cleanLayout.Controls.Add($cleanActionLayout, 0, 4)
$grpClean.Controls.Add($cleanLayout)

$contentLayout.Controls.Add($grpExtract, 0, 0)
$contentLayout.Controls.Add($grpClean, 1, 0)
$contentLayout.Controls.Add($grpSearch, 0, 1)
$contentLayout.SetColumnSpan($grpSearch, 2)

$searchLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchLayout.Dock = 'Fill'
$searchLayout.RowCount = 6
$searchLayout.ColumnCount = 1
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$searchLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
$searchLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$lblSearchFolder = New-Object System.Windows.Forms.Label
$lblSearchFolder.Text = 'Folder to search'
$lblSearchFolder.AutoSize = $true

$txtSearchFolder = New-Object System.Windows.Forms.TextBox
$txtSearchFolder.Dock = 'Fill'
$txtSearchFolder.BorderStyle = 'FixedSingle'

$btnBrowseSearchFolder = New-Object System.Windows.Forms.Button
$btnBrowseSearchFolder.Text = 'Browse...'
$btnBrowseSearchFolder.Dock = 'Fill'

$searchFolderLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchFolderLayout.Dock = 'Fill'
$searchFolderLayout.ColumnCount = 2
$searchFolderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$searchFolderLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchFolderLayout.Controls.Add($txtSearchFolder, 0, 0)
$searchFolderLayout.Controls.Add($btnBrowseSearchFolder, 1, 0)

$lblKeyword = New-Object System.Windows.Forms.Label
$lblKeyword.Text = 'Keyword 1'
$lblKeyword.AutoSize = $true
$lblKeyword.TextAlign = 'MiddleLeft'

$txtKeyword = New-Object System.Windows.Forms.TextBox
$txtKeyword.Dock = 'Fill'
$txtKeyword.BorderStyle = 'FixedSingle'

$lblKeyword2 = New-Object System.Windows.Forms.Label
$lblKeyword2.Text = 'Keyword 2'
$lblKeyword2.AutoSize = $true
$lblKeyword2.TextAlign = 'MiddleLeft'

$txtKeyword2 = New-Object System.Windows.Forms.TextBox
$txtKeyword2.Dock = 'Fill'
$txtKeyword2.BorderStyle = 'FixedSingle'

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'Search'
$btnSearch.Width = 110

$searchInputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchInputLayout.Dock = 'Fill'
$searchInputLayout.ColumnCount = 3
$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$searchInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchInputLayout.Controls.Add($lblKeyword, 0, 0)
$searchInputLayout.Controls.Add($txtKeyword, 1, 0)
$searchInputLayout.Controls.Add($btnSearch, 2, 0)

$searchInputLayout2 = New-Object System.Windows.Forms.TableLayoutPanel
$searchInputLayout2.Dock = 'Fill'
$searchInputLayout2.ColumnCount = 3
$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$searchInputLayout2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchInputLayout2.Controls.Add($lblKeyword2, 0, 0)
$searchInputLayout2.Controls.Add($txtKeyword2, 1, 0)
$searchInputLayout2.Controls.Add((New-Object System.Windows.Forms.Panel), 2, 0)

$listResults = New-Object System.Windows.Forms.ListView
$listResults.Dock = 'Fill'
$listResults.View = 'Details'
$listResults.FullRowSelect = $true
$listResults.GridLines = $true
$listResults.HideSelection = $false
$listResults.MultiSelect = $false
$listResults.Columns.Add('File', 520) | Out-Null
$listResults.Columns.Add('Matches', 90) | Out-Null

$btnOpenFile = New-Object System.Windows.Forms.Button
$btnOpenFile.Text = 'Open file'
$btnOpenFile.Width = 120
$btnOpenFile.Enabled = $false

$btnOpenFolder = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text = 'Open folder'
$btnOpenFolder.Width = 120
$btnOpenFolder.Enabled = $false

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export'
$btnExport.Width = 120
$btnExport.Enabled = $false

$searchActionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$searchActionLayout.Dock = 'Fill'
$searchActionLayout.ColumnCount = 7
$searchActionLayout.RowCount = 1
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 12)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 12)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
$searchActionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$btnOpenFile.Anchor = 'None'
$btnOpenFolder.Anchor = 'None'
$btnExport.Anchor = 'None'
$searchActionLayout.Controls.Add($btnOpenFile, 1, 0)
$searchActionLayout.Controls.Add($btnOpenFolder, 3, 0)
$searchActionLayout.Controls.Add($btnExport, 5, 0)

$searchLayout.Controls.Add($lblSearchFolder, 0, 0)
$searchLayout.Controls.Add($searchFolderLayout, 0, 1)
$searchLayout.Controls.Add($searchInputLayout, 0, 2)
$searchLayout.Controls.Add($searchInputLayout2, 0, 3)
$searchLayout.Controls.Add($listResults, 0, 4)
$searchLayout.Controls.Add($searchActionLayout, 0, 5)
$grpSearch.Controls.Add($searchLayout)

$grpLog = New-Object System.Windows.Forms.GroupBox
$grpLog.Text = 'Status'
$grpLog.Dock = 'Fill'
$grpLog.Padding = New-Object System.Windows.Forms.Padding(12, 18, 12, 12)
$grpLog.BackColor = [System.Drawing.Color]::White

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $false
$txtLog.ScrollBars = 'None'
$txtLog.ReadOnly = $true
$txtLog.Dock = 'Fill'
$txtLog.BorderStyle = 'FixedSingle'
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$txtLog.Font = New-Object System.Drawing.Font('Bahnschrift', 9.5)
$txtLog.TextAlign = 'Left'
$txtLog.Text = 'Ready.'
$txtLog.Height = 40

$grpLog.Controls.Add($txtLog)

$mainLayout.Controls.Add($header, 0, 0)
$mainLayout.Controls.Add($contentLayout, 0, 1)
$mainLayout.Controls.Add($grpLog, 0, 2)
$form.Controls.Add($mainLayout)

function Set-PrimaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.FlatStyle = 'Flat'
    $Button.BackColor = [System.Drawing.Color]::FromArgb(0, 121, 107)
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.FlatAppearance.BorderSize = 0
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 137, 123)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 105, 92)
    $Button.Height = 30
}

function Set-SecondaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)
    $Button.FlatStyle = 'Flat'
    $Button.BackColor = [System.Drawing.Color]::FromArgb(236, 239, 241)
    $Button.ForeColor = [System.Drawing.Color]::FromArgb(33, 33, 33)
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(220, 224, 227)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(207, 216, 220)
    $Button.Height = 28
}

Set-PrimaryButtonStyle -Button $btnExtract
Set-PrimaryButtonStyle -Button $btnClean
Set-PrimaryButtonStyle -Button $btnSearch
Set-SecondaryButtonStyle -Button $btnBrowse
Set-SecondaryButtonStyle -Button $btnBrowseFolder
Set-SecondaryButtonStyle -Button $btnBrowseSearchFolder
Set-SecondaryButtonStyle -Button $btnOpenFile
Set-SecondaryButtonStyle -Button $btnOpenFolder
Set-SecondaryButtonStyle -Button $btnExport

function Write-Log {
    param([string]$Message)
    $txtLog.Text = $Message
}

function Open-FileWithEditor {
    param([Parameter(Mandatory = $true)][string]$Path)

    $npp64 = 'C:\Program Files\Notepad++\notepad++.exe'
    $npp32 = 'C:\Program Files (x86)\Notepad++\notepad++.exe'
    $editor = $null

    if (Test-Path -LiteralPath $npp64) {
        $editor = $npp64
    } elseif (Test-Path -LiteralPath $npp32) {
        $editor = $npp32
    }

    if ($editor) {
        Start-Process -FilePath $editor -ArgumentList @("`"$Path`"")
    } else {
        Start-Process -FilePath $Path
    }
}

function Clean-FolderKeepLogsTxt {
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [switch]$KeepLogsAndTxt,
        [switch]$PurgeByAge,
        [int]$MaxAgeDays
    )

    if (-not (Test-Path -LiteralPath $RootPath)) {
        throw "Folder not found: $RootPath"
    }

    $cutoff = $null
    if ($PurgeByAge -and $MaxAgeDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-1 * $MaxAgeDays)
    }

    if ($KeepLogsAndTxt -or $cutoff) {
        $files = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $remove = $false
            if ($KeepLogsAndTxt) {
                $ext = $file.Extension.ToLowerInvariant()
                if ($ext -ne '.log' -and $ext -ne '.txt') {
                    $remove = $true
                }
            }
            if (-not $remove -and $cutoff) {
                if ($file.LastWriteTime -lt $cutoff) {
                    $remove = $true
                }
            }
            if ($remove) {
                Remove-Item -LiteralPath $file.FullName -Force
            }
        }
    }

    $dirs = Get-ChildItem -Path $RootPath -Recurse -Directory -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending
    foreach ($dir in $dirs) {
        $items = Get-ChildItem -LiteralPath $dir.FullName -Force -ErrorAction SilentlyContinue
        if (-not $items) {
            Remove-Item -LiteralPath $dir.FullName -Force
        }
    }
}

function Search-FilesForKeywords {
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [string]$Keyword1,
        [string]$Keyword2,
        [Parameter(Mandatory = $true)][System.Windows.Forms.ListView]$ResultsList
    )

    $ResultsList.Items.Clear()
    $ResultsList.BeginUpdate()
    $filesMatched = 0
    $totalMatches = 0
    try {
        $comparison = [System.StringComparison]::OrdinalIgnoreCase
        $hasKeyword1 = -not [string]::IsNullOrWhiteSpace($Keyword1)
        $hasKeyword2 = -not [string]::IsNullOrWhiteSpace($Keyword2)
        $files = Get-ChildItem -Path $RootPath -Recurse -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                $lines = Get-Content -Path $file.FullName -ReadCount 0 -ErrorAction Stop
                if ($lines -is [string]) {
                    $lines = @($lines)
                }
            } catch {
                continue
            }

            $matchCount = 0
            $matchedLines = @()
            for ($i = 0; $i -lt $lines.Count; $i++) {
                $line = $lines[$i]
                if ([string]::IsNullOrWhiteSpace($line)) {
                    continue
                }
                $match = $false
                if ($hasKeyword1 -and $hasKeyword2) {
                    $match = ($line.IndexOf($Keyword1, $comparison) -ge 0 -and $line.IndexOf($Keyword2, $comparison) -ge 0)
                } elseif ($hasKeyword1) {
                    $match = ($line.IndexOf($Keyword1, $comparison) -ge 0)
                } elseif ($hasKeyword2) {
                    $match = ($line.IndexOf($Keyword2, $comparison) -ge 0)
                }

                if ($match) {
                    $matchCount++
                    $lineNumber = $i + 1
                    $matchedLines += ("{0}`t{1}" -f $lineNumber, $line)
                }
            }

            if ($matchCount -gt 0) {
                $item = New-Object System.Windows.Forms.ListViewItem($file.FullName)
                $item.SubItems.Add($matchCount.ToString()) | Out-Null
                $item.Tag = [PSCustomObject]@{
                    Path = $file.FullName
                    Lines = $matchedLines
                }
                $ResultsList.Items.Add($item) | Out-Null
                $filesMatched++
                $totalMatches += $matchCount
            }
        }
    } finally {
        $ResultsList.EndUpdate()
    }

    return @{
        Files = $filesMatched
        Matches = $totalMatches
    }
}

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = 'EIP/ZIP (*.eip;*.zip)|*.eip;*.zip|All files (*.*)|*.*'
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq 'OK') {
        $txtPath.Text = $dialog.FileName
    }
})

$btnBrowseFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq 'OK') {
        $txtFolder.Text = $dialog.SelectedPath
    }
})

$btnBrowseSearchFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dialog.ShowDialog() -eq 'OK') {
        $txtSearchFolder.Text = $dialog.SelectedPath
    }
})

$chkPurgeAge.Add_CheckedChanged({
    $nudDays.Enabled = $chkPurgeAge.Checked
})

$listResults.Add_SelectedIndexChanged({
    $hasSelection = $listResults.SelectedItems.Count -gt 0
    $btnOpenFile.Enabled = $hasSelection
    $btnOpenFolder.Enabled = $hasSelection
    $btnExport.Enabled = $listResults.Items.Count -gt 0
})

$listResults.Add_DoubleClick({
    if ($listResults.SelectedItems.Count -eq 0) {
        return
    }
    $tag = $listResults.SelectedItems[0].Tag
    $path = if ($tag -is [string]) { $tag } else { [string]$tag.Path }
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
        Open-FileWithEditor -Path $path
    }
})

$btnOpenFile.Add_Click({
    if ($listResults.SelectedItems.Count -eq 0) {
        return
    }
    $tag = $listResults.SelectedItems[0].Tag
    $path = if ($tag -is [string]) { $tag } else { [string]$tag.Path }
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
        Open-FileWithEditor -Path $path
    }
})

$btnOpenFolder.Add_Click({
    if ($listResults.SelectedItems.Count -eq 0) {
        return
    }
    $tag = $listResults.SelectedItems[0].Tag
    $path = if ($tag -is [string]) { $tag } else { [string]$tag.Path }
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$path`""
    }
})

$btnExport.Add_Click({
    if ($listResults.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No results to export.', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = 'Text file (*.txt)|*.txt|All files (*.*)|*.*'
    $dialog.FileName = 'search-results.txt'
    $dialog.OverwritePrompt = $true

    if ($dialog.ShowDialog() -ne 'OK') {
        return
    }

    $lines = @()
    $lines += "Folder: $($txtSearchFolder.Text.Trim())"
    $lines += "Keyword 1: $($txtKeyword.Text.Trim())"
    $lines += "Keyword 2: $($txtKeyword2.Text.Trim())"
    $lines += "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += ''
    $lines += "File`tMatches"

    foreach ($item in $listResults.Items) {
        $tag = $item.Tag
        $path = if ($tag -is [string]) { $item.Text } else { [string]$tag.Path }
        $occ = $item.SubItems[1].Text
        $lines += ''
        $lines += "File: $path"
        $lines += "Matches: $occ"
        $lines += "Line`tText"
        if ($tag -and $tag.PSObject.Properties.Match('Lines').Count -gt 0) {
            foreach ($matchLine in $tag.Lines) {
                $lines += $matchLine
            }
        }
    }

    Set-Content -Path $dialog.FileName -Value $lines -Encoding UTF8
    Write-Log "Export completed: $($dialog.FileName)"
    Open-FileWithEditor -Path $dialog.FileName
})

$btnExtract.Add_Click({
    $zipPath = $txtPath.Text.Trim()
    if (-not (Test-Path -LiteralPath $zipPath)) {
        [System.Windows.Forms.MessageBox]::Show('File not found.', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $destRoot = Join-Path -Path 'C:\temp' -ChildPath ([IO.Path]::GetFileNameWithoutExtension($zipPath))
    if (-not (Test-Path -LiteralPath 'C:\temp')) {
        New-Item -ItemType Directory -Path 'C:\temp' | Out-Null
    }

    $btnExtract.Enabled = $false
    $form.UseWaitCursor = $true
    try {
        Write-Log "Start: extracting to $destRoot"
        Expand-ZipRecursive -ZipPath $zipPath -DestinationPath $destRoot
        $txtFolder.Text = $destRoot
        $txtSearchFolder.Text = $destRoot
        Write-Log 'Done: extraction'
        [System.Windows.Forms.MessageBox]::Show('Extraction completed.', 'Success', 'OK', 'Information') | Out-Null
    } catch {
        Write-Log "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $form.UseWaitCursor = $false
        $btnExtract.Enabled = $true
    }
})

$btnClean.Add_Click({
    $rootPath = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($rootPath)) {
        [System.Windows.Forms.MessageBox]::Show('Select a folder.', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $btnClean.Enabled = $false
    $form.UseWaitCursor = $true
    try {
        $daysValue = [int]$nudDays.Value
        $ageSuffix = ''
        if ($chkPurgeAge.Checked) {
            $ageSuffix = " (older than $daysValue days)"
        }
        Write-Log "Start: cleaning files in $rootPath$ageSuffix"
        Clean-FolderKeepLogsTxt -RootPath $rootPath -KeepLogsAndTxt:($chkKeepLogs.Checked) -PurgeByAge:($chkPurgeAge.Checked) -MaxAgeDays $daysValue
        Write-Log 'Done: cleanup'
        [System.Windows.Forms.MessageBox]::Show('Cleanup completed.', 'Success', 'OK', 'Information') | Out-Null
    } catch {
        Write-Log "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $form.UseWaitCursor = $false
        $btnClean.Enabled = $true
    }
})

$btnSearch.Add_Click({
    $rootPath = $txtSearchFolder.Text.Trim()
    $keyword1 = $txtKeyword.Text.Trim()
    $keyword2 = $txtKeyword2.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($rootPath)) {
        [System.Windows.Forms.MessageBox]::Show('Select a folder.', 'Error', 'OK', 'Error') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($keyword1) -and [string]::IsNullOrWhiteSpace($keyword2)) {
        [System.Windows.Forms.MessageBox]::Show('Enter at least one keyword.', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $btnSearch.Enabled = $false
    $btnExport.Enabled = $false
    $form.UseWaitCursor = $true
    try {
        if (-not [string]::IsNullOrWhiteSpace($keyword1) -and -not [string]::IsNullOrWhiteSpace($keyword2)) {
            Write-Log "Start: search '$keyword1' and '$keyword2' in $rootPath"
        } elseif (-not [string]::IsNullOrWhiteSpace($keyword1)) {
            Write-Log "Start: search '$keyword1' in $rootPath"
        } else {
            Write-Log "Start: search '$keyword2' in $rootPath"
        }
        $result = Search-FilesForKeywords -RootPath $rootPath -Keyword1 $keyword1 -Keyword2 $keyword2 -ResultsList $listResults
        Write-Log ("Done: search ({0} files, {1} matches)" -f $result.Files, $result.Matches)
        $btnExport.Enabled = $listResults.Items.Count -gt 0
    } catch {
        Write-Log "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $form.UseWaitCursor = $false
        $btnSearch.Enabled = $true
    }
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
