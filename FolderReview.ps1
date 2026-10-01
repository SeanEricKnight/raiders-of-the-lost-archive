<#
.SYNOPSIS
  Inventory one source folder and its descendants on an archaeological expedition,
  however many layers of archived civilization stand between them and the workbook.
.DESCRIPTION
  The selected folder tree receives an archaeological survey. Always produces a six-column inventory CSV, plus an XLSX when Excel is available
  and, when needed, an errors CSV for places that declined
  to be observed. No existing workbook
  is opened or updated. The six columns match the shared review template.
  Copy reviewed rows into the team's shared workbook manually.

  Prompts for source and report folders when paths are omitted. The source menu
  offers a Windows folder picker, typed paths, and desktop/documents/downloads.
  A full path or drive letter can also be entered directly at the menu.
  OwnerAlias defaults to no substitutions. ReviewerNotes preserves the true owner
  for any alias supplied. ExcludeFolder applies to top-level folders of the source only.
  A recovery CSV is written incrementally during scanning and retained after success.
  Pressing Esc during a scan saves the rows collected so far to a recovery CSV.
  The report folder may not be inside the source folder.
  Long-path support depends on the Windows/.NET host; failures are logged explicitly.
  Source files and their ownership remain untouched. The expedition is strictly observational.
.EXAMPLE
  .\FolderReview.ps1 -FolderName 'Templates'
.EXAMPLE
  .\FolderReview.ps1
  Prompts for paths, then offers the entire source folder or a child folder.
.EXAMPLE
  .\FolderReview.ps1 -SourceRoot 'C:\Documents' -ReportFolder 'C:\Reports' -ScanRoot
.EXAMPLE
  .\FolderReview.ps1 -SourceRoot '\\server\share\Security' -ExcludeFolder '00 Proposed Folder Structure' -OwnerAlias @{ 'DOMAIN\olduser' = 'DOMAIN\newuser' }
.NOTES
  Windows PowerShell 5.1 or PowerShell 7. Excel is optional and enables XLSX output.
  Compatible VT terminals get color effects and an arrow-key menu. Other hosts retain
  the numbered menu and original display. -PlainOutput disables VT rendering;
  -NumberedMenu retains colors but uses numbered selection.
#>
[CmdletBinding()]
param(
    [string]$FolderName,
    [string]$SourceRoot,
    [switch]$ScanRoot,
    [string]$ReportFolder,
    [switch]$SkipWindowsProperties,
    [switch]$PlainOutput,
    [switch]$NumberedMenu,
    [ValidateNotNull()][hashtable]$OwnerAlias = @{},
    [string[]]$ExcludeFolder = @(),
    [string[]]$ExcludeFile = @('~$*', 'desktop.ini', 'Thumbs.db')
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Esc = [char]27
$UseVT = $false
try {
    $UseVT = -not $PlainOutput -and [bool]$Host.UI.SupportsVirtualTerminal -and
        ($PSVersionTable.PSVersion.Major -ge 7 -or -not [string]::IsNullOrEmpty($env:WT_SESSION)) -and
        -not [Console]::IsOutputRedirected
} catch { $UseVT = $false }
$CursorHidden = $false

function Get-TerminalWidth {
    try { return [Math]::Max(20, [Console]::WindowWidth) }
    catch { return 80 }
}

function Write-WrappedMessage {
    param([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Yellow)
    $Width = [Math]::Max(1, ((Get-TerminalWidth) - 1))
    foreach ($Paragraph in ($Text -split "`r?`n")) {
        $Remaining = $Paragraph
        while ($Remaining.Length -gt $Width) {
            $Break = $Remaining.LastIndexOf(' ', $Width - 1, $Width)
            if ($Break -le 0) { $Break = $Width }
            Write-Host $Remaining.Substring(0, $Break) -ForegroundColor $Color
            $Remaining = $Remaining.Substring($Break).TrimStart()
        }
        Write-Host $Remaining -ForegroundColor $Color
    }
}

function Get-ExceptionSummary {
    param([Exception]$Exception)
    $Types = New-Object 'System.Collections.Generic.List[string]'
    $Current = $Exception
    while ($null -ne $Current) {
        $Types.Add($Current.GetType().FullName)
        $Current = $Current.InnerException
    }
    ($Types -join ' -> ') + ': ' + $Exception.Message
}

function Test-IsLink {
    # True links (symlinks, junctions) only. Dedup'd or cloud-tiered files carry the
    # reparse flag but are ordinary files, so PowerShell 7 uses LinkTarget instead.
    param($Item)
    try {
        if ($Item.PSObject.Properties['LinkTarget']) {
            return -not [string]::IsNullOrEmpty([string]$Item.LinkTarget)
        }
    } catch {}
    return (($Item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Select-SourceFolderDialog {
    # A dedicated STA runspace supports the Windows picker in both PS 5.1
    # and PS 7, including sessions whose main thread uses MTA.
    $PickerRunspace = $null
    $PickerPowerShell = $null
    try {
        $PickerRunspace = [RunspaceFactory]::CreateRunspace()
        $PickerRunspace.ApartmentState = [Threading.ApartmentState]::STA
        $PickerRunspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $PickerRunspace.Open()
        $PickerPowerShell = [PowerShell]::Create()
        $PickerPowerShell.Runspace = $PickerRunspace
        [void]$PickerPowerShell.AddScript({
            Add-Type -AssemblyName System.Windows.Forms
            $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            try {
                $Dialog.Description = 'Choose the source folder to review'
                $Dialog.ShowNewFolderButton = $false
                $Dialog.SelectedPath = [Environment]::GetFolderPath('MyDocuments')
                if ($Dialog.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
                    $Dialog.SelectedPath
                }
            } finally { $Dialog.Dispose() }
        }.ToString())
        $Selection = $PickerPowerShell.Invoke()
        if ($PickerPowerShell.HadErrors) {
            throw ($PickerPowerShell.Streams.Error | Out-String)
        }
        if ($Selection.Count -gt 0) { return [string]$Selection[0] }
        return ''
    } finally {
        if ($null -ne $PickerPowerShell) { $PickerPowerShell.Dispose() }
        if ($null -ne $PickerRunspace) { $PickerRunspace.Dispose() }
    }
}

function Resolve-SourceInput {
    param([string]$Text)
    $Text = $Text.Trim().Trim('"')
    switch -Regex ($Text) {
        '^desktop$'          { return [Environment]::GetFolderPath('Desktop') }
        '^(documents|docs)$' { return [Environment]::GetFolderPath('MyDocuments') }
        '^downloads$'        { return (Join-Path $env:USERPROFILE 'Downloads') }
        '^[A-Za-z]:?$'       { return ($Text.Substring(0,1).ToUpper() + ':\') }
    }
    return [Environment]::ExpandEnvironmentVariables($Text)
}

function Export-RowsCsv {
    param($Rows, [string]$Path)
    if (@($Rows).Count -gt 0) {
        $Rows | Select-Object FullPath,
            @{Name='LastModified'; Expression={$_.LastModified.ToString('o')}},
            FileSystemOwner, DocumentAuthor, ReviewDecision, ReviewerNotes |
            Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    } else {
        '"FullPath","LastModified","FileSystemOwner","DocumentAuthor","ReviewDecision","ReviewerNotes"' |
            Set-Content -LiteralPath $Path -Encoding UTF8
    }
}

function Write-Banner {
    if (-not $script:UseVT) { return }
    $Rule = ([string][char]0x2501) * [Math]::Min(62, ((Get-TerminalWidth) - 1))
    Write-Host ''
    Write-ColorLine $Rule '70;225;255' Cyan
    Write-Host "${Esc}[38;2;70;225;255m  F O L D E R   R E V I E W${Esc}[0m"
    Write-ColorLine '  raiders of the lost archive' '205;215;225' Gray
    Write-ColorLine $Rule '70;225;255' Cyan
    Write-Host ''
}

function Write-Certificate {
    param([string]$Title, [string[]]$Lines)
    $W = [int](@($Title) + $Lines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $H = [string][char]0x2550; $V = [string][char]0x2551
    $TL = [string][char]0x2554; $TR = [string][char]0x2557
    $ML = [string][char]0x2560; $MR = [string][char]0x2563
    $BL = [string][char]0x255A; $BR = [string][char]0x255D
    $Items = @([pscustomobject]@{K='top';T=''}, [pscustomobject]@{K='title';T=$Title}, [pscustomobject]@{K='mid';T=''}) +
        @($Lines | ForEach-Object { [pscustomobject]@{K='line';T=$_} }) +
        @([pscustomobject]@{K='bot';T=''})
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $c = "${Esc}[38;2;255;205;70m"
        $t = $Items[$i].T
        switch ($Items[$i].K) {
            'top' { $Text = "$c$TL$($H * ($W + 2))$TR" }
            'mid' { $Text = "$c$ML$($H * ($W + 2))$MR" }
            'bot' { $Text = "$c$BL$($H * ($W + 2))$BR" }
            'title' {
                $Pad = [int][Math]::Floor(($W - $t.Length) / 2)
                $Text = "$c$V ${Esc}[1;97m" + ((' ' * $Pad) + $t).PadRight($W) + "${Esc}[0m $c$V"
            }
            'line' { $Text = "$c$V ${Esc}[97m" + $t.PadRight($W) + "${Esc}[0m $c$V" }
        }
        Write-Host "$Text${Esc}[0m"
    }
}

function Show-FancyProgress {
    param([string]$Stage, [int]$Done, [int]$Total, [int]$Folders, [switch]$Force)
    # Show-ProgressPanel applies the same 80 ms throttle to both renderers.
    $Frames = [char[]]@(0x280B,0x2819,0x2839,0x2838,0x283C,0x2834,0x2826,0x2827,0x2807,0x280F)
    $Spin = $Frames[[int][Math]::Floor((Get-Date).Millisecond / 100)]
    $Available = (Get-TerminalWidth) - 1
    if ($Stage -eq 'Counting') {
        $Text = " $Spin COUNTING  $Done files  $Folders folders"
        if ($Text.Length -gt $Available) { $Text = $Text.Substring(0, $Available) }
        Write-Host -NoNewline ("`r${Esc}[38;2;70;225;255m$Text${Esc}[0m${Esc}[K")
        return
    }
    $Percent = if ($Total -le 0) { 100 } else { [int][Math]::Floor(100 * $Done / $Total) }
    $Percent = [Math]::Max(0, [Math]::Min(100, $Percent))
    $Label = " $Spin SCANNING  {0}/{1}  {2,3}%  " -f $Done, $Total, $Percent
    if ($Label.Length -ge $Available) { $Label = " $Percent% " }
    $Width = [Math]::Max(1, [Math]::Min(40, ($Available - $Label.Length)))
    $Filled = [int][Math]::Floor($Width * $Percent / 100)
    $sb = New-Object Text.StringBuilder
    for ($i = 0; $i -lt $Width; $i++) {
        if ($i -lt $Filled) {
            [void]$sb.Append("${Esc}[38;2;70;225;255m").Append([char]0x2588)
        } else {
            [void]$sb.Append("${Esc}[38;2;150;165;180m").Append([char]0x2591)
        }
    }
    [void]$sb.Append("${Esc}[0m")
    Write-Host -NoNewline ("`r${Esc}[96m$Label${Esc}[0m" + $sb.ToString() + "${Esc}[K")
}

function Select-ReviewFolder {
    # Folders are objects with Name (display text) and Item (the DirectoryInfo to return).
    param([object[]]$Folders)
    # Console.ReadKey is unavailable in ISE, redirected input, and many hosted sessions.
    $CanUseKeys = $false
    try {
        $CanUseKeys = $script:UseVT -and -not $script:NumberedMenu -and
            $Host.Name -eq 'ConsoleHost' -and -not [Console]::IsInputRedirected
        if ($CanUseKeys) { $null = [Console]::KeyAvailable }
    } catch { $CanUseKeys = $false }
    if ($CanUseKeys) {
        $Selected = 0
        $Drawn = 0
        Write-WrappedMessage 'Up/Down: choose a folder. Enter: confirm. Esc: cancel.' Gray
        Write-Host ''
        try {
            Write-Host -NoNewline "${Esc}[?25l"
            while ($true) {
                # Page the list so cursor movement never runs beyond the viewport.
                $Height = 20
                try { $Height = [Math]::Max(1, [Math]::Min(20, ([Console]::WindowHeight - 6))) } catch {}
                $Start = [int][Math]::Floor($Selected / $Height) * $Height
                $End = [Math]::Min($Folders.Count, ($Start + $Height))
                if ($Drawn -gt 0) { Write-Host -NoNewline "${Esc}[${Drawn}A" }
                $Width = (Get-TerminalWidth) - 1
                # Clear leftover rows if the final page is shorter.
                $PageRows = $End - $Start
                $RenderRows = [Math]::Max($PageRows, $Drawn)
                for ($j = 0; $j -lt $RenderRows; $j++) {
                    Write-Host -NoNewline "`r${Esc}[2K"
                    if ($j -lt $PageRows) {
                        $i = $Start + $j
                        $Mark = if ($i -eq $Selected) { '>' } else { ' ' }
                        $Label = '{0} {1,2}. {2}' -f $Mark, ($i + 1), $Folders[$i].Name
                        if ($Label.Length -gt $Width) { $Label = $Label.Substring(0, [Math]::Max(1, ($Width - 3))) + '...' }
                        if ($i -eq $Selected) {
                            Write-Host "${Esc}[48;2;20;40;60m${Esc}[38;2;70;225;255m$Label${Esc}[0m"
                        } else { Write-ColorLine $Label '205;215;225' Gray }
                    } else { Write-Host '' }
                }
                $Drawn = $RenderRows
                $Key = [Console]::ReadKey($true)
                switch ($Key.Key) {
                    'UpArrow' { $Selected = ($Selected + $Folders.Count - 1) % $Folders.Count }
                    'DownArrow' { $Selected = ($Selected + 1) % $Folders.Count }
                    'Home' { $Selected = 0 }
                    'End' { $Selected = $Folders.Count - 1 }
                    'Enter' { return $Folders[$Selected].Item }
                    'Escape' { throw [OperationCanceledException]::new('Folder selection cancelled.') }
                }
            }
        } finally { Write-Host -NoNewline "${Esc}[0m${Esc}[?25h" }
    }
    for ($i = 0; $i -lt $Folders.Count; $i++) {
        $Label = '{0,2}. {1}' -f ($i + 1), $Folders[$i].Name
        Write-ColorLine $Label $script:MenuRgbColors[$i % $script:MenuRgbColors.Count] $script:MenuColors[$i % $script:MenuColors.Count]
    }
    do {
        $Choice = Read-Host 'Choose a folder to review'
        $Number = 0
        $ValidChoice = [int]::TryParse($Choice, [ref]$Number) -and $Number -ge 1 -and $Number -le $Folders.Count
        if (-not $ValidChoice) {
            Write-Host ('Choose a number from 1 to {0}. Folder {1} is not on this map.' -f $Folders.Count, $Choice) -ForegroundColor Yellow
        }
    } until ($ValidChoice)
    return $Folders[$Number - 1].Item
}

$MenuColors = @('Red', 'DarkYellow', 'Yellow', 'DarkYellow', 'Green', 'DarkGreen',
                'DarkGreen', 'Cyan', 'Cyan', 'Blue', 'Blue', 'Magenta')
$MenuRgbColors = @(
    '255;75;75', '255;120;55', '255;165;50', '255;210;60',
    '215;235;65', '160;235;70', '90;225;110', '45;215;160',
    '35;205;210', '55;175;245', '100;145;255', '155;120;245'
)
$ClosingRgbColors = @('255;95;95', '255;165;45', '255;220;60',
                      '100;185;255', '210;120;255', '255;115;205')

function Write-ColorLine {
    param([string]$Text, [string]$Rgb, [ConsoleColor]$FallbackColor)
    if ($script:UseVT) {
        $Escape = [char]27
        Write-Host "${Escape}[38;2;${Rgb}m${Text}${Escape}[0m"
    } else {
        Write-Host $Text -ForegroundColor $FallbackColor
    }
}

function Format-DateValue {
    param($Value)

    if ($null -eq $Value) { return "" }

    try {
        if ($Value -is [datetime]) {
            return $Value.ToString("yyyy-MM-dd HH:mm:ss")
        }

        $Parsed = [datetime]$Value
        return $Parsed.ToString("yyyy-MM-dd HH:mm:ss")
    }
    catch {
        return [string]$Value
    }
}

function Convert-PropertyValueToString {
    param($Value)

    if ($null -eq $Value) { return "" }

    if ($Value -is [System.Array]) {
        return (($Value | ForEach-Object { [string]$_ }) -join "; ")
    }

    return [string]$Value
}

function Get-FileSystemOwnerSafe {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    try {
        if ($PSVersionTable.PSEdition -eq 'Core') {
            $Acl = [IO.FileSystemAclExtensions]::GetAccessControl($File, [Security.AccessControl.AccessControlSections]::Owner)
        } else {
            $Acl = $File.GetAccessControl([Security.AccessControl.AccessControlSections]::Owner)
        }
        try { $Owner = $Acl.GetOwner([Security.Principal.NTAccount]).Value }
        catch [Security.Principal.IdentityNotMappedException] {
            $Owner = $Acl.GetOwner([Security.Principal.SecurityIdentifier]).Value
        }
        return [PSCustomObject]@{
            Owner  = [string]$Owner
            Status = "OK"
            Error  = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Owner  = ""
            Status = "ERROR"
            Error  = Get-ExceptionSummary $_.Exception
        }
    }
}

function Get-OpenXmlCoreProperties {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $Result = [ordered]@{
        Author          = ""
        LastSavedBy     = ""
        DocumentCreated = ""
        Source          = ""
        Status          = "NotApplicable"
        Error           = ""
    }

    $OpenXmlExtensions = @(
        ".docx", ".docm", ".dotx", ".dotm",
        ".xlsx", ".xlsm", ".xltx", ".xltm",
        ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"
    )

    $Extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($OpenXmlExtensions -notcontains $Extension) {
        return [PSCustomObject]$Result
    }

    $Archive = $null
    $FileStream = $null
    $Stream  = $null
    $Reader  = $null

    try {
        # Shared reads reduce conflicts; exclusive locks can still deny access.
        $FileStream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $Archive = New-Object IO.Compression.ZipArchive($FileStream, [IO.Compression.ZipArchiveMode]::Read, $true)
        $Entry = $Archive.GetEntry("docProps/core.xml")

        if ($null -eq $Entry) {
            $Result.Status = "NoCoreProperties"
            return [PSCustomObject]$Result
        }

        $Stream = $Entry.Open()
        $Reader = New-Object System.IO.StreamReader($Stream)
        $XmlText = $Reader.ReadToEnd()

        [xml]$Xml = $XmlText

        $Ns = New-Object System.Xml.XmlNamespaceManager($Xml.NameTable)
        $Ns.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
        $Ns.AddNamespace("dc", "http://purl.org/dc/elements/1.1/")
        $Ns.AddNamespace("dcterms", "http://purl.org/dc/terms/")

        $CreatorNode = $Xml.SelectSingleNode("/cp:coreProperties/dc:creator", $Ns)
        $LastAuthorNode = $Xml.SelectSingleNode("/cp:coreProperties/cp:lastModifiedBy", $Ns)
        $CreatedNode = $Xml.SelectSingleNode("/cp:coreProperties/dcterms:created", $Ns)

        if ($null -ne $CreatorNode) {
            $Result.Author = $CreatorNode.InnerText.Trim()
        }

        if ($null -ne $LastAuthorNode) {
            $Result.LastSavedBy = $LastAuthorNode.InnerText.Trim()
        }

        if ($null -ne $CreatedNode) {
            $Result.DocumentCreated = Format-DateValue $CreatedNode.InnerText.Trim()
        }

        $Result.Source = "OpenXML core properties"

        if (($Result.Author -ne "") -or ($Result.LastSavedBy -ne "") -or ($Result.DocumentCreated -ne "")) {
            $Result.Status = "OK"
        }
        else {
            $Result.Status = "NoAuthorMetadata"
        }
    }
    catch {
        $Result.Status = "ERROR"
        $Result.Error = Get-ExceptionSummary $_.Exception
    }
    finally {
        if ($null -ne $Reader) {
            $Reader.Dispose()
        }
        elseif ($null -ne $Stream) {
            $Stream.Dispose()
        }

        if ($null -ne $Archive) { $Archive.Dispose() }
        if ($null -ne $FileStream) { $FileStream.Dispose() }
    }

    return [PSCustomObject]$Result
}

function Get-WindowsPropertyMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        $ShellApplication,

        [Parameter(Mandatory = $true)]
        [hashtable]$FolderCache
    )

    $Result = [ordered]@{
        Author      = ""
        LastSavedBy = ""
        Source      = ""
        Status      = "Unavailable"
        Error       = ""
    }

    $Item = $null
    try {
        $FolderPath = [IO.Path]::GetDirectoryName($Path)
        $FileName   = [IO.Path]::GetFileName($Path)

        if (-not $FolderCache.ContainsKey($FolderPath)) {
            $FolderCache[$FolderPath] = $ShellApplication.Namespace($FolderPath)
        }

        $Folder = $FolderCache[$FolderPath]
        if ($null -eq $Folder) {
            $Result.Status = "FolderNamespaceUnavailable"
            return [PSCustomObject]$Result
        }

        $Item = $Folder.ParseName($FileName)
        if ($null -eq $Item) {
            $Result.Status = "ItemUnavailable"
            return [PSCustomObject]$Result
        }

        $Author = $Item.ExtendedProperty("System.Author")
        $LastAuthor = $Item.ExtendedProperty("System.Document.LastAuthor")

        $Result.Author = Convert-PropertyValueToString $Author
        $Result.LastSavedBy = Convert-PropertyValueToString $LastAuthor
        $Result.Source = "Windows Property System"

        if (($Result.Author -ne "") -or ($Result.LastSavedBy -ne "")) {
            $Result.Status = "OK"
        }
        else {
            $Result.Status = "NoAuthorMetadata"
        }
    }
    catch {
        $Result.Status = "ERROR"
        $Result.Error = Get-ExceptionSummary $_.Exception
    }

    finally {
        if ($null -ne $Item) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Item) } catch {}
        }
    }
    return [PSCustomObject]$Result
}


function New-ReviewWorkbook {
    param([string]$Path, [System.Collections.IList]$ReviewRows, [string]$WorksheetName, [System.Collections.IList]$FolderSizeRows, [string]$SizeStatus)
    $Excel = $null; $Book = $null; $Sheet = $null; $Range = $null; $SizeSheet = $null
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        $Book = $Excel.Workbooks.Add()
        $Sheet = $Book.Worksheets.Item(1)
        # Excel sheet names allow at most 31 characters and exclude these symbols.
        $TabName = ($WorksheetName -replace '[\\/\?\*\[\]:\x00-\x1F]', '-').Trim().Trim("'")
        if ([string]::IsNullOrWhiteSpace($TabName)) { $TabName = 'Review' }
        if ($TabName.Length -gt 31) { $TabName = $TabName.Substring(0, 31).TrimEnd().TrimEnd("'") }
        if ($TabName -ieq 'History') { $TabName = 'History Review' }
        if ($TabName -ieq 'Folder Sizes') { $TabName = 'Folder Sizes Review' }
        $Sheet.Name = $TabName
        $Headers = @('FullPath','LastModified','FileSystemOwner','DocumentAuthor','ReviewDecision','ReviewerNotes')
        $HeaderCells = New-Object 'object[,]' 1, 6
        for ($c = 0; $c -lt 6; $c++) { $HeaderCells[0,$c] = $Headers[$c] }
        $Sheet.Range('A1','F1').Value2 = $HeaderCells
        $Sheet.Range('A1','F1').Font.Bold = $true
        if ($ReviewRows.Count -gt 1048575) {
            throw 'The inventory exceeds the Excel worksheet row limit; use the recovery CSV.'
        }
        # Keep metadata literal so leading '=' cannot become an Excel formula.
        $Sheet.Range('A:F').NumberFormat = '@'
        $Sheet.Range('B:B').NumberFormat = 'mm/dd/yyyy'
        $ChunkSize = 5000
        for ($Offset = 0; $Offset -lt $ReviewRows.Count; $Offset += $ChunkSize) {
            $Count = [Math]::Min($ChunkSize, $ReviewRows.Count - $Offset)
            $Data = New-Object 'object[,]' $Count, 6
            for ($i = 0; $i -lt $Count; $i++) {
                $Row = $ReviewRows[$Offset + $i]
                $Data[$i,0] = [string]$Row.FullPath
                $Data[$i,1] = ([datetime]$Row.LastModified).ToOADate()
                $Data[$i,2] = [string]$Row.FileSystemOwner
                $Data[$i,3] = [string]$Row.DocumentAuthor
                $Data[$i,4] = [string]$Row.ReviewDecision
                $Data[$i,5] = [string]$Row.ReviewerNotes
            }
            $Range = $Sheet.Range(('A' + ($Offset + 2)), ('F' + ($Offset + $Count + 1)))
            try { $Range.Value2 = $Data }
            finally {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Range)
                $Range = $null
            }
        }
        # Assign controlled HYPERLINK formulas in batches; escape quotes in paths.
        # Other metadata remains literal text. Formulas avoid creating thousands
        # of individual COM hyperlink objects.
        $Sheet.Range('A:A').NumberFormat = 'General'
        for ($Offset = 0; $Offset -lt $ReviewRows.Count; $Offset += $ChunkSize) {
            $Count = [Math]::Min($ChunkSize, $ReviewRows.Count - $Offset)
            $Links = New-Object 'object[,]' $Count, 1
            for ($i = 0; $i -lt $Count; $i++) {
                $LiteralPath = ([string]$ReviewRows[$Offset + $i].FullPath).Replace('"', '""')
                $Links[$i,0] = '=HYPERLINK("' + $LiteralPath + '","' + $LiteralPath + '")'
            }
            $Range = $Sheet.Range(('A' + ($Offset + 2)), ('A' + ($Offset + $Count + 1)))
            try { $Range.Formula = $Links }
            finally {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Range)
                $Range = $null
            }
        }
        $Sheet.Columns.Item(1).ColumnWidth = 90
        $Sheet.Columns.Item(2).ColumnWidth = 15
        $Sheet.Columns.Item(3).ColumnWidth = 28
        $Sheet.Columns.Item(4).ColumnWidth = 30
        $Sheet.Columns.Item(5).ColumnWidth = 20
        $Sheet.Columns.Item(6).ColumnWidth = 55
        try {
            $Sheet.Range('A1', ('F' + [Math]::Max(1, $ReviewRows.Count + 1))).AutoFilter() | Out-Null
            $Sheet.Activate() | Out-Null
            $Sheet.Range('A2').Select() | Out-Null
            $Excel.ActiveWindow.FreezePanes = $true
        } catch { Write-Warning "Could not set Excel filters or freeze the header: $($_.Exception.Message)" }
        if ($FolderSizeRows.Count -gt 1048572) { throw 'Folder sizes exceed the worksheet row limit; use the CSV reports.' }
        $SizeSheet = $Book.Worksheets.Add()
        $SizeSheet.Name = 'Folder Sizes'
        $SizeSheet.Range('A:D').NumberFormat = '@'
        $SizeSheet.Range('B:C').NumberFormat = '0'
        $SizeSheet.Range('D:D').NumberFormat = '0.00'
        $SizeSheet.Range('A1').Value2 = 'Folder sizes include subfolders. Parent and child totals overlap.'
        $SizeSheet.Range('A2').Value2 = $SizeStatus
        $SizeHeaders = New-Object 'object[,]' 1, 4
        $SizeHeaders[0,0] = 'Folder'
        $SizeHeaders[0,1] = 'TotalBytes'
        $SizeHeaders[0,2] = 'FileCount'
        $SizeHeaders[0,3] = 'SizeGiB'
        $SizeSheet.Range('A4','D4').Value2 = $SizeHeaders
        $SizeSheet.Range('A4','D4').Font.Bold = $true
        for ($Offset = 0; $Offset -lt $FolderSizeRows.Count; $Offset += $ChunkSize) {
            $Count = [Math]::Min($ChunkSize, $FolderSizeRows.Count - $Offset)
            $SizeData = New-Object 'object[,]' $Count, 4
            for ($i = 0; $i -lt $Count; $i++) {
                $Entry = $FolderSizeRows[$Offset + $i]
                $SizeData[$i,0] = [string]$Entry.Folder
                $SizeData[$i,1] = [double]$Entry.TotalBytes
                $SizeData[$i,2] = [double]$Entry.FileCount
                $SizeData[$i,3] = [double]$Entry.SizeGiB
            }
            $Range = $SizeSheet.Range(('A' + ($Offset + 5)), ('D' + ($Offset + $Count + 4)))
            try { $Range.Value2 = $SizeData }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Range); $Range = $null }
        }
        $SizeSheet.Columns.Item(1).ColumnWidth = 90
        $SizeSheet.Columns.Item(2).ColumnWidth = 20
        $SizeSheet.Columns.Item(3).ColumnWidth = 15
        $SizeSheet.Columns.Item(4).ColumnWidth = 15
        $SizeSheet.Range('A4', ('D' + [Math]::Max(4, $FolderSizeRows.Count + 4))).AutoFilter() | Out-Null
        $Sheet.Activate() | Out-Null
        $Book.SaveAs($Path, 51) # XLSX
    }
    finally {
        if ($null -ne $Book) { try { $Book.Close($false) } catch {} }
        if ($null -ne $Excel) { try { $Excel.Quit() } catch {} }
        foreach ($Item in @($Range,$SizeSheet,$Sheet,$Book,$Excel)) {
            if ($null -ne $Item) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Item) } catch {} }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Restore terminal state on completion, errors, and normal PowerShell cancellation.
try {
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'FolderReview requires Windows.'
}
if ($ScanRoot -and -not [string]::IsNullOrWhiteSpace($FolderName)) {
    throw 'Use either -ScanRoot or -FolderName, not both.'
}
Write-Banner
if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
    while ($true) {
        Write-Host ''
        Write-WrappedMessage 'Choose a source folder:' Cyan
        Write-Host ''
        Write-WrappedMessage '1. Browse for a folder' Gray
        Write-WrappedMessage '2. Enter a local or network path' Gray
        Write-WrappedMessage '3. Desktop' Gray
        Write-WrappedMessage '4. Documents' Gray
        Write-WrappedMessage '5. Downloads' Gray
        Write-WrappedMessage 'Q. Quit' Gray
        Write-Host ''
        $Choice = (Read-Host 'Choose 1-5, enter a path, or Q').Trim()
        if ([string]::IsNullOrWhiteSpace($Choice) -or $Choice -ieq 'q') {
            Write-Host 'Cancelled.' -ForegroundColor Yellow
            return
        }
        $Entry = ''
        switch ($Choice) {
            '1' {
                try { $Entry = Select-SourceFolderDialog }
                catch {
                    Write-WrappedMessage ('Folder browser unavailable: ' + $_.Exception.Message)
                    Write-WrappedMessage 'Choose option 2 to enter the path instead.'
                    continue
                }
            }
            '2' { $Entry = Read-Host 'Source path (Enter to return to the menu)' }
            '3' { $Entry = 'desktop' }
            '4' { $Entry = 'documents' }
            '5' { $Entry = 'downloads' }
            default { $Entry = $Choice }
        }
        # Cancelling the picker or leaving the path blank returns to the menu.
        if ([string]::IsNullOrWhiteSpace($Entry)) { continue }
        $Candidate = Resolve-SourceInput $Entry
        $Ok = $false
        try { $Ok = Test-Path -LiteralPath $Candidate -PathType Container } catch {}
        if ($Ok) { $SourceRoot = $Candidate; break }
        Write-WrappedMessage "Folder not found or inaccessible: $Candidate"
    }
} elseif (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "The tomb entrance is blocked. Cannot access the source folder: $SourceRoot"
}
$SourceDirectory = Get-Item -LiteralPath $SourceRoot -Force
if ($SourceDirectory -isnot [IO.DirectoryInfo]) { throw 'The source must be a filesystem folder.' }
$SourceRoot = $SourceDirectory.FullName
$Desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($Desktop)) { $Desktop = [Environment]::GetFolderPath('MyDocuments') }
$DefaultReportFolder = Join-Path $Desktop 'FolderReviewReports'
$SourceNormal = $SourceRoot.TrimEnd('\') + '\'
while ($true) {
    if ([string]::IsNullOrWhiteSpace($ReportFolder)) {
        Write-Host ''
        Write-WrappedMessage 'Choose where to save reports:' Cyan
        Write-Host ''
        Write-WrappedMessage "Press Enter for: $DefaultReportFolder" Gray
        $ReportFolder = (Read-Host 'Report folder').Trim().Trim('"')
        if ([string]::IsNullOrWhiteSpace($ReportFolder)) { $ReportFolder = $DefaultReportFolder }
    }
    $ReportFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportFolder)
    $ReportNormal = $ReportFolder.TrimEnd('\') + '\'
    if (-not $ReportNormal.StartsWith($SourceNormal, [StringComparison]::OrdinalIgnoreCase)) { break }
    Write-Host ''
    Write-WrappedMessage 'Choose a report folder outside the source folder, so nothing is written to the tree being audited.'
    Write-WrappedMessage "Source folder: $SourceRoot" Gray
    Write-WrappedMessage "Report folder: $ReportFolder" Gray
    Write-WrappedMessage 'Enter a different report folder. To cancel, press Ctrl+C.'
    $ReportFolder = ''
}
# Only enumerate child choices when a child selection is needed.
$TopLevelFolders = @()
if (-not $ScanRoot) {
    $TopLevelFolders = @(Get-ChildItem -LiteralPath $SourceRoot -Directory -Force |
        Where-Object { $ExcludeFolder -notcontains $_.Name } |
        Sort-Object Name)
}
if ($ScanRoot) {
    $TargetFolder = $SourceDirectory
} elseif ([string]::IsNullOrWhiteSpace($FolderName)) {
    Write-Host ''
    Write-ColorLine 'Choose a folder to review:' '180;130;255' Magenta
    Write-Host ''
    Write-ColorLine "Source folder: $SourceRoot" '205;215;225' Gray
    Write-Host ''
    Write-WrappedMessage 'The first entry scans the entire source folder, including files directly inside it.' Gray
    Write-Host ''
    $RootLabel = if ([string]::IsNullOrWhiteSpace($SourceDirectory.Name)) { $SourceRoot } else { $SourceDirectory.Name }
    $Choices = @([PSCustomObject]@{ Name = "(entire source folder: $RootLabel)"; Item = $SourceDirectory }) +
        @($TopLevelFolders | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Item = $_ } })
    try { $TargetFolder = Select-ReviewFolder -Folders $Choices }
    catch [OperationCanceledException] {
        Write-Host 'Cancelled.' -ForegroundColor Yellow
        return
    }
} else {
    $TargetFolder = $TopLevelFolders | Where-Object { $_.Name -ieq $FolderName } | Select-Object -First 1
    if ($null -eq $TargetFolder) { throw "That folder is missing from the expedition map: $FolderName" }
}

[void](New-Item -ItemType Directory -Path $ReportFolder -Force)
Add-Type -AssemblyName System.IO.Compression
$Shell = $null
$FolderCache = @{}
if (-not $SkipWindowsProperties) {
    try { $Shell = New-Object -ComObject Shell.Application }
    catch { Write-Warning "Windows property lookup unavailable: $($_.Exception.Message)" }
}
$Rows = New-Object 'System.Collections.Generic.List[object]'
$Errors = New-Object 'System.Collections.Generic.List[object]'
$FolderCount = 0
$FileCount = 0
$TotalFiles = 0
$CountedFolders = 0
$FilesToReview = New-Object 'System.Collections.Generic.List[object]'
$FolderTotals = @{}
$LastProgressDraw = [datetime]::MinValue
$ProgressPanelWidth = 58

# Report file names are settled up front so an early stop can still save its rows.
$FolderLabel = $TargetFolder.Name
if ([string]::IsNullOrWhiteSpace($FolderLabel) -or $null -eq $TargetFolder.Parent) {
    $FolderLabel = $TargetFolder.FullName.TrimEnd('\')
    if ($FolderLabel -match '^([A-Za-z]):$') { $FolderLabel = $Matches[1].ToUpper() + ' Drive' }
}
$SafeName = ($FolderLabel -replace '[<>:"/\\|?*\x00-\x1F]', '-').Trim().TrimEnd('.')
if ([string]::IsNullOrWhiteSpace($SafeName)) { $SafeName = 'Root' }
$Stamp = (Get-Date).ToString('yyyy-MM-dd')
$Suffix = 0
while ($true) {
    $Number = if ($Suffix -eq 0) { '' } else { " ($Suffix)" }
    $Xlsx = Join-Path $ReportFolder "$SafeName - Review - $Stamp$Number.xlsx"
    $ErrorCsv = Join-Path $ReportFolder "$SafeName - Issues - $Stamp$Number.csv"
    $Backup = Join-Path $ReportFolder "$SafeName - Inventory - $Stamp$Number.csv"
    $FolderSizeCsv = Join-Path $ReportFolder "$SafeName - Folder Sizes - $Stamp$Number.csv"
    if (-not ((Test-Path -LiteralPath $Xlsx) -or (Test-Path -LiteralPath $ErrorCsv) -or (Test-Path -LiteralPath $Backup) -or (Test-Path -LiteralPath $FolderSizeCsv))) { break }
    $Suffix++
}

function Format-ProgressPanelLine {
    param([string]$Text)
    $InnerWidth = $script:ProgressPanelWidth - 2
    if ($Text.Length -gt $InnerWidth) {
        $Text = $Text.Substring(0, $InnerWidth - 3) + '...'
    }
    '[' + $Text.PadRight($InnerWidth) + ']'
}

function Show-ProgressPanel {
    param([string]$Stage, [int]$Done, [int]$Total, [int]$Folders, [switch]$Force)
    $Now = Get-Date
    if (-not $Force -and ($Now - $script:LastProgressDraw).TotalMilliseconds -lt 80) { return }
    $script:LastProgressDraw = $Now
    if ($script:UseVT) { Show-FancyProgress @PSBoundParameters; return }
    if ($Stage -eq 'Counting') {
        $Line = Format-ProgressPanelLine (" COUNTING   $Done files   $Folders folders")
    } else {
        $Percent = if ($Total -eq 0) { 100 } else { [int][Math]::Floor(100 * $Done / $Total) }
        $Counts = " SCANNING   $Done/$Total files   $Percent%  "
        $Width = [Math]::Max(1, $script:ProgressPanelWidth - 4 - $Counts.Length)
        $Filled = if ($Total -eq 0) { $Width } else { [int][Math]::Floor($Width * $Done / $Total) }
        $Filled = [Math]::Min($Width, $Filled)
        $Bar = ([string][char]0x2588 * $Filled) + ([string][char]0x2591 * ($Width - $Filled))
        $Line = Format-ProgressPanelLine ($Counts + [char]0x2502 + $Bar + [char]0x2502)
    }
    Write-Host -NoNewline "`r$Line" -ForegroundColor Cyan
}

function Test-StopRequested {
    if (-not $script:CanReadStopKey) { return }
    try {
        while ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true)
            if ($Key.Key -eq [ConsoleKey]::Escape) {
                throw [OperationCanceledException]::new('Stopped by user.')
            }
        }
    } catch [OperationCanceledException] { throw }
    catch {
        $script:CanReadStopKey = $false
        Write-Host ''
        Write-Warning 'Keyboard polling became unavailable. Use Ctrl+C to stop.'
    }
}

function Count-ReviewFiles {
    param($Folder, [switch]$SelectedRoot)
    Test-StopRequested
    # Reports inside the source tree are excluded from the inventory.
    if ($Folder.FullName.TrimEnd('\') -ieq $script:ReportFolder.TrimEnd('\')) { return }
    # The explicitly selected root may be a link; child links remain excluded.
    if (-not $SelectedRoot -and (Test-IsLink $Folder)) {
        Add-ScanError $Folder.FullName 'ReparsePointSkipped' 'Directory link was not traversed.'
        return
    }
    $script:FolderTotals[$Folder.FullName] = [pscustomobject]@{ Folder=$Folder.FullName; TotalBytes=[long]0; FileCount=[long]0 }
    $script:CountedFolders++
    Show-ProgressPanel -Stage Counting -Done $script:TotalFiles -Folders $script:CountedFolders
    # Enumerate each network directory once without creating PowerShell provider objects.
    try { $Items = $Folder.GetFileSystemInfos() }
    catch {
        Add-ScanException $Folder.FullName 'FolderReadError' $_.Exception
        return
    }
    foreach ($Item in $Items) {
        Test-StopRequested
        if ($Item -is [System.IO.DirectoryInfo]) {
            # ExcludeFolder applies to top-level folders of the source only.
            $SkipIt = ($script:ExcludeFolder -contains $Item.Name) -and
                      ($Item.Parent.FullName.TrimEnd('\') -ieq $script:SourceRoot.TrimEnd('\'))
            if (-not $SkipIt) { Count-ReviewFiles -Folder $Item }
        }
        else {
            $Excluded = $false
            foreach ($Pattern in $script:ExcludeFile) {
                if ($Item.Name -like $Pattern) { $Excluded = $true; break }
            }
            if ($Excluded) { continue }
            if (-not (Test-IsLink $Item)) {
                try {
                    $Bytes = [long]$Item.Length
                    $ParentFolder = $Item.Directory
                    while ($null -ne $ParentFolder -and $script:FolderTotals.ContainsKey($ParentFolder.FullName)) {
                        $Total = $script:FolderTotals[$ParentFolder.FullName]
                        $Total.TotalBytes += $Bytes
                        $Total.FileCount++
                        $ParentFolder = $ParentFolder.Parent
                    }
                } catch { Add-ScanException $Item.FullName 'FileSizeError' $_.Exception }
            }
            $script:FilesToReview.Add($Item)
            $script:TotalFiles++
            Show-ProgressPanel -Stage Counting -Done $script:TotalFiles -Folders $script:CountedFolders
        }
    }
}

function Show-ScanProgress {
    Show-ProgressPanel -Stage Scanning -Done $script:FileCount -Total $script:TotalFiles -Folders $script:FolderCount
}

function Add-ScanError {
    param([string]$Path, [string]$Kind, [string]$Message)
    if ($Message -match 'PathTooLongException') { $Kind = 'PathTooLongException' }
    $script:Errors.Add([PSCustomObject]@{ FullPath=$Path; Issue=$Kind; Details=$Message })
}
function Add-ScanException {
    param([string]$Path, [string]$Kind, [Exception]$Exception)
    $Current = $Exception
    while ($null -ne $Current) {
        if ($Current -is [IO.PathTooLongException]) { $Kind = 'PathTooLongException' }
        $Current = $Current.InnerException
    }
    Add-ScanError $Path $Kind (Get-ExceptionSummary $Exception)
}

function Add-ReviewFile {
    param($File)
    $Path = $File.FullName
    $script:FileCount++
    Show-ScanProgress
    if (Test-IsLink $File) {
        Add-ScanError $Path 'ReparsePointSkipped' 'File link was not read.'
        return
    }
    $Owner = Get-FileSystemOwnerSafe -File $File
    if ($Owner.Error) { Add-ScanError $Path 'OwnerLookupError' $Owner.Error }
    $Author = ''
    $OpenXml = Get-OpenXmlCoreProperties -Path $Path
    if ($OpenXml.Status -eq 'ERROR') { Add-ScanError $Path 'DocumentMetadataError' $OpenXml.Error }
    $Author = $OpenXml.Author
    if (-not $Author -and $null -ne $script:Shell) {
        $Windows = Get-WindowsPropertyMetadata -Path $Path -ShellApplication $script:Shell -FolderCache $script:FolderCache
        if ($Windows.Status -eq 'ERROR') { Add-ScanError $Path 'WindowsMetadataError' $Windows.Error }
        elseif ($Windows.Status -in @('FolderNamespaceUnavailable', 'ItemUnavailable')) {
            Add-ScanError $Path 'WindowsMetadataUnavailable' "Windows property lookup failed ($($Windows.Status)); DocumentAuthor may be blank. Long paths commonly cause this."
        }
        if ($Windows.Author) { $Author = $Windows.Author }
    }
    $ReportedOwner = $Owner.Owner
    $OwnerNote = ''
    if ($script:OwnerAlias.ContainsKey($ReportedOwner)) {
        $ReportedOwner = [string]$script:OwnerAlias[$ReportedOwner]
        $OwnerNote = "Owner reporting alias: $($Owner.Owner) -> $ReportedOwner. Actual filesystem owner: $($Owner.Owner)."
    }
    $script:Rows.Add([PSCustomObject]@{
        FullPath=$Path
        LastModified=$File.LastWriteTime
        FileSystemOwner=$ReportedOwner
        DocumentAuthor=$Author
        ReviewDecision=''
        ReviewerNotes=$OwnerNote
    })
}
$CanReadStopKey = $false
try {
    $CanReadStopKey = $Host.Name -eq 'ConsoleHost' -and -not [Console]::IsInputRedirected
    if ($CanReadStopKey) { $null = [Console]::KeyAvailable }
} catch { $CanReadStopKey = $false }
Write-Host ''
if ($CanReadStopKey) {
    Write-ColorLine 'Press Esc to stop counting or scanning.' '205;215;225' Gray
} else {
    Write-Host 'Use Ctrl+C to stop counting or scanning.' -ForegroundColor Gray
}
Write-Host ''
if ($UseVT) {
    $CursorHidden = $true
    Write-Host -NoNewline "${Esc}[?25l"
}
# Flush each completed row to disk. Ctrl+C or a closed terminal can bypass
# cancellation handlers, so recovery must not depend on those handlers running.
$RecoveryWriter = $null
try {
    $RecoveryWriter = New-Object IO.StreamWriter($Backup, $false, (New-Object Text.UTF8Encoding($true)))
    $RecoveryWriter.AutoFlush = $true
    $RecoveryWriter.WriteLine('"FullPath","LastModified","FileSystemOwner","DocumentAuthor","ReviewDecision","ReviewerNotes"')
    Count-ReviewFiles -Folder $TargetFolder -SelectedRoot
    $FolderCount = $CountedFolders
    Show-ProgressPanel -Stage Counting -Done $TotalFiles -Folders $CountedFolders -Force
    Show-ScanProgress
    foreach ($File in $FilesToReview) {
        Test-StopRequested
        $PreviousRowCount = $Rows.Count
        try { Add-ReviewFile -File $File }
        catch { Add-ScanException $File.FullName 'FileReadError' $_.Exception }
        if ($Rows.Count -gt $PreviousRowCount) {
            $CompletedRow = $Rows[$Rows.Count - 1]
            $CsvLines = @($CompletedRow | Select-Object FullPath,
                @{Name='LastModified'; Expression={$_.LastModified.ToString('o')}},
                FileSystemOwner, DocumentAuthor, ReviewDecision, ReviewerNotes |
                ConvertTo-Csv -NoTypeInformation)
            # Keep write failures outside the per-file catch: do not continue
            # scanning while the recovery copy can no longer be updated.
            $RecoveryWriter.WriteLine($CsvLines[1])
        }
    }
    Test-StopRequested
    Show-ProgressPanel -Stage Scanning -Done $FileCount -Total $TotalFiles -Folders $FolderCount -Force
} catch [OperationCanceledException] {
    Write-Host ''
    Write-Host 'Stopped - partial inventory saved. No review workbook was created.' -ForegroundColor Yellow
    Write-Host "Completed rows were saved to: $Backup" -ForegroundColor Yellow
    if ($Errors.Count -gt 0) {
        try { $Errors | Export-Csv -LiteralPath $ErrorCsv -NoTypeInformation -Encoding UTF8 } catch {}
    }
    return
} finally {
    if ($null -ne $RecoveryWriter) { $RecoveryWriter.Dispose() }
    foreach ($CachedFolder in $FolderCache.Values) {
        if ($null -ne $CachedFolder) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($CachedFolder) } catch {}
        }
    }
    $FolderCache.Clear()
    if ($null -ne $Shell) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Shell) } catch {}
        $Shell = $null
    }
}
$SortedRows = @($Rows | Sort-Object @{Expression={[IO.Path]::GetDirectoryName($_.FullPath)}}, FullPath)
# Keep the flushed recovery CSV in scan order. The workbook uses SortedRows.
# Do not replace the recovery file after scanning: it is already complete.
if ($Errors.Count -gt 0) { $Errors | Export-Csv -LiteralPath $ErrorCsv -NoTypeInformation -Encoding UTF8 }
Write-Host ''
Write-Host ''
$SizeStatus = if ($Errors.Count -gt 0) { 'Partial totals: scan issues were recorded. Totals reflect included files only.' } else { 'Totals reflect included files only; configured exclusions and links are omitted.' }
$FolderSizeRows = @($FolderTotals.Values | Sort-Object @{Expression='TotalBytes';Descending=$true}, Folder | ForEach-Object {
    [pscustomobject]@{ Folder=$_.Folder; TotalBytes=$_.TotalBytes; FileCount=$_.FileCount; SizeGiB=([double]$_.TotalBytes / 1GB); Status=$SizeStatus }
})
$FolderSizeRows | Export-Csv -LiteralPath $FolderSizeCsv -NoTypeInformation -Encoding UTF8
$RootTotal = $FolderTotals[$TargetFolder.FullName]
$WorkbookCreated = $false
$ExcelAvailable = $false
try { $ExcelAvailable = $null -ne [Type]::GetTypeFromProgID('Excel.Application') } catch {}
if ($ExcelAvailable) {
    Write-WrappedMessage 'Building the workbook in Excel. Large inventories can take a minute...' DarkGray
    try {
        New-ReviewWorkbook -Path $Xlsx -ReviewRows $SortedRows -WorksheetName $FolderLabel -FolderSizeRows $FolderSizeRows -SizeStatus $SizeStatus
        $WorkbookCreated = $true
    } catch {
        $ExcelError = $_.Exception.Message
        if (Test-Path -LiteralPath $Xlsx) {
            try { Remove-Item -LiteralPath $Xlsx -Force }
            catch { Write-WrappedMessage "Partial workbook could not be removed: $Xlsx" }
        }
        Write-WrappedMessage "XLSX creation failed: $ExcelError"
        Write-WrappedMessage 'The inventory CSV was saved successfully and remains usable.'
    }
} else {
    Write-WrappedMessage 'Desktop Excel is unavailable. The inventory is saved as CSV.' Gray
}
# CSV is a standard report, retained whether or not XLSX creation succeeds.
if ($CursorHidden) {
    Write-Host -NoNewline "${Esc}[0m${Esc}[?25h"
    $CursorHidden = $false
}
Write-Host ''
$CompletionStatus = if ($Errors.Count -gt 0) { 'Complete with issues' } else { 'Complete' }
$CompletionStatus += if ($WorkbookCreated) { ' - CSV and XLSX saved' } else { ' - CSV saved' }
Write-ColorLine $CompletionStatus '70;225;255' Cyan
if ($Errors.Count -gt 0) {
    Write-Warning "$($Errors.Count) errors or skipped paths; the inventory may be incomplete."
}
Write-Host ''
$CertificateTitle = 'EXPEDITION FIELD REPORT'
$CertificateFields = @(
    [PSCustomObject]@{ Label='Files inventoried:'; Value=$SortedRows.Count },
    [PSCustomObject]@{ Label='Folders scanned:'; Value=$FolderCount },
    [PSCustomObject]@{ Label='Issues found:'; Value=$Errors.Count },
    [PSCustomObject]@{ Label='Total size:'; Value=('{0:N2} GiB{1}' -f ([double]$RootTotal.TotalBytes / 1GB), $(if ($Errors.Count -gt 0) { ' (partial)' } else { '' })) }
)
$LabelWidth = ($CertificateFields | ForEach-Object { $_.Label.Length } | Measure-Object -Maximum).Maximum + 1
$CertificateLines = @($CertificateFields | ForEach-Object {
    $_.Label.PadRight($LabelWidth) + [string]$_.Value
})
if ($UseVT) {
    Write-Certificate -Title $CertificateTitle -Lines $CertificateLines
} else {
    $CertificateWidth = (@($CertificateTitle) + $CertificateLines | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    $CertificateBorder = '+' + ('-' * ($CertificateWidth + 2)) + '+'
    $CertificateDivider = '|' + ('-' * ($CertificateWidth + 2)) + '|'
    Write-ColorLine $CertificateBorder '255;205;70' Yellow
    Write-ColorLine ('| ' + $CertificateTitle.PadRight($CertificateWidth) + ' |') '255;205;70' Yellow
    Write-ColorLine $CertificateDivider '255;205;70' Yellow
    foreach ($CertificateLine in $CertificateLines) {
        Write-ColorLine ('| ' + $CertificateLine.PadRight($CertificateWidth) + ' |') '255;205;70' Yellow
    }
    Write-ColorLine $CertificateBorder '255;205;70' Yellow
}
Write-Host ''
if ($WorkbookCreated) { Write-WrappedMessage "Review workbook (XLSX): $Xlsx" Blue }
Write-WrappedMessage "Inventory (CSV): $Backup" Gray
Write-WrappedMessage "Folder sizes (CSV): $FolderSizeCsv" Gray
Write-Host ''
if ($Errors.Count -gt 0) {
    Write-WrappedMessage 'Issue summary:' Yellow
    $IssueLabels = @{
        FolderReadError = 'Folders could not be read'
        FileReadError = 'Files could not be inventoried'
        FileSizeError = 'File sizes could not be read'
        OwnerLookupError = 'Filesystem owners could not be read'
        DocumentMetadataError = 'Document author metadata could not be read'
        WindowsMetadataError = 'Windows document metadata lookup failed'
        WindowsMetadataUnavailable = 'Windows document metadata was unavailable'
        ReparsePointSkipped = 'Links were skipped to avoid following other locations'
        PathTooLongException = 'Paths exceeded the host-supported length'
    }
    foreach ($Group in @($Errors | Group-Object Issue | Sort-Object Count -Descending)) {
        $Label = if ($IssueLabels.ContainsKey($Group.Name)) { $IssueLabels[$Group.Name] } else { $Group.Name }
        Write-WrappedMessage ('  {0}: {1}' -f $Group.Count, $Label) Yellow
        # Names provide context without repeating long paths in the field report.
        $Examples = @($Group.Group | Select-Object -First 3 | ForEach-Object {
            $Name = [IO.Path]::GetFileName($_.FullPath.TrimEnd('\'))
            if ([string]::IsNullOrWhiteSpace($Name)) { $Name = '(root folder)' }
            $Name
        })
        Write-WrappedMessage ('    Affected: ' + ($Examples -join '; ')) Gray
        $Denied = @($Group.Group | Where-Object { $_.Details -match 'UnauthorizedAccessException|Access is denied|access.*denied' }).Count
        if ($Denied -gt 0) { Write-WrappedMessage ("    Access denied was reported for $Denied of these issues.") Gray }
        if ($Group.Count -gt 3) { Write-WrappedMessage ('    Additional affected items: ' + ($Group.Count - 3)) Gray }
    }
    Write-WrappedMessage 'Full paths and error details are in the Issues CSV.' Gray
    Write-Host ''
}
if ($Errors.Count -gt 0) { Write-ColorLine "Issues (CSV): $ErrorCsv" '255;220;60' Yellow }
Write-Host ''
$ClosingRemark = Get-Random -InputObject @(
    'The inventory belongs in a museum. The spreadsheet will have to do.',
    'No files were altered. The boulder has been referred to Facilities.',
    'Artifacts catalogued. Please do not replace the workbook with a bag of sand.',
    'The archive is mapped. The snakes remain outside the scope of this review.',
    'The archive has been inventoried. Its original purpose remains disputed.',
    'Three expeditions preceded this one. Their reports are in this folder.',
    'The oldest folder contains a newer folder named Old.',
    'The expedition requests that Final_v7_REALLYFINAL remain undisturbed.')
if ($UseVT) { Write-ColorLine $ClosingRemark '210;120;255' Magenta }
else { Write-ColorLine $ClosingRemark (Get-Random -InputObject $ClosingRgbColors) Magenta }
Write-Host ''

 } catch {
    # Print the actual error as ordinary wrapped text, avoiding the host's
    # shortened source-line error display. Retain a failing process exit code.
    Write-Host ''
    Write-WrappedMessage ('Folder Review stopped: ' + $_.Exception.Message) Red
    exit 1
} finally {
    if ($CursorHidden) {
        Write-Host -NoNewline "${Esc}[0m${Esc}[?25h"
    }
}