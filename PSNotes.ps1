#==================================
#---- CONFIGURATION VARIABLES ----================================================================================
#==================================

# Where notes are stored on disk
# Default: $env:USERPROFILE\notes ( = C:\Users\<you>\notes )
$script:NOTESROOT = "$env:USERPROFILE\notes"

# When printing a vertical list of notes, this limits the maximum that will be shown per "page"
# Default: 9
# TODO: Anything 10 or greater currently will not work!
$script:NOTESLISTMAXENTRIES = 9

# Currently, every note will open with vscode (code)
$script:OPENNOTECOMMAND = "code"

# Known file types that the system will recognize as "notes"
$script:ValidNoteTypes = @(
    @{
        Type = "Markdown"
        ShortType = "md"
        Extension = "md"
        ContentSearch = "text"
    },
    @{
        Type = "Jupyter"
        ShortType = "Jupyter"
        Extension = "ipynb"
        ContentSearch = "jupyter"
    }
)

# List of standard Powershell properties to automatically sort in reverse (i.e.- use '-Descending') for a list of
# notes.
# Primarily used to sort date properties like LastAccessTime by most-recent by default.
$script:ReversedOrderProperties = @(
    "CreationTime",
    "CreationTimeUtc",
    "LastAccessTime",
    "LastAccessTimeUtc",
    "LastWriteTime",
    "LastWriteTimeUtc")

#===================================
#---- PRIVATE HELPER VARIABLES ----===============================================================================
#===================================

$script:NotesRootRegexable = $script:NOTESROOT.replace("\", "\\")
$script:CachedNoteSelectionFiles = @()


#====================================
#---- PRIVATE UTILITY FUNCTIONS ----==============================================================================
#====================================

function _IsValidNoteTypeName {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$TypeName
    )

    foreach ($validNoteType in $ValidNoteTypes) {
        if (($TypeName -eq $validNoteType.Type) -or ($TypeName -eq $validNoteType.ShortType)) {
            return $true
        }
    }

    return $false
}

# Defines rules for a valid note name
function _IsValidNoteName {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteName
    )

    if ($NoteName -eq "") {
        return $false
    }

    # Allowed characters
    if (!($NoteName -match "^[-+!@#%&_\w\.]+$")) {
        return $false
    }

    # Must have other characters between '.'s
    if ($NoteName -match "\.\.") {
        return $false
    }

    # Must have other characters before first '.'
    if ($NoteName -match "^\.") {
        return $false
    }

    # Must have other characters after last '.'
    if ($NoteName -match "\.$") {
        return $false
    }

    return $true
}

function _HasValidNoteTypeExtension {
    param(
        [Parameter(Mandatory, Position=0)]
        [string]$NoteName
    )

    $potentialExtension = ($NoteName | sls "\.([^\.]+)$") # Everything after last '.'

    if ($null -eq $potentialExtension) {
        return $false
    }

    $extension = $potentialExtension.Matches[0].Groups[1].Value

    foreach ($type in $ValidNoteTypes) {
        if ($extension -eq $type.Extension) {
            return $true
        }
    }

    return $false
}

function _GetExtensionForValidNoteTypeName {
    param(
        [ValidateScript({_IsValidNoteTypeName $_})]
        [string]$TypeName
    )

    foreach ($validNoteType in $ValidNoteTypes) {
        if (($TypeName -eq $validNoteType.Type) -or ($TypeName -eq $validNoteType.ShortType)) {
            return $validNoteType.Extension
        }
    }

    throw "Unexpected end of function!"
}

# Return note path of everything after $NOTESROOT
function _GetNoteTruncatedPath {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteFullPath
    )

    return ($NoteFullPath | sls "$script:NotesRootRegexable\\(.+)").Matches[0].Groups[1]
}

function _GetNoteNameFromTruncatedPath {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteTruncatedPath
    )

    return $NoteTruncatedPath.replace("\", ".")
}

function _WriteError {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]$ErrorString
    )

    Write-Host -ForegroundColor Red $ErrorString
}

# Only accepts a full note name with its extension included
# Example input: top.mid.leaf.md
# Example output: top/mid/top.mid.leaf.md (i.e.- hierarchy in directories and filename)
function _GetNotePathFormattedFromNoteNameWithExtension {
    param (
        [ValidateScript({(_IsValidNoteName $_) -and (_HasValidNoteTypeExtension $_)})]
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteNameWithExtension
    )

    $spaceToSlashName = $NoteNameWithExtension.replace(".", "\")
    $spaceToSlashName = ($spaceToSlashName -replace "(.*)\\(.*)", '$1.$2') # Replaces last '\' with a dot 

    $pathPrefix = $spaceToSlashName | sls ".+\\" # Note this is greedy matching, so we get last '\'!
    if ($null -eq $pathPrefix) {
        return $NoteNameWithExtension
    }
    
    return ($pathPrefix.Matches[0].Value + $NoteNameWithExtension)
}

function _WriteNumberedListOfNoteNames {
    param (
        [Parameter(Mandatory, Position = 0)]
        [string[]]$NoteNames
    )

    $max = [math]::Min($NoteNames.Length, 9)

    foreach ($i in 1..$max) {
        Write-Host -NoNewline "[$i] "
        Write-Host -ForegroundColor Cyan $NoteNames[$i - 1]
    }
}

# TODO: IMPLEMENT ASSOCIATED CONTENT (e.g.- CONTENT MATCH) DISPLAYING
function _DoMultiPageListExperience {
    param (
        [Parameter(Mandatory, Position=0)]
        [PSObject[]]$Items,

        [string]$PageHeaderPrefix,

        $ItemFilter = $null # TODO: IMPLEMENT FILTER FOR ITEMS IN LIST FOR ENUMERATION
    )

    # Simplist case: Nothing!
    if ($Items.Length -eq 0) {
        return;
    }

    # Simple case: only 1 page of results to show. Show it and return.
    if ($Items.Length -le $script:NOTESLISTMAXENTRIES) {
        $entriesToShow = $Items | Select-Object -ExpandProperty Name
        _WriteNumberedListOfNoteNames $entriesToShow
        $script:CachedNoteSelectionFiles = $Items
        return;
    }

    $numPages = -1 # -1 means we don't know yet
    $numPagesDisplayString = "?"
    if ($null -eq $ItemFilter) {
        $numPages = [int][math]::Ceiling($Items.Length / $script:NOTESLISTMAXENTRIES)
        $numPagesDisplayString = $numPages
    }

    $indexFirstItemCurrentPage = 0
    # $cachedHydratedEntries = @()

    # First page of entries upfront
    $pageEntries = $Items[0..($script:NOTESLISTMAXENTRIES - 1)] 
    # $cachedHydratedEntries += $pageEntries

    Write-Host -ForegroundColor Magenta "$PageHeaderPrefix (Page 1 of $numPagesDisplayString)"

    _WriteNumberedListOfNoteNames ($pageEntries | Select-Object -ExpandProperty Name)
    Write-Host ""
    Write-Host -NoNewline "(N/n) Next page, (Esc/Backspace/Enter) Stop"

    $hasNextPage = $true
    $hasPrevPage = $false
    while ($true) {
        $key = [System.Console]::ReadKey($true).Key

        switch ($key) {
            {@([ConsoleKey]::Escape, [ConsoleKey]::Enter, [ConsoleKey]::Backspace) -contains $_} {
                $script:CachedNoteSelectionFiles = $pageEntries
                return;
            }

            'N' {
                # Next list
                if (!($hasNextPage)) {
                    continue
                }

                $indexFirstItemCurrentPage += $script:NOTESLISTMAXENTRIES

                $remainingItemsCount = $Items.Length - $indexFirstItemCurrentPage
                $pageEndIndexAdd = [math]::Min($remainingItemsCount, 9)

                $pageEntries = $Items[$indexFirstItemCurrentPage..($indexFirstItemCurrentPage + $pageEndIndexAdd - 1)]

                # $cachedHydratedEntries += $pageEntries # TODO: HYDRATION

                $currentPage = ($indexFirstItemCurrentPage / $script:NOTESLISTMAXENTRIES) + 1
                $numPagesDisplayString = $numPagesDisplayString # TODO: ONLY SET IF ALL PAGES HYDRATED WITH FILTER PRESENT

                Write-Host ""
                Write-Host ""
                Write-Host -ForegroundColor Magenta "$PageHeaderPrefix (Page $currentPage of $numPagesDisplayString)"

                _WriteNumberedListOfNoteNames ($pageEntries | Select-Object -ExpandProperty Name)
                Write-Host ""

                $hasPrevPage = $true
                if ($remainingItemsCount -gt $script:NOTESLISTMAXENTRIES) {
                    $hasNextPage = $true
                    Write-Host -NoNewline "(N/n) Next page, (P/p) Previous page, (Esc/Backspace/Enter) Stop"
                } else {
                    $hasNextPage = $false
                    Write-Host -NoNewline "(P/p) Previous page, (Esc/Backspace/Enter) Stop"
                }
            }

            'P' {
                # Previous list
                if (!($hasPrevPage)) {
                    continue
                }

                $indexFirstItemCurrentPage -= $script:NOTESLISTMAXENTRIES

                $pageEntries = $Items[$indexFirstItemCurrentPage..($indexFirstItemCurrentPage + $script:NOTESLISTMAXENTRIES - 1)]
                
                $currentPage = ($indexFirstItemCurrentPage / $script:NOTESLISTMAXENTRIES) + 1

                Write-Host ""
                Write-Host ""
                Write-Host -ForegroundColor Magenta "$PageHeaderPrefix (Page $currentPage of $numPagesDisplayString)"
                
                _WriteNumberedListOfNoteNames ($pageEntries | Select-Object -ExpandProperty Name)
                Write-Host ""

                $hasNextPage = $true
                if ($indexFirstItemCurrentPage -eq 0) {
                    $hasPrevPage = $false
                    Write-Host -NoNewline "(N/n) Next page, (Esc/Backspace/Enter) Stop"
                } else {
                    $hasPrevPage = $true
                    Write-Host -NoNewline "(N/n) Next page, (P/p) Previous page, (Esc/Backspace/Enter) Stop"
                }
            }

            default {
                continue
            }
        }
    }
}

#======================
#---- ENTRY SETUP ----============================================================================================
#======================

# Register 'notes' as a PowerShell alias for 'Invoke-Notes' common entry point
Set-Alias -Name notes -Value Invoke-Notes


#=================================
#---- PUBLIC COMMANDLINE API ----=================================================================================
#=================================

function Find-Note {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$FindPattern,

        [string]$Type,

        [string]$SortBy = "LastAccessTime",
        [switch]$ReverseOrder,

        [switch]$NoPrintHeader
    )

    # Validate parameters
    if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
        _WriteError "Invalid note type: $Type"
        return;
    }

    $extension = $null
    if ($Type -ne "") {
        $extension = (_GetExtensionForValidNoteTypeName $Type)
    }

    # Sub in regex wildcards
    $wildcardReplacedFindPattern = $FindPattern.replace("*", "\S*")

    $allNotes = (Get-ChildItem -Recurse -File $NOTESROOT)

    $matchesFound = $null
    if ($null -ne $extension) {
        $matchesFound = $allNotes | Where-Object {$_.Name -match $wildcardReplacedFindPattern -and $_.Extension -match $extension}
    } else {
        $matchesFound = $allNotes | Where-Object {$_.Name -match $wildcardReplacedFindPattern}
    }

    if ($null -eq $matchesFound) {
        # No match
    }
    elseif ($matchesFound.Length -eq 0) { # 1 match
        if (!($NoPrintHeader)) {
            Write-Host -ForegroundColor Magenta "-- Notes Find Result"
        }

        $script:CachedNoteSelectionFiles = @($matchesFound[0])
        Write-Host -NoNewline "[1] "
        Write-Host -ForegroundColor Cyan $matchesFound.Name
    }
    elseif ($matchesFound.Length -gt 1) {
        if ($SortBy -in $script:ReversedOrderProperties) {
            $matchesFound = $matchesFound | Sort-Object -Descending -Property {$_.$SortBy}
        } else {
            $matchesFound = $matchesFound | Sort-Object -Property {$_.$SortBy}
        }

        _DoMultiPageListExperience $matchesFound -PageHeaderPrefix "-- Notes Find Result"
    }
}

function New-Note {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteName,

        [string]$Type
    )

    # Validate parameters
    if (!(_IsValidNoteName $NoteName)) {
        _WriteError "Invalid note name: $NoteName"
        return;
    }

    if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
        _WriteError "Invalid note type: $Type"
        return;
    }

    $noteFullPath = ""
    $noteNameWithExtension = ""
    if (_HasValidNoteTypeExtension $NoteName) {
        $noteNameWithExtension = $NoteName
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NoteName)
        $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }
    else {
        if ($null -eq $Type -or "" -eq $TYPE) {
            _WriteError "Note missing extension and Type, please provide extension in name or Type!"
            return;
        }

        $noteNameWithExtension = "$NoteName.$(_GetExtensionForValidNoteTypeName $Type)"
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
        $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }

    if (Test-Path $noteFullPath) {
        Write-Host -NoNewline -ForegroundColor Red "Note "
        Write-Host -NoNewline -ForegroundColor Cyan "$noteNameWithExtension"
        Write-Host -ForegroundColor Red " already exists!"
        return;
    }

    New-Item $noteFullPath -Type File -Force | Out-Null
    Write-Host -NoNewline -ForegroundColor Green "Note "
    Write-Host -NoNewline -ForegroundColor Cyan "$noteNameWithExtension"
    Write-Host -ForegroundColor Green " created!"

    Invoke-Expression "& $script:OPENNOTECOMMAND $noteFullPath"
}

function Remove-Note {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteName,

        [string]$Type,

        [string]$SortBy = "LastAccessTime",
        [switch]$ReverseOrder
    )

    # Validate parameters
    if (!(_IsValidNoteName $NoteName)) {
        _WriteError "Invalid note name: $NoteName"
        return;
    }

    if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
        _WriteError "Invalid note type: $Type"
        return;
    }

    $noteFullPath = ""
    $noteNameWithExtension = ""
    if (_HasValidNoteTypeExtension $NoteName) {
        $noteNameWithExtension = $NoteName
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NoteName)
        $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }
    else {
        if ($null -ne $Type -and "" -ne $TYPE) {
            $noteNameWithExtension = "$NoteName.$(_GetExtensionForValidNoteTypeName $Type)"
            $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
            $noteFullPath = "$script:NOTESROOT\$notePartialPath"
        }
    }

    if (!(Test-Path $noteFullPath)) {
        # Attempt to find close notes
        Find-Note -NoPrintHeader -FindPattern "*$NoteName*" -Type "$Type" -SortBy $SortBy -ReverseOrder:$ReverseOrder
        Write-Host ""

        $selectionMax = $script:CachedNoteSelectionFiles.Length
        Write-Host -NoNewline "Press 1-$selectionMax to select match. Press any other key to abort Delete."

        $key = [System.Console]::ReadKey($true).KeyChar
        if ($key -lt "1" -or $key -gt "$selectionMax") {
            # User aborted
            return;
        }

        # Put new line in now
        Write-Host ""

        $selection = $script:CachedNoteSelectionFiles[[int]::Parse($key) - 1]
        $noteFullPath = $selection.Fullname
    }

    # Make sure the user is really Sure about this
    Write-Host -NoNewline "Really DELETE note "
    Write-Host -NoNewline -ForegroundColor Cyan "$noteNameWithExtension"
    Write-Host -NoNewline "? This is NOT reversible! (Y/y to confirm)"
    $key = [System.Console]::ReadKey($true).KeyChar

    if ($key -ne 'Y' -and $key -ne 'y') {
        Write-Host ""
        Write-Host "Note not deleted."
        return;
    }

    Remove-Item -Force $noteFullPath

    Write-Host ""
    Write-Host -NoNewline -ForegroundColor Green "Note "
    Write-Host -NoNewline -ForegroundColor Cyan "$noteNameWithExtension"
    Write-Host -ForegroundColor Green " deleted!"
    
    # Remove any potentially now empty directories
    $emptyDirs = (Get-ChildItem -Recurse -Directory $script:NOTESROOT | Where-Object { $_.GetFileSystemInfos().Count -eq 0 })
    foreach ($d in $emptyDirs) {
        Remove-Item $d -Force -Recurse
    }
}

function Open-Note {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$NoteName,

        [string]$Type,

        [string]$SortBy = "LastAccessTime",
        [switch]$ReverseOrder
    )

    # Validate parameters
    if (!(_IsValidNoteName $NoteName)) {
        _WriteError "Invalid note name: $NoteName"
        return;
    }

    if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
        _WriteError "Invalid note type: $Type"
        return;
    }

    $noteFullPath = ""
    $noteNameWithExtension = ""
    if (_HasValidNoteTypeExtension $NoteName) {
        $noteNameWithExtension = $NoteName
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NoteName)
        $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }
    elseif ("" -ne $Type) {
        $noteNameWithExtension = "$NoteName.$(_GetExtensionForValidNoteTypeName $Type)"
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
        $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }

    if ("" -ne $noteFullPath -and (Test-Path $noteFullPath)) {
        Invoke-Expression "& $script:OPENNOTECOMMAND $noteFullPath"
        return;
    }

    # Attempt to find close notes
    Find-Note -NoPrintHeader -FindPattern "*$NoteName*" -Type "$Type" -SortBy $SortBy -ReverseOrder:$ReverseOrder
    Write-Host ""

    $selectionMax = $script:CachedNoteSelectionFiles.Length
    Write-Host -NoNewline "Press 1-$selectionMax to select match. Press any other key to abort Open."

    $key = [System.Console]::ReadKey($true).KeyChar
    if ($key -lt "1" -or $key -gt "$selectionMax") {
        # User aborted
        return;
    }

    $selection = $script:CachedNoteSelectionFiles[[int]::Parse($key) - 1]
    Invoke-Expression "& $script:OPENNOTECOMMAND $($selection.FullName)"
}

function Search-Notes {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$QueryString,

        [string]$NoteNameFilter


    )
}

# Main 'notes' function with dispatching for given $Command
function Invoke-Notes {
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Command,

        [string]$SortBy = "LastAccessTime",
        [switch]$ReverseOrder
    )

    DynamicParam
    {
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        if ($Command -eq "find") { # FindPattern required
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $true
                Position = 1
            }

            $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attributeCollection.Add($parameterAttribute)

            $dynParam1 = [System.Management.Automation.RuntimeDefinedParameter]::new(
                'FindPattern', [String], $attributeCollection
            )
            
            $paramDictionary.Add('FindPattern', $dynParam1)
        }
        elseif ($Command -eq "create" -or $Command -eq "delete" -or $Command -eq "open") { # NoteName required
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $true
                Position = 1
            }

            $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attributeCollection.Add($parameterAttribute)

            $dynParam1 = [System.Management.Automation.RuntimeDefinedParameter]::new(
                'NoteName', [String], $attributeCollection
            )

            $paramDictionary.Add('NoteName', $dynParam1)
        }

        $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
            Mandatory = $false
        }

        $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
        $attributeCollection.Add($parameterAttribute)

        $dynParam2 = [System.Management.Automation.RuntimeDefinedParameter]::new(
            'Type', [String], $attributeCollection
        )

        $paramDictionary.Add('Type', $dynParam2)

        return $paramDictionary
    }

    process 
    {
        switch ($Command) {
            "find" {
                if ($null -ne $PSBoundParameters.Type) {
                    (Find-Note -FindPattern $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
                else {
                    (Find-Note $PSBoundParameters.FindPattern -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
            }

            "create" {
                if ($null -ne $PSBoundParameters.Type) {
                    (New-Note -NoteName $PSBoundParameters.NoteName -Type $PSBoundParameters.Type)
                }
                else {
                    (New-Note $PSBoundParameters.NoteName)
                }
            }

            "delete" {
                if ($null -ne $PSBoundParameters.Type) {
                    (Remove-Note -NoteName $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
                else {
                    (Remove-Note $PSBoundParameters.NoteName -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
            }

            "open" {
                if ($null -ne $PSBoundParameters.Type) {
                    (Open-Note -NoteName $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
                else {
                    (Open-Note $PSBoundParameters.NoteName -SortBy $SortBy -ReverseOrder:$ReverseOrder)
                }
            }

            default
            {
                _WriteError "Invalid notes command: $Command"
            }
        }
    }
}