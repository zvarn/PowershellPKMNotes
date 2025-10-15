# MIT License

# Copyright (c) 2025 Zach Varnadore

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.


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

#===================================
#---- INITIALIZATION ----===============================================================================
#===================================

# Check for existence of notes root. Create it if it does not exist.
if (-not (Test-Path $script:NOTESROOT)) {
    New-Item -ItemType "Directory" $script:NOTESROOT
}

#====================================
#---- PRIVATE UTILITY FUNCTIONS ----==============================================================================
#====================================

function _IsValidNoteTypeName {
    param (
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Mandatory=$true, Position=0)]
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
        [string]$TypeName
    )

    foreach ($validNoteType in $ValidNoteTypes) {
        if (($TypeName -eq $validNoteType.Type) -or ($TypeName -eq $validNoteType.ShortType)) {
            return $validNoteType.Extension
        }
    }

    return $null;
}

function _WriteError {
    param (
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Mandatory=$true, Position=0)]
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
        [Parameter(Position=0)]
        [string]$FindPattern,

        [string]$Type,

        [string]$SortBy = "Name"
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

    # Handle regex-specific characters
    $wildcardReplacedFindPattern = $FindPattern.replace("+", "\+")

    $allNotes = (Get-ChildItem -Recurse -File $script:NOTESROOT)
    if ($null -eq $allNotes) {
        _WriteError "You have no notes!"
        return;
    }

    if ($null -ne $extension) {
        $allNotes = $allNotes | Where-Object {$_.Extension -match $extension}
    }
    if ($null -eq $allNotes -or $allNotes.Length -eq 0) {
        return;
    }

    if ($SortBy -in $script:ReversedOrderProperties) {
        $allNotes = $allNotes | Sort-Object -Descending -Property {$_.$SortBy}
    } else {
        $allNotes = $allNotes | Sort-Object -Property {$_.$SortBy}
    }

    if ("" -eq $wildcardReplacedFindPattern) {
        # Jump straight to fuzzy search
        $foundNote = $allNotes |
            Select-Object -ExpandProperty FullName |
            fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom'
        if ($null -ne $foundNote -and "" -ne $foundNote) {
            $noteFile = Get-Item $foundNote
            $script:CachedNoteSelectionFiles = @($noteFile)

            Write-Host -NoNewline "[1] "
            Write-Host -ForegroundColor Cyan $noteFile.Name
        }
    }
    else {
        $matchesFound = $allNotes | Where-Object {$_.Name -match $wildcardReplacedFindPattern}

        if ($null -eq $matchesFound) {
            # No match
        }
        elseif ($matchesFound.Length -eq 0) {
            $script:CachedNoteSelectionFiles = @($matchesFound)
            Write-Host -NoNewline "[1] "
            Write-Host -ForegroundColor Cyan $matchesFound.Name
        }
        elseif ($matchesFound.Length -lt $script:NOTESLISTMAXENTRIES) {
            _DoMultiPageListExperience $matchesFound
        }
        else {
            $foundNote = $matchesFound | 
                Select-Object -ExpandProperty FullName |
                fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom'
            if ($null -ne $foundNote -and "" -ne $foundNote) {
                $noteFile = Get-Item $foundNote
                $script:CachedNoteSelectionFiles = @($noteFile)

                Write-Host -NoNewline "[1] "
                Write-Host -ForegroundColor Cyan $noteFile.Name
            }
        }
    }
}

function New-Note {
    param(
        [Parameter(Mandatory=$true, Position=0)]
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

    New-Item $noteFullPath -Type File | Out-Null
    Write-Host -NoNewline -ForegroundColor Green "Note "
    Write-Host -NoNewline -ForegroundColor Cyan "$noteNameWithExtension"
    Write-Host -ForegroundColor Green " created!"

    Invoke-Expression "& $script:OPENNOTECOMMAND $noteFullPath"
}

function Remove-Note {
    param(
        [Parameter(Position=0)]
        [string]$NoteName,

        [string]$Type,

        [string]$SortBy = "Name"
    )

    $noteFullPath = ""
    $noteNameWithExtension = ""

    # Check for cached note selection
    if ($NoteName.Length -gt 1 -and $NoteName[0] -eq '/') {
        if ($null -eq $script:CachedNoteSelectionFiles -or 0 -eq $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Note cache currently does not exist!"
            return;
        }

        $indexSelected = 0
        try {
            $indexSelected = ([int]$NoteName.Substring(1)) - 1
        } catch {
            _WriteError "Invalid note cache selection."
            return;
        }

        if ($indexSelected -lt 0 -or $indexSelected -ge $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Invalid note cache selection."
            return;
        }

        $noteFullPath = $script:CachedNoteSelectionFiles[$indexSelected].FullName
        $noteNameWithExtension = $script:CachedNoteSelectionFiles[$indexSelected].Name
    }
    else {
        # Validate parameters
        if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
            _WriteError "Invalid note type: $Type"
            return;
        }

        $allNotes = (Get-ChildItem -Recurse -File $script:NOTESROOT)
        if ($null -eq $allNotes) {
            _WriteError "You have no notes!"
            return;
        }

        if ($null -ne $extension) {
            $allNotes = $allNotes | Where-Object {$_.Extension -match $extension}
        }
        if ($null -eq $allNotes -or $allNotes.Length -eq 0) {
            return;
        }

        if ($SortBy -in $script:ReversedOrderProperties) {
            $allNotes = $allNotes | Sort-Object -Descending -Property {$_.$SortBy}
        } else {
            $allNotes = $allNotes | Sort-Object -Property {$_.$SortBy}
        }

        # Locate the note - fuzzy find if no exact note given from args
        if ($null -eq $NoteName -or "" -eq $NoteName) {
            # Jump right into fuzzy find.
            $noteFullPath = $allNotes |
                Select-Object -ExpandProperty FullName |
                fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom'
            if ($null -eq $noteFullPath -or "" -eq $noteFullPath) {
                _WriteError "No note specified"
                return;
            }

            $noteNameWithExtension = ($noteFullPath | Select-String ".+\\(.+)$").Matches[0].Groups[1].Value
        }
        else {
            if (_HasValidNoteTypeExtension $NoteName) {
                $noteNameWithExtension = $NoteName
                $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NoteName)
                $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
            }
            else {
                if ("" -ne $Type) {
                    $noteNameWithExtension = "$NoteName.$(_GetExtensionForValidNoteTypeName $Type)"
                    $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
                    $noteFullPath = "$script:NOTESROOT\$notePartialPath"
                } else {
                    # We don't have a complete note name, just setting up for fuzzy find.
                    $noteNameWithExtension = $NoteName
                }
            }

            if ("" -eq $noteFullPath -or !(Test-Path $noteFullPath)) {
                # Start fuzzy find to find note
                $noteFullPath = $allNotes |
                    Select-Object -ExpandProperty FullName |
                    fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom' --query "$noteNameWithExtension"
                if ($null -eq $noteFullPath -or "" -eq $noteFullPath) {
                    _WriteError "No note specified"
                    return;
                }
        
                $noteNameWithExtension = ($noteFullPath | Select-String ".+\\(.+)$").Matches[0].Groups[1].Value
            }
        }
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
    # Sort by FullName length (longest to shortest) to work up the directory tree
    $allDirs = (Get-ChildItem -Recurse -Directory $script:NOTESROOT | Sort-Object -Descending {$_.FullName.Length})
    foreach ($d in $allDirs) {
        if ($d.GetFileSystemInfos().Count -eq 0) {
            Remove-Item $d
        }
    }
}

function Open-Note {
    param(
        [Parameter(Position=0)]
        [string]$NoteName,

        [string]$Type,

        [string]$SortBy = "LastAccessTime"
    )

    $noteFullPath = ""

    # Check for cached note selection
    if ($NoteName.Length -gt 1 -and $NoteName[0] -eq '/') {
        if ($null -eq $script:CachedNoteSelectionFiles -or 0 -eq $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Note cache currently does not exist!"
            return;
        }

        $indexSelected = 0
        try {
            $indexSelected = ([int]$NoteName.Substring(1)) - 1
        } catch {
            _WriteError "Invalid note cache selection."
            return;
        }

        if ($indexSelected -lt 0 -or $indexSelected -ge $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Invalid note cache selection."
            return;
        }

        $noteFullPath = $script:CachedNoteSelectionFiles[$indexSelected].FullName
    }
    else {
        # Validate parameters
        if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
            _WriteError "Invalid note type: $Type"
            return;
        }

        $allNotes = (Get-ChildItem -Recurse -File $script:NOTESROOT)
        if ($null -eq $allNotes) {
            _WriteError "You have no notes!"
            return;
        }

        if ($null -ne $extension) {
            $allNotes = $allNotes | Where-Object {$_.Extension -match $extension}
        }
        if ($null -eq $allNotes -or $allNotes.Length -eq 0) {
            return;
        }

        if ($SortBy -in $script:ReversedOrderProperties) {
            $allNotes = $allNotes | Sort-Object -Descending -Property {$_.$SortBy}
        } else {
            $allNotes = $allNotes | Sort-Object -Property {$_.$SortBy}
        }

        # Locate the note - fuzzy find if no exact note given from args
        if ($null -eq $NoteName -or "" -eq $NoteName) {
            # Jump right into fuzzy find.
            $noteFullPath = $allNotes |
                Select-Object -ExpandProperty FullName |
                fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom'
            if ($null -eq $noteFullPath -or "" -eq $noteFullPath) {
                _WriteError "No note specified"
                return;
            }
        }
        else {
            $noteNameWithExtension = ""
            if (_HasValidNoteTypeExtension $NoteName) {
                $noteNameWithExtension = $NoteName
                $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NoteName)
                $noteFullPath = "$script:NOTESROOT\$notePartialPath" 
            }
            else {
                if ("" -ne $Type) {
                    $noteNameWithExtension = "$NoteName.$(_GetExtensionForValidNoteTypeName $Type)"
                    $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
                    $noteFullPath = "$script:NOTESROOT\$notePartialPath"
                } else {
                    # We don't have a complete note name, just setting up for fuzzy find.
                    $noteNameWithExtension = $NoteName
                }
            }

            if ("" -eq $noteFullPath -or !(Test-Path $noteFullPath)) {
                # Start fuzzy find to find note
                $noteFullPath = $allNotes |
                    Select-Object -ExpandProperty FullName |
                    fzf --delimiter "\" --with-nth=-1 --preview "bat --color=always --style=numbers --line-range=:500 {}" --preview-window 'up,60%,border-bottom' --query "$noteNameWithExtension"
                if ($null -eq $noteFullPath -or "" -eq $noteFullPath) {
                    _WriteError "No note specified"
                    return;
                }
            }
        }
    }

    Invoke-Expression "& $script:OPENNOTECOMMAND $noteFullPath"
}

function Rename-Note {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$CurrentNoteName,

        [Parameter(Mandatory=$true, Position=1)]
        [string]$NewNoteName,

        [string]$Type
    )

    $currentNoteFullPath = ""

    if ($CurrentNoteName.Length -gt 1 -and $CurrentNoteName[0] -eq '/') {
        if ($null -eq $script:CachedNoteSelectionFiles -or 0 -eq $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Note cache currently does not exist!"
            return;
        }

        $indexSelected = 0
        try {
            $indexSelected = ([int]$CurrentNoteName.Substring(1)) - 1
        } catch {
            _WriteError "Invalid note cache selection 1."
            return;
        }

        if ($indexSelected -lt 0 -or $indexSelected -ge $script:CachedNoteSelectionFiles.Length) {
            _WriteError "Invalid note cache selection 2."
            return;
        }

        $currentNoteFullPath = $script:CachedNoteSelectionFiles[$indexSelected].FullName
    } else {
        if ("" -ne $Type -and !(_IsValidNoteTypeName $Type)) {
            _WriteError "Invalid note type: $Type"
            return;
        }

        if (_HasValidNoteTypeExtension $CurrentNoteName) {
            $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $CurrentNoteName)
            $currentNoteFullPath = "$script:NOTESROOT\$notePartialPath" 
        }
        else {
            if ("" -ne $Type) {
                $noteNameWithExtension = "$CurrentNoteName.$(_GetExtensionForValidNoteTypeName $Type)"
                $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
                $CurrentNoteName = "$script:NOTESROOT\$notePartialPath"
            }
        }
    }

    if (-not (Test-Path $currentNoteFullPath)) {
        _WriteError "Current note specified does not exist!"
        return;
    }

    $newNoteFullPath = ""
    if (-not (_HasValidNoteTypeExtension $NewNoteName)) {
        $noteNameWithExtension = "$NewNoteName.$(_GetExtensionForValidNoteTypeName $Type)"
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $noteNameWithExtension)
        $newNoteFullPath = "$script:NOTESROOT\$notePartialPath" 
    } else {
        $notePartialPath = (_GetNotePathFormattedFromNoteNameWithExtension $NewNoteName)
        $newNoteFullPath = "$script:NOTESROOT\$notePartialPath" 
    }

    # Create destination path, or fail if it already exists - user must do explicit delete of destination note if it
    # already exists.
    if (Test-Path $newNoteFullPath) {
        _WriteError "Existing note with new note name already exists!"
        return;
    }

    New-Item -Force -ItemType "File" $newNoteFullPath | Out-Null

    Move-Item -Force -Path $currentNoteFullPath -Destination $newNoteFullPath
}

function Search-Notes {
    $RgPrefix="rg --column --line-number --no-heading --color=always --smart-case "

    $allNotes = (Get-ChildItem -Recurse -File $script:NOTESROOT)
    if ($null -eq $allNotes) {
        _WriteError "You have no notes!"
        return;
    }

    # Remove old query cache files
    if (Test-Path $USERGLOBALTEMP\notes-search-f) {
        Remove-Item -Force $USERGLOBALTEMP\notes-search-f
    }
    if (Test-Path $USERGLOBALTEMP\notes-search-r) {
        Remove-Item -Force $USERGLOBALTEMP\notes-search-r
    }

    $fzfResult = fzf --ansi --disabled --with-nth=-1 `
        --bind "start:reload($RgPrefix `"`" `"$script:NOTESROOT`")+unbind(ctrl-r)" `
        --bind "change:reload:powershell -Command `"Start-Sleep -Milliseconds 100`"; $RgPrefix {q} '$script:NOTESROOT'" `
        --bind "ctrl-f:unbind(change,ctrl-f)+change-prompt(2. fzf> )+enable-search+rebind(ctrl-r)+execute(echo {q} > $USERGLOBALTEMP\notes-search-r)+transform-query:powershell -Command `"(Get-Content $USERGLOBALTEMP\notes-search-f) -replace('^^\`"', '') -replace(' $', '')`"" `
        --bind "ctrl-r:unbind(ctrl-r)+change-prompt(1. rg> )+disable-search+reload($RgPrefix {q} `"$script:NOTESROOT`")+rebind(change,ctrl-f)+execute-silent(powershell -Command `"Set-Content $USERGLOBALTEMP\notes-search-f {q}`")+transform-query:powershell -Command `"(Get-Content $USERGLOBALTEMP\notes-search-r) -replace('\`"', '') -replace(' ', '')`"" `
        --color "hl:-1:underline,hl+:-1:underline:reverse" `
        --prompt '1. rg> ' `
        --delimiter "\" `
        --header '-- CTRL-R (rg mode), CTRL-F (fzf mode)' `
        --preview "powershell -Command `"`$splitStr = '{}' -split ':'; bat --color=always (`$splitStr[0][1] + ':' + `$splitStr[1]) --highlight-line `$splitStr[2]`"" `
        --preview-window 'up,60%,border-bottom,+{3}+3/3,~3'

    if ($null -ne $fzfResult -and "" -ne $fzfResult) {
        # Because the items in fzf were formatted with more than just the file name, we need to truncate it
        $matchGroups = $fzfResult | Select-String "(.+?):[0-9]+:[0-9]+:" | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Groups
        $fileTarget = $matchGroups[1].Value

        $noteFile = Get-Item $fileTarget
        $script:CachedNoteSelectionFiles = @($noteFile)

        Write-Host -NoNewline "[1] "
        Write-Host -ForegroundColor Cyan $noteFile.Name
    }
}

# Main 'notes' function with dispatching for given $Command
function Invoke-Notes {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,

        [string]$SortBy = "Name"
    )

    DynamicParam
    {
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

        if ($Command -eq "find") { # FindPattern required
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $false
                Position = 1
            }

            $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attributeCollection.Add($parameterAttribute)

            $dynParam1 = [System.Management.Automation.RuntimeDefinedParameter]::new(
                'FindPattern', [String], $attributeCollection
            )
            
            $paramDictionary.Add('FindPattern', $dynParam1)
        }
        elseif ($Command -eq "create" -or $Command -eq "delete" -or $Command -eq "open") {
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $false
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
                    (Find-Note -FindPattern $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy)
                }
                else {
                    (Find-Note $PSBoundParameters.FindPattern -SortBy $SortBy)
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
                    (Remove-Note -NoteName $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy)
                }
                else {
                    (Remove-Note $PSBoundParameters.NoteName -SortBy $SortBy)
                }
            }

            "open" {
                if ($null -ne $PSBoundParameters.Type) {
                    (Open-Note -NoteName $PSBoundParameters.NoteName -Type $PSBoundParameters.Type -SortBy $SortBy)
                }
                else {
                    (Open-Note $PSBoundParameters.NoteName -SortBy $SortBy)
                }
            }

            default
            {
                _WriteError "Invalid notes command: $Command"
            }
        }
    }
}