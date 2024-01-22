<#
.SYNOPSIS
Extracts detailed information about video files using MediaInfo CLI, applies user-defined filters based on specified criteria, and optionally copies the files to a target destination.

.DESCRIPTION
This script searches a designated folder and its subfolders for video files, utilizing MediaInfo CLI to extract comprehensive details. It enables filtering based on user-defined criteria, including format, bitrate, resolution, etc. Additionally, it offers an option to copy the filtered video files to a specified destination.

.PARAMETER FolderPath
Specifies the path to the root folder containing video files.

.PARAMETER MediaInfocliPath
Specifies the path to the MediaInfo CLI executable. Default: "C:\Program Files\MediaInfo_CLI\MediaInfo.exe".

.PARAMETER Recursive
If specified, the script searches for video files recursively in subfolders.

.PARAMETER FormatFilter
Specifies the desired video format to filter by.

.PARAMETER FormatNotFilter
Specifies the desired video format to not filter by.

.PARAMETER MinBitrate
Specifies the minimum video bitrate (bps) to filter by.

.PARAMETER MaxBitrate
Specifies the maximum video bitrate (bps) to filter by.

.PARAMETER MinFileSize
Specifies the minimum file size (bytes) to filter by.

.PARAMETER MaxFileSize
Specifies the maximum file size (bytes) to filter by.

.PARAMETER MinWidth
Specifies the minimum video width (pixels) to filter by.

.PARAMETER MaxWidth
Specifies the maximum video width (pixels) to filter by.

.PARAMETER ExactWidth
Specifies the exact video width (pixels) to filter by.

.PARAMETER MinHeight
Specifies the minimum video height (pixels) to filter by.

.PARAMETER MaxHeight
Specifies the maximum video height (pixels) to filter by.

.PARAMETER ExactHeight
Specifies the exact video height (pixels) to filter by.

.PARAMETER EncoderFilter
Specifies a keyword to filter video encoders.

.PARAMETER FileNameFilter
Specifies a keyword to filter file names.

.PARAMETER TargetDestination
Specifies the target destination to copy the filtered video files.

.PARAMETER CopyRelatedFiles
If provided, instructs the script to copy jpg and png images with the same File BaseName as the video to the target.

.EXAMPLE
.\ProcessVideos.ps1 -FolderPath "C:\Videos" -FormatFilter "mp4" -MinBitrate 1000000 -TargetDestination "D:\FilteredVideos"

This example searches for MP4 video files in the "C:\Videos" folder and its subfolders, with a minimum bitrate of 1 Mbps.
The filtered videos are then copied to the "D:\FilteredVideos" directory.
#>
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $FolderPath,

    [Parameter(Position = 1)]
    [string] $MediaInfocliPath = "C:\Program Files\MediaInfo_CLI\MediaInfo.exe",

    [Parameter()]
    [switch] $Recursive,

    [Parameter()]
    [string] $FormatFilter,
    
    [Parameter()]
    [string] $FormatNotFilter,

    [Parameter()]
    [int] $MinBitrate,

    [Parameter()]
    [int] $MaxBitrate,
 
    [Parameter()]
    [int] $MinFileSize,

    [Parameter()]
    [int] $MaxFileSize,
    
    [Parameter()]
    [int] $MinWidth,
    
    [Parameter()]
    [int] $MaxWidth,
    
    [Parameter()]
    [int] $ExactWidth,

    [Parameter()]
    [int] $MinHeight,

    [Parameter()]
    [int] $MaxHeight,

    [Parameter()]
    [int] $ExactHeight,

    [Parameter()]
    [string] $AudioLanguageFilter,
 
    [Parameter()]
    [string] $AudioLanguageNotFilter,

    [Parameter()]
    [string] $EncoderFilter,
 
    [Parameter()]
    [string] $EncoderNotFilter,

    [Parameter()]
    [string] $FileNameFilter,

    [Parameter()]
    [string] $TargetDestination,
    
    [Parameter()]
    [switch] $CopyRelatedFiles
)

# Check if MediaInfo executable exists
if (-not (Test-Path $MediaInfocliPath)) {
    Write-Host "Error: MediaInfo CLI executable not found at the specified path: $MediaInfocliPath"
    Write-Host "Please provide the correct path to the MediaInfo CLI executable using the -MediaInfocliPath parameter."
    Exit
}

<#
.SYNOPSIS
    Converts bitrate from bits per second to a human-readable format.
.DESCRIPTION
    This function takes a bitrate value in bits per second and converts it to a
    more human-readable format. It categorizes the bitrate into Mbps, Kbps, or
    displays it in bits per second, based on the magnitude of the input value.
.PARAMETER bitratePerSecond
    Specifies the bitrate value in bits per second that needs to be converted.
    This parameter is mandatory and can accept input from the pipeline.
.INPUTS
    System.Double. Bitrate values in bits per second.
.OUTPUTS 
    System.String. The converted bitrate value along with the appropriate unit
    (Mbps, Kbps, or b/s).
.EXAMPLE
    Convert-BitRate -bitratePerSecond 1500000
    # Output: '1.50 Mb/s'
    Description: Converts the bitrate value 1500000 b/s to Mbps.
.EXAMPLE
    7500 | Convert-BitRate
    # Output: '7.50 Kb/s'
    Description: Converts the piped-in bitrate value 7500 b/s to Kbps.
.EXAMPLE
    Convert-BitRate -bitratePerSecond 500
    # Output: '500 b/s'
    Description: Displays the bitrate value 500 b/s in bits per second.
#>

function Convert-BitRate {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$bitratePerSecond
    )

    switch ($bitratePerSecond) {
        { $_ -ge 1000000 } {
            # Convert to Mb/s
            '{0:N2} Mb/s' -f ($bitratePerSecond / 1000000)
            break
        }
        { $_ -ge 1000 } {
            # Convert to Kb/s
            '{0:N2} Kb/s' -f ($bitratePerSecond / 1000)
            break
        }
        default {
            # Display in bits if less than 1 bits/s
            "$bitratePerSecond b/s"
        }
    }
}

# Function to extract video information using MediaInfo CLI
function Get-VideoInfo($filePath, $MediaInfocliPath) {
    <#
    .SYNOPSIS
    Retrieves detailed information about a video file using MediaInfo CLI.

    .DESCRIPTION
    This function takes a video file path and the path to the MediaInfo CLI executable as inputs.
    It uses MediaInfo to extract information about the video, such as Format, dimensions, bitrate, and encoder.

    .PARAMETER filePath
    Specifies the path to the video file for which information needs to be extracted.

    .PARAMETER MediaInfocliPath
    Specifies the path to the MediaInfo executable.

    .EXAMPLE
    Get-VideoInfo -filePath "C:\Videos\video.mp4" -MediaInfocliPath "C:\Program Files\FFmpeg\MediaInfo.exe"
    This example retrieves information about the video file "video.mp4" using MediaInfo.

    .NOTES
    This function requires MediaInfo to be installed on the system and the MediaInfocliPath parameter to point to its location.
    #>
    $MediaInfoOutput = & $MediaInfocliPath --output=JSON --Full "$filePath" | ConvertFrom-Json

    $singleVideoInfo = $null

    $generalTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'General' }
    $videoTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'Video' }

    # Array to hold audio languages
    $languages = @()

    # Loop through each audio stream and extract language
    foreach ($stream in $MediaInfoOutput.media.track) {
        if ($stream.StreamKind -eq "Audio" -and $null -ne $stream.Language) {
            $languages += $stream.Language.ToUpper()
        }
    }

    # Join the languages into a string or keep it empty if no languages found
    if ($languages.Count -gt 0) {
        $audioLanguages = $languages -join ' | '
    } else {
        $audioLanguages = ""
    }
    
    $format = $videoTrack.Format_String
    $codec = $videoTrack.CodecID
    $videoWidth = if ($videoTrack.Width) {
        [int]$videoTrack.Width 
    } else {
        $null 
    }
    $videoHeight = if ($videoTrack.Height) {
        [int]$videoTrack.Height 
    } else {
        $null 
    }
    
    if ($videoTrack.BitRate) {
        $rawVideoBitRate = [int]$videoTrack.BitRate
        $videoBitRate = Convert-BitRate -bitratePerSecond $rawVideoBitRate
    } else {
        $videoBitRate = $null
    }
    
    if ($generalTrack.OverallBitRate) {
        $rawTotalBitRate = [int]$generalTrack.OverallBitRate
        $totalBitRate = Convert-BitRate -bitratePerSecond $rawTotalBitRate
    } else {
        $totalBitRate = $null
    }
    $encodedApplication = $generalTrack.Encoded_Application_String
    
    if ($videoTrack.Duration) {
        # Extracting the duration
        $rawDuration = [decimal]$videoTrack.Duration
    } else {
        $rawDuration = [decimal]$generalTrack.Duration
    }
    # Rounding video duration
    $videoDuration = [math]::Floor($rawDuration)

    $FileInfo = Get-Item -LiteralPath $filePath

    $singleVideoInfo = [PSCustomObject]@{
        ParentFolder    = $FileInfo.Directory.FullName
        FileName        = $FileInfo.BaseName
        FullPath        = $filePath
        Format          = $format
        Codec           = $codec
        VideoWidth      = $videoWidth
        VideoHeight     = $videoHeight
        VideoBitrate    = $videoBitRate
        TotalBitrate    = $totalBitRate
        FileSize        = $(Format-Size -SizeInBytes $FileInfo.Length)
        AudioLanguages  = $audioLanguages
        FileSizeByte    = $FileInfo.Length
        VideoDuration   = $VideoDuration   
        Encoder         = $encodedApplication
        RawVideoBitrate = $rawVideoBitRate   
        RawTotalBitrate = $rawTotalBitRate  
    }
    return $singleVideoInfo
}

function Format-Size() {
    <#
    .SYNOPSIS
    Takes bytes and converts it to KB.MB,GB,TB,PB
    
    .DESCRIPTION
    Takes bytes and converts it to KB.MB,GB,TB,PB
    
    .PARAMETER SizeInBytes
    Input bytes
    
    .EXAMPLE
    Format-Size -SizeInBytes 864132
    843,88 KB
    	
    Format-Size -SizeInBytes 8641320
    8,24 MB
    
    .NOTES
    General notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$SizeInBytes
    )
    switch ([math]::Max($SizeInBytes, 0)) {
        { $_ -ge 1PB } {
            '{0:N2} PB' -f ($SizeInBytes / 1PB); break
        }
        { $_ -ge 1TB } {
            '{0:N2} TB' -f ($SizeInBytes / 1TB); break
        }
        { $_ -ge 1GB } {
            '{0:N2} GB' -f ($SizeInBytes / 1GB); break
        }
        { $_ -ge 1MB } {
            '{0:N2} MB' -f ($SizeInBytes / 1MB); break
        }
        { $_ -ge 1KB } {
            '{0:N2} KB' -f ($SizeInBytes / 1KB); break
        }
        default {
            "$SizeInBytes Bytes"
        }
    }
}

# Recursive function to search for video files and extract information
function Get-VideosRecursively($folderPath, $MediaInfocliPath) {
    <#
    .SYNOPSIS
    Recursively searches for video files in a folder and its subfolders and extracts information using MediaInfo.

    .DESCRIPTION
    This function searches for video files (with extensions mp4, mkv, avi, mov, wmv) in the specified folder
    and its subfolders. For each video file found, it calls the Get-VideoInfo function to extract detailed information.

    .PARAMETER folderPath
    Specifies the path of the folder to start the search from.

    .PARAMETER MediaInfocliPath
    Specifies the path to the MediaInfo executable.

    .EXAMPLE
    Get-VideosRecursively -folderPath "C:\Videos" -MediaInfocliPath "C:\Program Files\FFmpeg\MediaInfo.exe"
    This example searches for video files in the "C:\Videos" folder and its subfolders and extracts information using MediaInfo.

    .NOTES
    This function requires the Get-VideoInfo function and MediaInfo to be installed on the system.
    #>
    $videoFiles = @()
    if ($Recursive) {
        $videoFiles = Get-ChildItem -Recurse -Path $folderPath -File | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }
    } else {
        $videoFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }
    }

    $totalFilesToScan = ($videoFiles).Count
    $FilesScanned = 0

    $allVideoInfo = @()
    foreach ($file in $videoFiles) {
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $($FilesScanned + 1) of $totalFilesToScan" -Status "Reading media info: $file" -PercentComplete $progressPercent
        $videoInfo = Get-VideoInfo $file.FullName $MediaInfocliPath
        if ($videoInfo) {
            $allVideoInfo += $videoInfo
        }
        $FilesScanned++
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $FilesScanned of $totalFilesToScan" -Status "Reading media info: $file" -PercentComplete $progressPercent
    }
    return $allVideoInfo
}
function AddFilterCriteria($name, $value) {
    <#
    .SYNOPSIS
    Adds a filter criteria to the global criteria list based on the provided name and value.

    .DESCRIPTION
    This function constructs filter criteria based on user-provided parameters (name and value) and appends them to the global criteria list for filtering video files.

    .PARAMETER name
    Specifies the name of the filter criteria to be added.

    .PARAMETER value
    Specifies the value associated with the filter criteria.

    .EXAMPLE
    AddFilterCriteria -name "FormatFilter" -value "mp4"
    This example adds a filter criteria "FormatFilter=mp4" to the global criteria list.

    .NOTES
    This function aids in dynamically constructing filtering criteria for video files based on user-defined parameters.
    #>
    if ($value) {
        $global:criteria += "$name=$value"
    }
}

#* Start of script
Clear-Host

# Start searching for video files and extracting information
$videoInfoList = Get-VideosRecursively $FolderPath $MediaInfocliPath

# Create filter description based on applied criteria
$criteria = @()

AddFilterCriteria "FormatFilter" $FormatFilter
AddFilterCriteria "FormatNotFilter" $FormatNotFilter
AddFilterCriteria "MinBitrate" $MinBitrate
AddFilterCriteria "MaxBitrate" $MaxBitrate
AddFilterCriteria "MinFileSize" $MinFileSize
AddFilterCriteria "MaxFileSize" $MaxFileSize
AddFilterCriteria "MinWidth" $MinWidth
AddFilterCriteria "MaxWidth" $MaxWidth
AddFilterCriteria "MinHeight" $MinHeight
AddFilterCriteria "MaxHeight" $MaxHeight
AddFilterCriteria "ExactWidth" $ExactWidth
AddFilterCriteria "ExactHeight" $ExactHeight
AddFilterCriteria "AudioLanguageFilter" $AudioLanguageFilter
AddFilterCriteria "AudioLanguageNotFilter" $AudioLanguageNotFilter
AddFilterCriteria "EncoderFilter" $EncoderFilter
AddFilterCriteria "EncoderNotFilter" $EncoderNotFilter
AddFilterCriteria "FileNameFilter" $FileNameFilter

if ($criteria -eq $null -or $criteria.Length -eq 0) {
    $filterDescription = " - No filters applied"
} else {
    $filterDescription = " - Filters: " + ($criteria -join ", ")
}

# Filter based on provided criteria
$videoInfoList = $videoInfoList | Where-Object {
    (!$FormatFilter -or $_.Format -eq $FormatFilter) -and
    (!$FormatNotFilter -or $_.Format -ne $FormatNotFilter) -and
    (!$MinBitrate -or $_.RawTotalBitrate -ge $MinBitrate) -and
    (!$MaxBitrate -or $_.RawTotalBitrate -le $MaxBitrate) -and
    (!$MinFileSize -or $_.FileSizeByte -ge $MinFileSize) -and
    (!$MaxFileSize -or $_.FileSizeByte -le $MaxFileSize) -and
    (!$MinWidth -or $_.VideoWidth -ge $MinWidth) -and
    (!$MaxWidth -or $_.VideoWidth -le $MaxWidth) -and
    (!$MinHeight -or $_.VideoHeight -ge $MinHeight) -and
    (!$MaxHeight -or $_.VideoHeight -le $MaxHeight) -and
    (!$ExactWidth -or $_.VideoWidth -eq $ExactWidth) -and
    (!$ExactHeight -or $_.VideoHeight -eq $ExactHeight) -and
    (!$AudioLanguageFilter -or $_.AudioLanguages -like "*$AudioLanguageFilter*") -and
    (!$AudioLanguageNotFilter -or $_.AudioLanguages -notlike "*$AudioLanguageNotFilter*") -and
    (!$EncoderFilter -or $_.Encoder -like "*$EncoderFilter*") -and
    (!$EncoderNotFilter -or $_.Encoder -notlike "*$EncoderNotFilter*") -and
    (!$FileNameFilter -or $_.FileName -like "*$FileNameFilter*")
}

$sortedVideoInfo = @()
$sortedVideoInfo += $videoInfoList | Sort-Object -Property Format, @{Expression = "VideoWidth"; Descending = $true }, @{Expression = "RawTotalBitrate"; Descending = $true }
$sortedVideoInfo | Select-Object -Property ParentFolder, FileName, Format, VideoWidth, VideoHeight, VideoBitrate, TotalBitrate, RawTotalBitrate, FileSize, FileSizeByte, AudioLanguages, Encoder | Out-GridView -Title "Video information$filterDescription"

# Copy files to the target destination if specified
if ($TargetDestination) {
    $response = Read-Host "Do you want to start the copy of the videos listed? (Y/N)"

    if ($response -eq 'Y' -or $response -eq 'y') {
        $totalFiles = $sortedVideoInfo.Count
        $copiedFiles = 0

        foreach ($videoInfo in $sortedVideoInfo) {
            $sourceFilePath = $videoInfo.FullPath
            $relativePath = $videoInfo.FullPath.Substring($FolderPath.Length)
            $destinationFilePath = Join-Path $TargetDestination $relativePath

            # Create the destination directory if it doesn't exist
            $destinationDirectory = Split-Path -Path $destinationFilePath -Parent
            if (-not (Test-Path $destinationDirectory)) {
                $null = New-Item -ItemType Directory -Path $destinationDirectory
            }
        
            # Write Progress
            $progressPercent = ($copiedFiles / $totalFiles) * 100
            Write-Progress -Activity "Copying Files: $($copiedFiles + 1) of $totalFiles" -Status "Copying $sourceFilePath" -PercentComplete $progressPercent

            # Copy the file to the destination
            $null = Copy-Item -LiteralPath $sourceFilePath -Destination $destinationFilePath -Force

            if ($CopyRelatedFiles) {
                # Check for other files with different extensions but the same BaseName
                $otherExtensions = @("jpg", "png")  # Add more extensions as needed
                $videoFolder = Split-Path -Path $sourceFilePath -Parent
                foreach ($ext in $otherExtensions) {
                    $otherFilePath = Join-Path $videoFolder "$($videoInfo.FileName).$ext"
                    if (Test-Path $otherFilePath) {
                        $otherRelativePath = $otherFilePath.Substring($FolderPath.Length)
                        $otherDestinationFilePath = Join-Path $TargetDestination $otherRelativePath
                
                        # Create the destination directory if it doesn't exist
                        $otherDestinationDirectory = Split-Path -Path $otherDestinationFilePath -Parent
                        if (-not (Test-Path $otherDestinationDirectory)) {
                            $null = New-Item -ItemType Directory -Path $otherDestinationDirectory
                        }
                
                        $null = Copy-Item -LiteralPath $otherFilePath -Destination $otherDestinationFilePath -Force
                    }
                }
            }
        
            $copiedFiles++

            # Write Progress
            $progressPercent = ($copiedFiles / $totalFiles) * 100
            Write-Progress -Activity "Copying Files: $($copiedFiles) of $totalFiles" -Status "Copying $sourceFilePath" -PercentComplete $progressPercent
        }
    } elseif ($response -eq 'N' -or $response -eq 'n') {
        Write-Host "Run script again with different criteria"
    }
}
