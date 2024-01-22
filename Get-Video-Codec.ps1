<#
.SYNOPSIS
    Extracts detailed information about video files using MediaInfo CLI, applies user-defined filters based on specified criteria, and optionally copies the files to a target destination.
.DESCRIPTION
    This script searches a designated folder and its subfolders for video files, utilizing MediaInfo CLI to extract comprehensive details. It enables filtering based on user-defined criteria, including codec, bitrate, resolution, etc. Additionally, it offers an option to copy the filtered video files to a specified destination.
.PARAMETER FolderPath
    Specifies the path to the root folder containing video files.
.PARAMETER MediaInfocliPath
    Specifies the path to the MediaInfo CLI executable. Default: "C:\Program Files\MediaInfo_CLI\MediaInfo.exe".
.PARAMETER Recursive
    If specified, the script searches for video files recursively in subfolders.
.PARAMETER CodecFilter
    Specifies the desired video codec to filter by.
.PARAMETER CodecNotFilter
    Specifies the desired video codec to not filter by.
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
    .\ProcessVideos.ps1 -FolderPath "C:\Videos" -CodecFilter "mp4" -MinBitrate 1000000 -TargetDestination "D:\FilteredVideos"

    This example searches for MP4 video files in the "C:\Videos" folder and its subfolders, with a minimum bitrate of 1 Mbps.
    The filtered videos are then copied to the "D:\FilteredVideos" directory.
#>
param (
    [Parameter(Mandatory = $true)]
    [string] $FolderPath,

    [Parameter(Mandatory = $false)]
    [string] $MediaInfocliPath = "C:\Program Files\MediaInfo_CLI\MediaInfo.exe",

    [Parameter(Mandatory = $false)]
    [switch] $Recursive,

    [Parameter(Mandatory = $false)]
    [string] $CodecFilter,
    
    [Parameter(Mandatory = $false)]
    [string] $CodecNotFilter,

    [Parameter(Mandatory = $false)]
    [int] $MinBitrate,

    [Parameter(Mandatory = $false)]
    [int] $MaxBitrate,
 
    [Parameter(Mandatory = $false)]
    [int] $MinFileSize,

    [Parameter(Mandatory = $false)]
    [int] $MaxFileSize,
    
    [Parameter(Mandatory = $false)]
    [int] $MinWidth,
    
    [Parameter(Mandatory = $false)]
    [int] $MaxWidth,
    
    [Parameter(Mandatory = $false)]
    [int] $ExactWidth,

    [Parameter(Mandatory = $false)]
    [int] $MinHeight,

    [Parameter(Mandatory = $false)]
    [int] $MaxHeight,

    [Parameter(Mandatory = $false)]
    [int] $ExactHeight,

    [Parameter(Mandatory = $false)]
    [string] $AudioLanguageFilter,
 
    [Parameter(Mandatory = $false)]
    [string] $AudioLanguageNotFilter,

    [Parameter(Mandatory = $false)]
    [string] $EncoderFilter,
 
    [Parameter(Mandatory = $false)]
    [string] $EncoderNotFilter,

    [Parameter(Mandatory = $false)]
    [string] $FileNameFilter,

    [Parameter(Mandatory = $false)]
    [string] $TargetDestination,
    
    [Parameter(Mandatory = $false)]
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
    Retrieves detailed information about video files in a specified folder, optionally recursively.
.DESCRIPTION
    This function scans a folder for video files (mp4, mkv, avi, mov, wmv) and retrieves detailed
    information using the specified MediaInfo CLI.
.PARAMETER folderPath
    Specifies the path to the folder containing video files.
.PARAMETER MediaInfoCliPath
    Specifies the path to the MediaInfo CLI executable.
.PARAMETER Recursive
    Switch parameter. If present, the function scans the folder and its subfolders recursively.
.INPUTS
    Accepts file objects as input, specifically video files with extensions: mp4, mkv, avi, mov, wmv.
.OUTPUTS
    Returns an array of custom objects containing detailed information about each video file.
.EXAMPLE
    Get-VideoInfoRecursively -folderPath "C:\Videos" -MediaInfoCliPath "C:\MediaInfo\MediaInfo.exe"
.EXAMPLE
    Get-VideoInfoRecursively -folderPath "D:\Movies" -MediaInfoCliPath "D:\Tools\MediaInfo.exe" -Recursive
#>
function Get-VideoInfoRecursively {
    param (
        [Parameter(Mandatory = $true)]
        [string]$folderPath,

        [Parameter(Mandatory = $true)]
        [string]$MediaInfoCliPath,
        
        [Parameter(Mandatory = $false)]
        [switch] $Recursive
    )
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
        $singleVideoInfo = $null
        $MediaInfoOutput = $null

        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $($FilesScanned + 1) of $totalFilesToScan" -Status "Reading media info: $file" -PercentComplete $progressPercent
        # $videoInfo = Get-VideoInfo -folderPath $file.FullName -MediaInfocliPath $MediaInfoCliPath
        
        try {
            $MediaInfoOutput = & $MediaInfoCliPath --output=JSON --Full $file.FullName | ConvertFrom-Json
        } catch {
            Write-Host 'Exception:' $_.Exception.Message -ForegroundColor Red
            Write-Host 'Problem running MediaInfo' -ForegroundColor Red
            exit 1
        }
    
        $generalTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'General' }
        $videoTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'Video' }
        
        # Array to hold audio languages
        $audioLanguages = @()
        $audioChannels = @()
        
        # Loop through each audio stream
        foreach ($stream in $MediaInfoOutput.media.track) {
            if ($stream.StreamKind -eq "Audio") {
                # Keep track of all languages
                if ($null -ne $stream.Language) {
                    $audioLanguages += $stream.Language.ToUpper()
                        
                } else {
                    $audioLanguages += "UND"
                }
    
                # Keep track of all Audio Channel info (formatted as 2/0/0)
                if ($null -ne $stream.ChannelPositions_String2) {
                    $audioChannels += $stream.ChannelPositions_String2
                } else {
                    $audioChannels += "UND"
                }
            }
        }
        
        # Join the languages into a string or keep it empty if no languages found
        if ($audioLanguages.Count -gt 0) {
            $audioLanguages = $audioLanguages -join ' | '
        } else {
            $audioLanguages = ""
        }
    
        # Join the Channels into a string or keep it empty if no Channels found
        if ($audioChannels.Count -gt 0) {
            $audioChannels = $audioChannels -join ' | '
        } else {
            $audioChannels = ""
        }
            
        $codec = $videoTrack.Format_String
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
        
        $FileInfo = Get-Item -LiteralPath $file.FullName
        
        $singleVideoInfo = [PSCustomObject]@{
            ParentFolder    = $FileInfo.Directory.FullName
            FileName        = $FileInfo.BaseName
            FullPath        = $file.FullName
            Codec           = $codec
            VideoWidth      = $videoWidth
            VideoHeight     = $videoHeight
            VideoBitrate    = $videoBitRate
            TotalBitrate    = $totalBitRate
            FileSize        = $(Format-Size -SizeInBytes $FileInfo.Length)
            AudioLanguages  = $audioLanguages
            AudioChannels   = $audioChannels
            FileSizeByte    = $FileInfo.Length
            VideoDuration   = $VideoDuration   
            Encoder         = $encodedApplication
            RawVideoBitrate = $rawVideoBitRate   
            RawTotalBitrate = $rawTotalBitRate  
        } 
                   
        if ($singleVideoInfo) {
            $allVideoInfo += $singleVideoInfo
        }
        $FilesScanned++
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $FilesScanned of $totalFilesToScan" -Status "Reading media info: $file" -PercentComplete $progressPercent
    }
    Write-Progress -Completed -Activity "Processing: Done"
    return $allVideoInfo
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

<#
.SYNOPSIS
    Formats a size in bytes into a human-readable format.
.DESCRIPTION
    This function takes a size in bytes as input and converts it into a human-readable format, 
    displaying the size in terabytes (TB), gigabytes (GB), megabytes (MB), kilobytes (KB), 
    or bytes based on the magnitude of the input.
.PARAMETER SizeInBytes
    Specifies the size in bytes that needs to be formatted.
.INPUTS
    Accepts a double-precision floating-point number representing the size in bytes.
.OUTPUTS 
    Returns a formatted string representing the size in TB, GB, MB, KB, or bytes.
.EXAMPLE
    Format-Size -SizeInBytes 150000000000
    # Output: "139.81 GB"
    # Description: Formats 150,000,000,000 bytes into gigabytes.
.EXAMPLE
    5000000 | Format-Size
    # Output: "4.77 MB"
    # Description: Pipes 5,000,000 bytes to the function and formats the size into megabytes.
#>

function Format-Size {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$SizeInBytes
    )

    switch ($SizeInBytes) {
        { $_ -ge 1PB } {
            # Convert to PB
            '{0:N2} PB' -f ($SizeInBytes / 1PB)
            break
        }
        { $_ -ge 1TB } {
            # Convert to TB
            '{0:N2} TB' -f ($SizeInBytes / 1TB)
            break
        }
        { $_ -ge 1GB } {
            # Convert to GB
            '{0:N2} GB' -f ($SizeInBytes / 1GB)
            break
        }
        { $_ -ge 1MB } {
            # Convert to MB
            '{0:N2} MB' -f ($SizeInBytes / 1MB)
            break
        }
        { $_ -ge 1KB } {
            # Convert to KB
            '{0:N2} KB' -f ($SizeInBytes / 1KB)
            break
        }
        default {
            # Display in bytes if less than 1KB
            "$SizeInBytes Bytes"
        }
    }
}

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
    AddFilterCriteria -name "CodecFilter" -value "mp4"
    This example adds a filter criteria "CodecFilter=mp4" to the global criteria list.
.NOTES
    This function aids in dynamically constructing filtering criteria for video files based on user-defined parameters.
#>
function AddFilterCriteria($name, $value) {
    if ($value) {
        $global:criteria += "$name=$value"
    }
}

#* Start of script
Clear-Host

# Start searching for video files and extracting information
# $videoInfoList = Get-VideosRecursively $FolderPath $MediaInfocliPath
$videoInfoList = Get-VideoInfoRecursively -folderPath $FolderPath -MediaInfoCliPath $MediaInfocliPath -Recursive

# Create filter description based on applied criteria
$criteria = @()

AddFilterCriteria "CodecFilter" $CodecFilter
AddFilterCriteria "CodecNotFilter" $CodecNotFilter
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

if ($null -eq $criteria -or $criteria.Length -eq 0) {
    $filterDescription = " - No filters applied"
} else {
    $filterDescription = " - Filters: " + ($criteria -join ", ")
}

# Filter based on provided criteria
$videoInfoList = $videoInfoList | Where-Object {
    (!$CodecFilter -or $_.Codec -eq $CodecFilter) -and
    (!$CodecNotFilter -or $_.Codec -ne $CodecNotFilter) -and
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
$sortedVideoInfo += $videoInfoList | Sort-Object -Property Codec, @{Expression = "VideoWidth"; Descending = $true }, @{Expression = "RawTotalBitrate"; Descending = $true }
$sortedVideoInfo | Select-Object -Property ParentFolder, FileName, Codec, VideoWidth, VideoHeight, VideoBitrate, TotalBitrate, RawTotalBitrate, FileSize, FileSizeByte, AudioLanguages, AudioChannels, Encoder | Out-GridView -Title "Video information$filterDescription"

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
