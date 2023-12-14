<#
.SYNOPSIS
This script extracts information about video files using MediaInfo CLI, filters the results based on specified criteria,
and optionally copies the files to a target destination.

.DESCRIPTION
This script searches for video files in a specified folder and its subfolders, extracts detailed information about them
using MediaInfo CLI, and applies various filters based on user-defined criteria such as format, bitrate, resolution, etc.
It also provides the option to copy the filtered files to a target destination.

.PARAMETER FolderPath
Specifies the path to the root folder containing video files.

.PARAMETER MediaInfocliPath
Specifies the path to the MediaInfo CLI executable. Default: "C:\Program Files\MediaInfo_CLI\MediaInfo.exe".

.PARAMETER Recursive
If specified, the script will search for video files recursively in subfolders.

.PARAMETER FormatFilter
Specifies the desired video format to filter by.

.PARAMETER MinBitrate
Specifies the minimum video bitrate (bps) to filter by.

.PARAMETER MaxBitrate
Specifies the maximum video bitrate (bps) to filter by.

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
Specifies a keyword to filter file name.

.PARAMETER TargetDestination
Specifies the target destination to copy the filtered video files.

.PARAMETER CopyRelatedFiles
If provided will tell the script to copy jpg and png images with the same File BaseName as the video to the target

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
    [int] $MinBitrate,

    [Parameter()]
    [int] $MaxBitrate,
    
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

# Function to convert bitrate to human-readable format with two decimal places
function Convert-BitRate($bitRate) {
    <#
    .SYNOPSIS
    This function converts a given Bitrate value into a human-readable format, including bps, kbps, and Mbps.
    
    .DESCRIPTION
    The Convert-BitRate function takes a Bitrate value as input and converts it into a more readable format. It calculates and rounds the Bitrate to kilobits per second (kbps) and megabits per second (Mbps) as appropriate, and then returns the formatted result with the corresponding unit.

    .PARAMETER bitRate
    Specifies the Bitrate value that needs to be converted. It should be provided in bits per second (bps).

    .EXAMPLE
    Example 1:
    Convert-BitRate -bitRate 2500000
    This example converts a Bitrate of 2500000 bps into 2.50 Mbps.
    #>
    if ($null -eq $bitRate) {
        return ""
    }

    $kbps = [math]::Round($bitRate / 1000, 2)
    $mbps = [math]::Round($kbps / 1000, 2)

    if ($mbps -ge 1) {
        return ("{0:N2}" -f $mbps) + " Mbps"
    } elseif ($kbps -ge 1) {
        return ("{0:N2}" -f $kbps) + " kbps"
    } else {
        return "${bitRate} bps"
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
        $videoBitRate = Convert-BitRate $rawVideoBitRate
    } else {
        $videoBitRate = $null
    }
    
    if ($generalTrack.OverallBitRate) {
        $rawTotalBitRate = [int]$generalTrack.OverallBitRate
        $totalBitRate = Convert-BitRate $rawTotalBitRate
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

    $parentFolder = (Get-Item -LiteralPath $filePath).Directory.FullName
    
    $singleVideoInfo = [PSCustomObject]@{
        ParentFolder    = $parentFolder
        FileName        = (Get-Item -LiteralPath $filePath).BaseName
        FullPath        = $filePath
        Format          = $format
        Codec           = $codec
        VideoWidth      = $videoWidth
        VideoHeight     = $videoHeight
        VideoBitrate    = $videoBitRate
        TotalBitrate    = $totalBitRate
        VideoDuration   = $VideoDuration   
        Encoder         = $encodedApplication
        RawVideoBitrate = $rawVideoBitRate   
        RawTotalBitrate = $rawTotalBitRate  
    }
    return $singleVideoInfo
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

# Start searching for video files and extracting information
$videoInfoList = Get-VideosRecursively $FolderPath $MediaInfocliPath

# Create filter description based on applied criteria
$filterDescription = "Filter: "
$filterDescription += if ($FormatFilter) {
    "FormatFilter=$FormatFilter, " 
}
$filterDescription += if ($MinBitrate) {
    "MinBitrate=$MinBitrate, " 
}
$filterDescription += if ($MaxBitrate) {
    "MaxBitrate=$MaxBitrate, " 
}
$filterDescription += if ($MinWidth) {
    "MinWidth=$MinWidth, " 
}
$filterDescription += if ($MaxWidth) {
    "MaxWidth=$MaxWidth, " 
}
$filterDescription += if ($ExactWidth) {
    "ExactWidth=$ExactWidth, " 
}
$filterDescription += if ($MinHeight) {
    "MinHeight=$MinHeight, " 
}
$filterDescription += if ($MaxHeight) {
    "MaxHeight=$MaxHeight, " 
}
$filterDescription += if ($ExactHeight) {
    "ExactHeight=$ExactHeight, " 
}
$filterDescription += if ($EncoderFilter) {
    "EncoderFilter=$EncoderFilter, " 
}
$filterDescription += if ($FileNameFilter) {
    "FileNameFilter=$FileNameFilter, " 
}
$filterDescription = $filterDescription.TrimEnd(", ")

# Filter based on provided criteria
$videoInfoList = $videoInfoList | Where-Object {
    (!$FormatFilter -or $_.Format -eq $FormatFilter) -and
    (!$MinBitrate -or $_.RawTotalBitrate -ge $MinBitrate) -and
    (!$MaxBitrate -or $_.RawTotalBitrate -le $MaxBitrate) -and
    (!$MinWidth -or $_.VideoWidth -ge $MinWidth) -and
    (!$MaxWidth -or $_.VideoWidth -le $MaxWidth) -and
    (!$MinHeight -or $_.VideoHeight -ge $MinHeight) -and
    (!$MaxHeight -or $_.VideoHeight -le $MaxHeight) -and
    (!$ExactWidth -or $_.VideoWidth -eq $ExactWidth) -and
    (!$ExactHeight -or $_.VideoHeight -eq $ExactHeight) -and
    (!$EncoderFilter -or $_.Encoder -like "*$EncoderFilter*") -and
    (!$EncoderNotFilter -or $_.Encoder -notlike "*$EncoderNotFilter*") -and
    (!$FileNameFilter -or $_.FileName -like "*$FileNameFilter*")
}

$sortedVideoInfo = @()
$sortedVideoInfo += $videoInfoList | Sort-Object -Property Format, @{Expression = "VideoWidth"; Descending = $true }, @{Expression = "RawTotalBitrate"; Descending = $true }
$sortedVideoInfo | Select-Object -Property ParentFolder, FileName, Format, VideoWidth, VideoHeight, VideoBitrate, TotalBitrate, RawTotalBitrate, Encoder | Out-GridView -Title "Video information $filterDescription"

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
            Write-Progress -Activity "Copying Files" -Status "Copying $sourceFilePath" -PercentComplete $progressPercent

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
            Write-Progress -Activity "Copying Files" -Status "Copied $sourceFilePath" -PercentComplete $progressPercent
        }
    } elseif ($response -eq 'N' -or $response -eq 'n') {
        Write-Host "Run script again with different criteria"
    }
}