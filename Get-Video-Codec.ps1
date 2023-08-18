param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $FolderPath,

    [Parameter(Position = 1)]
    [string] $FFprobePath = "C:\Program Files\FFmpeg\ffprobe.exe",

    [Parameter()]
    [switch] $Recursive,

    [Parameter()]
    [string] $CodecFilter,

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
    [string] $TargetDestination
)

# Check if ffprobe executable exists
if (-not (Test-Path $FFprobePath)) {
    Write-Host "Error: FFprobe executable not found at the specified path: $FFprobePath"
    Write-Host "Please provide the correct path to the FFprobe executable using the -FFprobePath parameter."
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

# Function to extract video information using FFprobe
function Get-VideoInfo($filePath, $ffprobePath) {
    <#
    .SYNOPSIS
    Retrieves detailed information about a video file using FFprobe.

    .DESCRIPTION
    This function takes a video file path and the path to the FFprobe executable as inputs.
    It uses FFprobe to extract information about the video, such as codec, dimensions, bitrate, and encoder.

    .PARAMETER filePath
    Specifies the path to the video file for which information needs to be extracted.

    .PARAMETER ffprobePath
    Specifies the path to the FFprobe executable.

    .EXAMPLE
    Get-VideoInfo -filePath "C:\Videos\video.mp4" -ffprobePath "C:\Program Files\FFmpeg\ffprobe.exe"
    This example retrieves information about the video file "video.mp4" using FFprobe.

    .NOTES
    This function requires FFprobe to be installed on the system and the ffprobePath parameter to point to its location.
    #>
    $ffprobeOutput = & $ffprobePath -v error -print_format json -show_format -show_streams "$filePath" | ConvertFrom-Json

    $videoInfo = $null

    foreach ($stream in $ffprobeOutput.streams) {
        if ($stream.codec_type -eq "video") {
            $codec = $stream.codec_name
            $videoWidth = $stream.width
            $videoHeight = $stream.height
            $bitRate = $stream.bit_rate
            $bitRateFormatted = Convert-BitRate $bitRate

            $format = $ffprobeOutput.format
            $totalBitRate = $format.bit_rate
            $totalBitRateFormatted = Convert-BitRate $totalBitRate
            $tags = $format.tags
            $encoder = $tags.encoder

            $videoInfo = [PSCustomObject]@{
                FileName        = (Get-Item $filePath).Name
                FullPath        = $filePath
                Codec           = $codec
                "Video Width"   = [int]$videoWidth
                "Video Height"  = [int]$videoHeight
                "Video Bitrate" = $bitRateFormatted
                "Total Bitrate" = $totalBitRateFormatted
                RawBitRate      = [int]$totalBitRate
                Encoder         = $encoder
            }
        }
    }

    return $videoInfo
}

# Recursive function to search for video files and extract information
function Get-VideosRecursively($folderPath, $ffprobePath) {
    <#
    .SYNOPSIS
    Recursively searches for video files in a folder and its subfolders and extracts information using FFprobe.

    .DESCRIPTION
    This function searches for video files (with extensions mp4, mkv, avi, mov, wmv) in the specified folder
    and its subfolders. For each video file found, it calls the Get-VideoInfo function to extract detailed information.

    .PARAMETER folderPath
    Specifies the path of the folder to start the search from.

    .PARAMETER ffprobePath
    Specifies the path to the FFprobe executable.

    .EXAMPLE
    Get-VideosRecursively -folderPath "C:\Videos" -ffprobePath "C:\Program Files\FFmpeg\ffprobe.exe"
    This example searches for video files in the "C:\Videos" folder and its subfolders and extracts information using FFprobe.

    .NOTES
    This function requires the Get-VideoInfo function and FFprobe to be installed on the system.
    #>

    $videoFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }

    $allVideoInfo = @()
    foreach ($file in $videoFiles) {
        $videoInfo = Get-VideoInfo $file.FullName $ffprobePath
        if ($videoInfo) {
            $allVideoInfo += $videoInfo
        }
    }

    if ($Recursive) {
        $subfolders = Get-ChildItem -Path $folderPath -Directory
        foreach ($subfolder in $subfolders) {
            $subVideoInfo = Get-VideosRecursively $subfolder.FullName $ffprobePath
            $allVideoInfo += $subVideoInfo
        }
    }

    return $allVideoInfo
}

# Start searching for video files and extracting information
$videoInfoList = Get-VideosRecursively $FolderPath $FFprobePath

# Filter based on provided criteria
$videoInfoList = $videoInfoList | Where-Object {
    (!$CodecFilter -or $_.Codec -eq $CodecFilter) -and
    (!$MinBitrate -or $_.RawBitRate -ge $MinBitrate) -and
    (!$MaxBitrate -or $_.RawBitRate -le $MaxBitrate) -and
    (!$MinWidth -or $_."Video Width" -ge $MinWidth) -and
    (!$MaxWidth -or $_."Video Width" -le $MaxWidth) -and
    (!$MinHeight -or $_."Video Height" -ge $MinHeight) -and
    (!$MaxHeight -or $_."Video Height" -le $MaxHeight) -and
    (!$ExactWidth -or $_."Video Width" -eq $ExactWidth) -and
    (!$ExactHeight -or $_."Video Height" -eq $ExactHeight) -and
    (!$EncoderFilter -or $_.Encoder -like "*$EncoderFilter*")
}

$sortedVideoInfo = $videoInfoList | Sort-Object -Property Codec, "Video Width", RawBitRate -Descending
$sortedVideoInfo | Format-Table -AutoSize FileName, Codec, "Video Width", "Video Height", "Video Bitrate", "Total Bitrate", RawBitRate, Encoder

# Copy files to the target destination if specified
if ($TargetDestination) {
    $totalFiles = $sortedVideoInfo.Count
    $copiedFiles = 0

    foreach ($videoInfo in $sortedVideoInfo) {
        $sourceFilePath = $videoInfo.FullPath
        $relativePath = $sourceFilePath.Substring($FolderPath.Length)
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
        $null = Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force

        $copiedFiles++
    }
}