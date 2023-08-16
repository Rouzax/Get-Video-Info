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
    [int] $MinHeight,

    [Parameter()]
    [int] $MaxHeight,

    [Parameter()]
    [int] $ExactWidth,

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
    if ($bitRate -eq $null) {
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
                FileName         = (Get-Item $filePath).Name
                FullPath         = $filePath
                Codec            = $codec
                "Video Width"    = [int]$videoWidth
                "Video Height"   = [int]$videoHeight
                "Video Bit Rate" = $bitRateFormatted
                "Total Bit Rate" = $totalBitRateFormatted
                RawBitRate       = [int]$totalBitRate
                Encoder          = $encoder
            }
        }
    }

    return $videoInfo
}

# Recursive function to search for video files and extract information
function Get-VideosRecursively($folderPath, $ffprobePath) {
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

if ($CodecFilter) {
    $videoInfoList = $videoInfoList | Where-Object { $_.Codec -eq $CodecFilter }
}

if ($MinBitrate) {
    $videoInfoList = $videoInfoList | Where-Object { $_.RawBitRate -ge $MinBitrate }
}

if ($MaxBitrate) {
    $videoInfoList = $videoInfoList | Where-Object { $_.RawBitRate -le $MaxBitrate }
}

if ($MinWidth) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Width" -ge $MinWidth }
}

if ($MaxWidth) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Width" -le $MaxWidth }
}

if ($MinHeight) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Height" -ge $MinHeight }
}

if ($MaxHeight) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Height" -le $MaxHeight }
}

if ($ExactWidth) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Width" -eq $ExactWidth }
}

if ($ExactHeight) {
    $videoInfoList = $videoInfoList | Where-Object { $_."Video Height" -eq $ExactHeight }
}

if ($EncoderFilter) {
    $videoInfoList = $videoInfoList | Where-Object { $_.Encoder -like "*$EncoderFilter*" }
}

$sortedVideoInfo = $videoInfoList | Sort-Object -Property Codec, "Video Width", RawBitRate -Descending
$sortedVideoInfo | Format-Table -AutoSize FileName, Codec, "Video Width", "Video Height", "Video Bit Rate", "Total Bit Rate", RawBitRate, Encoder

# Copy files to the target destination if specified
if ($TargetDestination) {
    foreach ($videoInfo in $sortedVideoInfo) {
        $sourceFilePath = $videoInfo.FullPath
        $relativePath = $sourceFilePath.Substring($FolderPath.Length)
        $destinationFilePath = Join-Path $TargetDestination $relativePath

        # Create the destination directory if it doesn't exist
        $destinationDirectory = [System.IO.Path]::GetDirectoryName($destinationFilePath)
        if (-not (Test-Path $destinationDirectory)) {
            $null = New-Item -ItemType Directory -Path $destinationDirectory
        }

        # Copy the file to the destination
        $null = Copy-Item -Path $sourceFilePath -Destination $destinationFilePath -Force
    }
}
