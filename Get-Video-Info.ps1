<#
.SYNOPSIS
    Retrieves detailed information about video files in a specified folder, optionally recursively.
.DESCRIPTION
    This function scans a folder for video files (mp4, mkv, avi, mov, wmv) and retrieves detailed
    information using the specified MediaInfo CLI and optionally allows for the filtered video files to be copied somewhere
.PARAMETER folderPath
    Specifies the path to the folder containing video files.
.PARAMETER MediaInfoCliPath
    Specifies the path to the MediaInfo CLI executable.
.PARAMETER Recursive
    Switch parameter. If present, the function scans the folder and its subfolders recursively.
.PARAMETER VideoCodecFilter
    Specifies the video codec to filter video files.
.PARAMETER VideoCodecNotFilter
    Specifies the video codec to exclude from the results.
.PARAMETER MinBitrate
    Specifies the minimum total bitrate for filtering video files.
.PARAMETER MaxBitrate
    Specifies the maximum total bitrate for filtering video files.
.PARAMETER MinFileSize
    Specifies the minimum file size (in bytes) for filtering video files.
.PARAMETER MaxFileSize
    Specifies the maximum file size (in bytes) for filtering video files.
.PARAMETER MinWidth
    Specifies the minimum video width for filtering video files.
.PARAMETER MaxWidth
    Specifies the maximum video width for filtering video files.
.PARAMETER ExactWidth
    Specifies the exact video width for filtering video files.
.PARAMETER MinHeight
    Specifies the minimum video height for filtering video files.
.PARAMETER MaxHeight
    Specifies the maximum video height for filtering video files.
.PARAMETER ExactHeight
    Specifies the exact video height for filtering video files.
.PARAMETER AudioCodecFilter
    Specifies the audio codec to filter video files.
.PARAMETER AudioCodecNotFilter
    Specifies the audio codec to exclude from the results.
.PARAMETER AudioLanguageFilter
    Specifies the audio language to filter video files.
.PARAMETER AudioLanguageNotFilter
    Specifies the audio language to exclude from the results.
.PARAMETER EncoderFilter
    Specifies the encoder to filter video files.
.PARAMETER EncoderNotFilter
    Specifies the encoder to exclude from the results.
.PARAMETER FileNameFilter
    Specifies a filter for video file names.
.PARAMETER TargetDestination
    Specifies the destination folder for copying filtered video files.
.PARAMETER CopyRelatedFiles
    If provided, instructs the script to copy jpg and png images
    with the same File BaseName as the video to the target.
.EXAMPLE
    .\Get-Video-Info.ps1 -folderPath "C:\Videos" -MediaInfoCliPath "C:\MediaInfo\MediaInfo.exe"
.EXAMPLE
    .\Get-Video-Info.ps1 -FolderPath "C:\Videos" -CodecFilter "AVC" -MinBitrate 1000000 -TargetDestination "D:\FilteredVideos"

    This example searches for AVC video files in the "C:\Videos" folder and its subfolders, 
    with a minimum bitrate of 1 Mbps. 
    The filtered videos are then copied to the "D:\FilteredVideos" directory.
#>
param (
    [Parameter(Mandatory = $true)]
    [string] $FolderPath,

    [Parameter(Mandatory = $false)]
    [string] $MediaInfocliPath,

    [Parameter(Mandatory = $false)]
    [switch] $Recursive,

    [Parameter(Mandatory = $false)]
    [string] $VideoCodecFilter,
    
    [Parameter(Mandatory = $false)]
    [string] $VideoCodecNotFilter,

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
    [string] $AudioCodecFilter,
    
    [Parameter(Mandatory = $false)]
    [string] $AudioCodecNotFilter,

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

<#
.SYNOPSIS
    Retrieves detailed information about video files in a specified folder, optionally recursively.
.DESCRIPTION
    This function scans a folder for video files (mp4, mkv, avi, mov, wmv) and retrieves detailed
    information using the specified MediaInfo CLI.
.PARAMETER videoFiles
    Object that holds all the video files that need to be scanned.
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
        [object]$videoFiles,

        [Parameter(Mandatory = $true)]
        [string]$MediaInfoCliPath
    )

    $totalFilesToScan = ($videoFiles).Count
    $FilesScanned = 0

    $allVideoInfo = @()

    foreach ($file in $videoFiles) {
        $singleVideoInfo = $null
        $MediaInfoOutput = $null

        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $($FilesScanned + 1) of $totalFilesToScan" -Status "Reading media info: $($file.Name)" -PercentComplete $progressPercent
        
        try {
            $MediaInfoOutput = & $MediaInfoCliPath --output=JSON $file.FullName | ConvertFrom-Json
        } catch {
            Write-Host 'Exception:' $_.Exception.Message -ForegroundColor Red
            Write-Host 'Problem running MediaInfo' -ForegroundColor Red
            exit 1
        }

        # Array initialization to hold audio languages
        $audioCodecs = @()
        $audioLanguages = @()
        $audioChannels = @()

        foreach ($stream in $MediaInfoOutput.media.track) {
            # Get information from General stream
            if ($stream.'@type' -eq 'General') {
                # Get the total Bitrate
                if ($stream.OverallBitRate) {
                    [int]$rawTotalBitRate = $stream.OverallBitRate
                    $totalBitRate = Convert-BitRate -bitratePerSecond $rawTotalBitRate
                } else {
                    $totalBitRate = $null
                }

                # Get encoding Application
                [string]$encodedApplication = $stream.Encoded_Application

                # Extracting the duration
                [decimal]$rawDuration = $stream.Duration
                # Rounding video duration
                $videoDuration = [math]::Floor($rawDuration)

                # Get File Container
                [string]$fileContainer = $stream.Format
            
                # Get File Extension
                [string]$fileExt = $stream.FileExtension
            }

            # Get information from Video stream
            elseif ($stream.'@type' -eq 'Video') {
                # Get Codec information from Video
                [string]$videoCodec = $stream.Format
                
                # Get Video dimensions
                [int]$videoWidth = $stream.Width 
                [int]$videoHeight = $stream.Height 
                
                # Get Video Bitrate
                if ($stream.BitRate) {
                    [int]$rawVideoBitRate = $stream.BitRate
                    $videoBitRate = Convert-BitRate -bitratePerSecond $rawVideoBitRate
                } else {
                    $videoBitRate = $null
                }

                # Get Video HDR
                if ($stream.HDR_Format_Compatibility) {
                    [string]$videoHDR = $stream.HDR_Format_Compatibility
                } else {
                    $videoHDR = "SDR"
                }
            } 

            # Get information from Audio stream
            elseif ($stream.'@type' -eq 'Audio') {
                # Keep track of all Audio Codec info
                if ($null -ne $stream.Format) {
                    $audioCodecs += $stream.Format
                } else {
                    $audioCodecs += "UND"
                }

                # Keep track of all languages
                if ($null -ne $stream.Language) {
                    $audioLanguages += $stream.Language.ToUpper()
                    
                } else {
                    $audioLanguages += "UND"
                }

                # Keep track of all Audio Channel info
                if ($null -ne $stream.Channels) {
                    $audioChannels += $stream.Channels
                } else {
                    $audioChannels += "UND"
                }
            } 

        }
                
        # Join the Codecs into a string or keep it empty if no Codecs found
        if ($audioCodecs.Count -gt 0) {
            $audioCodecs = $audioCodecs -join ' | '
        } else {
            $audioCodecs = ""
        }

        # Join the Languages into a string or keep it empty if no Languages found
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
        
        $singleVideoInfo = [PSCustomObject]@{
            ParentFolder    = $file.Directory.FullName
            FileName        = $file.BaseName
            FullPath        = $file.FullName
            FileExt         = $fileExt
            FileContainer   = $fileContainer
            VideoCodec      = $videoCodec
            VideoHDR        = $videoHDR
            VideoWidth      = $videoWidth
            VideoHeight     = $videoHeight
            VideoBitrate    = $videoBitRate
            TotalBitrate    = $totalBitRate
            FileSize        = $(Format-Size -SizeInBytes $file.Length)
            AudioCodecs     = $audioCodecs
            AudioLanguages  = $audioLanguages
            AudioChannels   = $audioChannels
            FileSizeByte    = $file.Length
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
        Write-Progress -Activity "Processing: $FilesScanned of $totalFilesToScan" -Status "Reading media info: $($file.Name)" -PercentComplete $progressPercent
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
    AddFilterCriteria -name "VideoCodecFilter" -value "mp4"
    This example adds a filter criteria "VideoCodecFilter=mp4" to the global criteria list.
.NOTES
    This function aids in dynamically constructing filtering criteria for video files based on user-defined parameters.
#>
function AddFilterCriteria($name, $value) {
    if ($value) {
        $global:criteria += "$name=$value"
    }
}

<#
.SYNOPSIS
    Updates MediaInfoCLI to the latest version.
.DESCRIPTION
    This function checks the local version of MediaInfoCLI and updates it to the latest
    version available on GitHub if a newer version exists. It also verifies if the script
    has the necessary permissions to write to the specified installation path.
.PARAMETER MediaInfoCLIPath
    Specifies the path where MediaInfoCLI should be installed or updated. Ensure the
    script has write permissions to this location. If elevation is required, the script
    prompts the user to run with administrator privileges.
.INPUTS
    None
.OUTPUTS 
    None
.EXAMPLE
    Update-MediaInfoCLI -MediaInfoCLIPath "C:\Program Files\HandBrake\MediaInfoCLI.exe"
    # Checks and updates MediaInfoCLI to the latest version in the specified path.
.EXAMPLE
    Update-MediaInfoCLI -MediaInfoCLIPath "D:\HandBrake\MediaInfoCLI.exe"
    # Checks and updates MediaInfoCLI to the latest version in the specified path.
#>
function Update-MediaInfoCLI {
    param(
        [string]$MediaInfoCLIPath
    )
    Write-Host "`nChecking if MediaInfo CLI is available and update is needed" -ForegroundColor Magenta
  
    # Extract the folder path without the executable
    $FolderPath = Split-Path $MediaInfoCLIPath
    
    # Check if $MediaInfoCLIPath already exists
    if (Test-Path $MediaInfoCLIPath) {
        Write-Host "    MediaInfo CLI is already installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$MediaInfoCLIPath'" -ForegroundColor Cyan  
        $fileVersion = (Get-Command $MediaInfoCLIPath).FileVersionInfo.ProductVersion
    } else {
        # MediaInfoCLI is not installed at the specified path
        Write-Host "    MediaInfo CLI is not installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$MediaInfoCLIPath'" -ForegroundColor Cyan  
        Write-Host "    Will try to get the latest version online" -ForegroundColor DarkGray 
    }
    
    # Check if writing to $FolderPath requires elevated permissions
    try {
        # Check if the folder exists, if not create it
        if (-not (Test-Path $FolderPath)) {
            New-Item -ItemType Directory -Path $FolderPath -ErrorAction Stop | Out-Null
        }

        # Check if writing to $FolderPath requires elevated permissions
        $testPath = Join-Path $FolderPath "Test-Permission.txt"
        $null | Out-File -FilePath $testPath -Force -ErrorAction Stop
        Remove-Item -Path $testPath -Force -ErrorAction Stop
    } catch {
        if ($fileVersion) {
            Write-Host "    Insufficient permissions to write to: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$FolderPath'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Insufficient permissions to write to: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$FolderPath'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please run the script as an administrator." -ForegroundColor DarkRed
            Exit
        }
    }

    # Define version file path
    $versionFilePath = Join-Path $FolderPath "MediaInfoCLI-Version.json"

    if (Test-Path $MediaInfoCLIPath) {
        # MediaInfoCLI is installed
        if (Test-Path $versionFilePath) {
            # Read the version information from the JSON file
            $versionInfo = Get-Content $versionFilePath | ConvertFrom-Json
            $currentVersion = $versionInfo.Version
            $lastLocalUpdate = $versionInfo.LastLocalUpdate
        } else {
            # No version file found, write current version and date to MediaInfoCLI-Version.json
            $lastLocalUpdate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $currentVersion = $fileVersion
            # Save the version information to a JSON file
            $versionObject = [PSCustomObject]@{
                Version         = $currentVersion
                LastLocalUpdate = $lastLocalUpdate
            }
            $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force
        }
        
        # Display installed version and last update date
        Write-Host "    Installed version of MediaInfo CLI: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan
        Write-Host "    Last local update: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$lastLocalUpdate" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 
    } else {
        # MediaInfoCLI is not installed
        $currentVersion = "0.0"
        $lastLocalUpdate = $null
    }

    # Define the download URL for MediaInfo_CLI
    $MediaInfoCLIdownloadUrl = "https://mediaarea.net/en/MediaInfo/Download/Windows"

    try {
        # Use Invoke-WebRequest to fetch the HTML content of the download page
        $response = Invoke-WebRequest -Uri $MediaInfoCLIdownloadUrl -UseBasicParsing -ErrorAction Stop
    } catch {
        if ($fileVersion) {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Parse the HTML response to extract the download links
    $downloadLinks = $response.Links | Where-Object { $_.href -match "/MediaInfo_CLI_[\d.]+_Windows_x64.zip" }

    # Check if download links are found
    if ($downloadLinks.Count -eq 0) {
        if ($fileVersion) {
            Write-Host "    No download links found at: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Did not find: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    No download links found at: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Did not find: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Initialize variables to store version and latest link
    $latestVersion = $null
    $latestLink = $null

    # Iterate through each download link to find the latest version
    foreach ($link in $downloadLinks) {
        if ($link.href -match '/MediaInfo_CLI_([\d.]+)_Windows_x64.zip') {
            $version = $Matches[1]
            if (-not $latestVersion -or [version]::Parse($version) -gt [version]::Parse($latestVersion)) {
                $latestVersion = $version
                $latestLink = $link
            }
        }
    }
    
    # Check if latest link is found
    if (-not $latestLink) {
        if ($fileVersion) {
            Write-Host "    Did not find: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Did not find: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Construct the full download URL
    $downloadUrl = "https:" + $latestLink.href
    
    if ($null -eq $downloadUrl) {
        # Download URL is empty
        if ($fileVersion) {
            Write-Host "    No download links found at: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    No download links found at: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Compare versions
    if ($latestVersion -gt $currentVersion) {
        # Newer version available, proceed with update

        # Display updating message
        Write-Host "    Updating MediaInfoCLI from version " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion " -ForegroundColor Cyan -NoNewline 
        Write-Host "to " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$latestVersion" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 

        # Define download path
        $downloadPath = Join-Path $FolderPath "MediaInfo_CLI_$($latestVersion)_Windows_x64.zip"

        # Download the zip file
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -ErrorAction Stop
        } catch {
            if ($fileVersion) {
                Write-Host "    Failed to download MediaInfoCLI." -ForegroundColor DarkYellow
    
                Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
                Write-Host $fileVersion -ForegroundColor Cyan
                return
            } else {
                Write-Host "    Failed to download MediaInfoCLI." -ForegroundColor DarkRed
    
                Write-Host "    No current version found" -ForegroundColor DarkRed
                Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
                Exit
            }
        }

        # Extract the contents
        Write-Host "    Extracting files to " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$FolderPath" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 
        try {
            Expand-Archive -Path $downloadPath -DestinationPath $FolderPath -Force -ErrorAction Stop
        } catch {
            Write-Host "    Failed to extract files. Aborting update." -ForegroundColor DarkRed
            Write-Host "    Error: $_" -ForegroundColor DarkRed
            exit
        }

        # Clean up the downloaded zip file
        Remove-Item -Path $downloadPath -Force

        # Update the version information in the JSON file
        $versionObject = [PSCustomObject]@{
            Version         = $latestVersion
            LastLocalUpdate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force

        Write-Host "    Update completed successfully." -ForegroundColor DarkGray
    } else {
        Write-Host "    MediaInfoCLI is already up to date (version :" -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan -NoNewline
        Write-Host ")." -ForegroundColor DarkGray 
    }
}

#* Start of script
Clear-Host

# Handle no MediaInfocliPath Path given as parameter
if (-not $PSBoundParameters.ContainsKey('MediaInfocliPath')) {
    $MediaInfocliPath = Join-Path -Path $PSScriptRoot -ChildPath "\MediaInfoCLI\MediaInfo.exe"
} else {
    # Check if path is Folder 
    $MediaInfocliFolder = (Get-Item -LiteralPath $MediaInfocliPath) -is [System.IO.DirectoryInfo]
    if ($MediaInfocliFolder) {
        $MediaInfocliPath = Join-Path -Path $MediaInfocliPath -ChildPath "\MediaInfo.exe"
    }
}

# Check version of HandbrakeCLi Path that was given and update if needed
Update-MediaInfoCLI -MediaInfoCLIPath $MediaInfocliPath


$videoFiles = @()

$fileExtensions = "mp4", "m4v", "mkv", "avi", "mov", "wmv", "ts", "flv", "webm", "mpeg", "mpg"
$videoFilesParams = @{
    Recurse = $Recursive
    Path    = $folderPath
    File    = $true
}

$videoFiles = Get-ChildItem @videoFilesParams | Where-Object { $_.Extension -match '\.({0})$' -f ($fileExtensions -join '|') } 
if ($null -eq $videoFiles) {
    Write-Host "No video files found in location: " -ForegroundColor DarkYellow -NoNewline
    Write-Host "'$folderPath'" -ForegroundColor Cyan
    Write-Host "Maybe try with '-Recursive parameter"-ForegroundColor DarkYellow
    exit
}
 
# filter out min and max files sizes
if ($MinFileSize -or $MaxFileSize) {
    $videoFiles = $videoFiles | Where-Object {
        (
                ($_.Length -ge $MinFileSize) -and
                ($_.Length -le $MaxFileSize)
        )
    }
}

# filter out file name
if ($FileNameFilter) {
    $videoFiles = $videoFiles | Where-Object {
        (
                ($_.BaseName -like "*$FileNameFilter*")
        )
    }
}

$videoInfoList = Get-VideoInfoRecursively -videoFiles $videoFiles -MediaInfoCliPath $MediaInfocliPath

# Create filter description based on applied criteria
$criteria = @()

AddFilterCriteria "VideoCodecFilter" $VideoCodecFilter
AddFilterCriteria "VideoCodecNotFilter" $VideoCodecNotFilter
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
AddFilterCriteria "AudioCodecFilter" $AudioCodecFilter
AddFilterCriteria "AudioCodecNotFilter" $AudioCodecNotFilter
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
    (!$VideoCodecFilter -or $_.VideoCodec -eq $VideoCodecFilter) -and
    (!$VideoCodecNotFilter -or $_.VideoCodec -ne $VideoCodecNotFilter) -and
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
    (!$AudioCodecFilter -or $_.AudioCodecs -like "*$AudioCodecFilter*") -and
    (!$AudioCodecNotFilter -or $_.AudioCodecs -notlike "*$AudioCodecNotFilter*") -and
    (!$AudioLanguageFilter -or $_.AudioLanguages -like "*$AudioLanguageFilter*") -and
    (!$AudioLanguageNotFilter -or $_.AudioLanguages -notlike "*$AudioLanguageNotFilter*") -and
    (!$EncoderFilter -or $_.Encoder -like "*$EncoderFilter*") -and
    (!$EncoderNotFilter -or $_.Encoder -notlike "*$EncoderNotFilter*") -and
    (!$FileNameFilter -or $_.FileName -like "*$FileNameFilter*")
}

$videoInfoList | Select-Object -Property ParentFolder, FileName, FileExt, FileContainer , VideoCodec, VideoHDR, VideoWidth, VideoHeight, VideoBitrate, TotalBitrate, RawTotalBitrate, FileSize, FileSizeByte, AudioLanguages, AudioCodecs, AudioChannels, Encoder | Out-GridView -Title "Video information$filterDescription"

# Copy files to the target destination if specified
if ($TargetDestination) {
    $response = Read-Host "Do you want to start the copy of the videos listed? (Y/N)"

    if ($response -eq 'Y' -or $response -eq 'y') {
        $totalFiles = $videoInfoList.Count
        $copiedFiles = 0

        foreach ($videoInfo in $videoInfoList) {
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
