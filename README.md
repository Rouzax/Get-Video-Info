# Video File Processing with PowerShell

![PowerShell Version](https://img.shields.io/badge/PowerShell-5.1%2B-blue)

## Overview

This PowerShell script enables comprehensive extraction and filtering of video file information using MediaInfo CLI. It offers extensive filtering options based on various criteria such as codec, bitrate, resolution, and more. Additionally, it provides an option to copy filtered videos to a specified destination.
![image](https://github.com/Rouzax/Get-Video-Info/assets/4103090/6cd64e68-ee44-4cc2-9dfc-254b0c075355)


## Features

- **File Information Extraction:** Utilizes MediaInfo CLI to extract detailed information about video files.
- **Flexible Filtering:** Allows filtering based on codec, bitrate, file size, resolution, audio language, encoder, and file name.
- **Recursive Search:** Optionally searches subfolders for video files.
- **Targeted File Copy:** Copies filtered videos to a specified destination folder.
- **Related Files Copy:** Optionally copies images (jpg, png) with the same base name as the video.

## Parameters

- `FolderPath`: Path to the root folder containing video files.
- `MediaInfocliPath`: Path to the MediaInfo CLI executable (Default: "C:\Program Files\MediaInfo_CLI\MediaInfo.exe").
- `Recursive`: Search for video files recursively in subfolders.
- `CodecFilter`, `MinBitrate`, `MaxBitrate`, `MinFileSize`, `MaxFileSize`: Filtering criteria for codec, bitrate, file size.
- `MinWidth`, `MaxWidth`, `ExactWidth`, `MinHeight`, `MaxHeight`, `ExactHeight`: Criteria for video resolution.
- `AudioLanguageFilter`, `AudioLanguageNotFilter`, `EncoderFilter`, `EncoderNotFilter`, `FileNameFilter`: Filtering based on audio language, encoder, and file name.
- `TargetDestination`: Destination folder for copied filtered video files.
- `CopyRelatedFiles`: Option to copy jpg and png images with the same base name as the video.

## Usage

Example:

```powershell
.\Get-Video-Codec.ps1 -FolderPath "C:\Videos" -CodecFilter "AVC" -MinBitrate 1000000 -TargetDestination "D:\FilteredVideos"
```

This example searches for AVC video files in the "C:\Videos" folder and its subfolders, with a minimum bitrate of 1 Mbps. Filtered videos are copied to the "D:\FilteredVideos" directory.

## Installation and Requirements

1. Ensure PowerShell version 5.1 or later is installed.
2. Download and install MediaInfo CLI from [MediaInfo website](https://mediaarea.net/MediaInfo).
3. Clone or download this script to your local machine.

## Notes

- Ensure MediaInfo CLI is installed at the specified path or update the `MediaInfocliPath` parameter accordingly.
- File copying requires appropriate permissions and sufficient disk space in the target destination.

## License

This script is licensed under the [MIT License](LICENSE).
