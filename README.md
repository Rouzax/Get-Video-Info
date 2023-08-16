markdown
# Video Information Filter and Copy Script

This PowerShell script allows you to filter and copy video files based on various criteria. It uses FFprobe to extract video information and then filters the files based on codec, bit rate, dimensions, and more.

## Prerequisites

- PowerShell
- FFprobe executable (part of FFmpeg) - Ensure that FFprobe is installed and accessible through the provided path or update the path to the FFprobe executable.

## Usage

1. Clone or download this repository to your local machine.

2. Open a PowerShell terminal.

3. Navigate to the directory containing the script using the `cd` command.

4. Run the script using the following command:

```powershell
.\Get-Video-Codec.ps1 -FolderPath <FolderPath> [Options]
```

Replace `<FolderPath>` with the path to the root folder containing your video files.

### Options

- `-FFprobePath` (Optional): Path to the FFprobe executable. Defaults to "C:\Program Files\FFmpeg\ffprobe.exe".

- `-Recursive` (Optional): Include this switch to enable recursive search through subdirectories.

- `-CodecFilter` (Optional): Filter videos by codec name.

- `-MinBitrate` and `-MaxBitrate` (Optional): Filter videos by minimum and maximum bit rate.

- `-MinWidth`, `-MaxWidth`, `-MinHeight`, `-MaxHeight` (Optional): Filter videos by dimensions.

- `-ExactWidth` and `-ExactHeight` (Optional): Filter videos by exact dimensions.

- `-EncoderFilter` (Optional): Filter videos by encoder value.

- `-TargetDestination` (Optional): Specify the target destination folder to copy the filtered videos. Videos will be copied while maintaining the folder structure.

## Examples

Filter and display video information:

```powershell
.\Get-Video-Codec.ps1 -FolderPath "C:\Videos" -CodecFilter "h264" -MinWidth 1920 -MaxBitrate 8000000
```

Filter and copy videos to a specific destination:

```powershell
.\Get-Video-Codec.ps1 -FolderPath "C:\Videos" -CodecFilter "h264" -TargetDestination "D:\FilteredVideos"
```

## Note

- Make sure FFprobe is correctly installed and the provided path is accurate.

- Always review and test the script on a small subset of files before using it on a large collection.

- This script provides a simple way to filter and copy videos, but it might not cover all edge cases.

- Use at your own risk. The author is not responsible for any data loss or damage caused by using this script.
```

Feel free to modify the README.md file according to your preferences and include any additional information you find relevant.