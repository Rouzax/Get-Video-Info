# Get Video Information PowerShell Script

This PowerShell script allows you to extract information from video files in a specified folder and its subfolders. It uses FFprobe to gather video details such as codec, bit rate, width, height, and more.

## Usage

1. Make sure you have FFprobe installed. You can download it as part of the FFmpeg package from https://ffmpeg.org/download.html.

2. Place the script (`Get-Video-Info.ps1`) and FFprobe executable (`ffprobe.exe`) in the same directory.

3. Open a PowerShell terminal.

4. Navigate to the directory containing the script and FFprobe executable.

5. Run the script using the following commands:

   ```powershell
   # Basic usage
   .\Get-Video-Info.ps1 -FolderPath "C:\Path\To\Your\Videos"

   # Additional options
   .\Get-Video-Info.ps1 -FolderPath "C:\Path\To\Your\Videos" -Recursive -CodecFilter "h264" -MinBitrate 1000000 -MaxBitrate 5000000 -MinWidth 1280 -MaxWidth 1920 -MinHeight 720 -MaxHeight 1080 -ExactWidth 1920 -ExactHeight 1080 -TargetDestination "C:\Path\To\Copy\Videos"
   ```

Parameters

    -FolderPath: The path to the folder containing the video files.

    -FFprobePath (Optional): Path to the FFprobe executable. Defaults to "C:\Program Files\FFmpeg\ffprobe.exe".

    -Recursive (Optional): Include subfolders in the search.

    -CodecFilter (Optional): Filter videos by codec.

    -MinBitrate (Optional): Filter videos with a minimum bit rate (in bps).

    -MaxBitrate (Optional): Filter videos with a maximum bit rate (in bps).

    -MinWidth, -MaxWidth (Optional): Filter videos by width range.

    -MinHeight, -MaxHeight (Optional): Filter videos by height range.

    -ExactWidth, -ExactHeight (Optional): Filter videos by exact width and height.

    -TargetDestination (Optional): Copy filtered videos to the specified destination path. If the path does not exist, it will be created.

Note: If you use filters, only videos meeting the specified criteria will be displayed or copied.
Example

To extract information from all videos in the folder C:\Videos and its subfolders, and copy videos with the codec "h264" and bit rate between 1Mbps and 5Mbps to a new directory C:\FilteredVideos, use the following command:

```powershell
.\Get-Video-Info.ps1 -FolderPath "C:\Videos" -Recursive -CodecFilter "h264" -MinBitrate 1000000 -MaxBitrate 5000000 -TargetDestination "C:\FilteredVideos"
```

Remember to adjust the paths and filters according to your needs.
