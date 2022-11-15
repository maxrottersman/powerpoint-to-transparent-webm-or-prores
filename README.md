# Powerpoint to WebM or ProRes 4444 Transparent Video
Exports slides as transparent webm or mov files for obs, Kdenlive, Premiere, Resolve Youtube, or other video editors or streaming solutions.

How it works
1. Graphics are placed on a slide with the background color set to solid green 0x00FF00 or blue 0x00000FF
2. The graphics areanimated (if desired)
3. Each desired slide is exported as a 4K MP4 video 
4. The MP4 video is then converted to and/or a WebM/VP9 or ProRes 4444 video file using FFMPEG

## Installation
Install FFMPEG
https://www.ffmpeg.org/download.html

Set the path to ffmpeg.exe in the Powerpoint Slide 1
Also set folders for which to place MP4 and WebM exports

## Known Issues
There is green or blue fringing around graphics that have shadows or motion, etc. Best to pick the background color that you will mind the least, if this is an issue.  I've tried some ffmpeg tricks to fix this but no luck.  Hopefully someone can solve.

Generate ffmpeg command *** WEBM / VP9 ***

sFFMPEG_config_webm = " -b:v 2M -vf " & Chr(34) & _
 "chromakey=" & sHEXColor & ":" & _
 sChromaSimilarity & ":" & _
 sChromaBlend & ",format=rgba,scale=" & txt_output_scale & Chr(34) & _
 " -c:v libvpx-vp9 -auto-alt-ref 0 -y "

Generate ffmpeg command *** MOV 4444 with Alpha ***
sFFMPEG_config_mov4444 = " -c:v prores_ks -profile:v 4 -vendor apl0 -bits_per_mb 8000  -vf " & Chr(34) & _
 "chromakey=" & sHEXColor & ":" & _
 sChromaSimilarity & ":" & _
 sChromaBlend & ",format=rgba,scale=" & txt_output_scale & Chr(34) & _
 " -pix_fmt yuva444p10le "![image](https://user-images.githubusercontent.com/512477/200072982-881a4405-a507-4154-9c9e-e5d017059c01.png)


My YouTube is "Maxotics".  I can also be reached at maxotics.com




