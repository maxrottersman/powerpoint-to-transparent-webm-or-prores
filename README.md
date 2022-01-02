# Powerpoint to WebM Transparent Video
Exports slides as transparent webm or mov files for obs, youtube, or other streaming

How it works
1. A graphic is placed on a slide with the background color set to solid green 0x00FF00 or blue 0x00000FF
2. The graphic is animated (if desired)
3. It is exported as a 4K MP4 video (MOV works too, though I see no difference)
4. The MP4 video is then converted to a WebM video file using FFMPEG

## Installation
Install FFMPEG
https://www.ffmpeg.org/download.html

Set the path to ffmpeg.exe in the Powerpoint Slide 1
Also set folders for which to place MP4 and WebM exports

## Known Issues
There is green or blue fringing around graphics that have shaddows or motion, etc. Best to pick the background color that you will mind the least if this is an issue.  I've tried some ffmpeg tricks to fix this but no luck.  Hopefully someone can solve.

## Basic/Current FFMPEG params
"file.mp4" -b:v 2M -vf "chromakey=0x00FF00:.20:.15,format=rgba,scale=1920x1080" -c:v libvpx-vp9 -auto-alt-ref 0 "ouput.webm"

For ProRes, Davinci Resolve won't work with WebM, currently (i'm sure this could be improved)
" file.mp4" -c:v qtrle -vf "chromakey=0x00FF00:.2:.15,format=rgba,scale=1920x1080" "output.mov"




