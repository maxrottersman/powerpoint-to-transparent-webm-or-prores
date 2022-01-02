# Powerpoint to WebM Transparent Video
Exports slides as transparent webm or mov files for obs, youtube, or other streaming

How it works
1. An animation of graphic is place on a slide with the background color set to solid green 0x00FF00 or blue 0x00000FF
2. It is exported as a 4K MP4 (MOV works too though I see no difference) video
3. The MP4 video is then converted to a WebM video file using FFMPEG

## Known Issues
There is green or blue fringing around graphics that have shaddows or motion, etc. Best to pick the background color that you will mind the least if this is an issue.  I've tried some ffmpeg tricks to fix this but no luck.  Hopefully someone can solve.


