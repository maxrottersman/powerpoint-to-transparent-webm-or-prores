Rem %1 is the filename that is dropped onto this batch filename
Rem %~n1.webm creates an output file name with the extension changed to webm
Rem add -y if you want to automatically overwrite an existing file
::
:: Set our path to ffmpeg\ffmpeg
set path_ffmpeg=J:\Files2021_MaxoticsYouTube\20211224_PowerPointToWebmTransparentGraphics\GenClips2\ffmpeg\
::

%path_ffmpeg%ffmpeg -i %1 -b:v 2M -vf "chromakey=0x00FF00:.20:.15,format=rgba,scale=1920x1080" -c:v libvpx-vp9 -auto-alt-ref 0 "%~n1.webm"