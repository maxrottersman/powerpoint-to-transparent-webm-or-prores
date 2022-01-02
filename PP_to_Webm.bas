Attribute VB_Name = "PP_to_Webm"

Global gSlideNumber ' when we only want to export one slide

Sub Export_Only_One_Slide()

gSlideNumber = Val(ActivePresentation.Slides(1).Shapes("txt_export_only_one_slide").TextFrame.TextRange.Text)

' Now run export routine
a__Create_Transparent_Webm_For_Each_Slide


End Sub

Sub a__Create_Transparent_Webm_For_Each_Slide()

Dim Slide As Object
Dim sPPPath, sMP4Path, sWEBMPath, sSlide As String
Dim sMP4FilNameAndPath, sWEBMFileNameAndPath As String
Dim sFFMPEG_config_1, sFFMPEG_WithPath, sFileExt As String
Dim sSlideName, sNotes As String
Dim iSlideCount, iSlideExport As Integer

' Create paths to folders
' Assume we have sub-folders off powerpoint folder \ffmpeg \mp4 and \webm
sPPPath = ActivePresentation.Path
sFFMPEG_WithPath = sPPPath & "\ffmpeg\ffmpeg -i "
sMP4Path = sPPPath & "\mp4\"
'sMP4Path = sPPPath & "\wmv\"
sWEBMPath = sPPPath & "\webm\"
'sWEBMPath = sPPPath & "\mkv\"
'sWEBMPath = sPPPath & "\webm_wmv\"
sFileExt = ".webm"
'sFileExt = ".mkv"
Dim sHEXColor As String
sHEXColor = "0x00FF00" ' default
sChromaSimilarity = ".2"
sChromaBlend = ".15"

' ffmpeg notes
' https://trac.ffmpeg.org/wiki/Encode/VP9
' -y = force overwrite
' -pix_fmt yuva420p ' don't need
' -metadata:s:v:0
' "-lossless 1 ' looked worse
' '  -b:v 0 -crf 30  ' constant quality
' or libvpx
' alpha_mode=" & _
'    Chr(34) & "1" & Chr(34) & " -auto-alt-ref 0
' Gray didn't work: 7E7E7E
' chromakey=:similarity:blend
' v1: "chromakey=0x00FF00:.20:.15,scale=1920x1080" & Chr(34) & _
' v2: "chromakey=0x00Fe00:.5.10,scale=1920x1080" & Chr(34) & _ ' didn't work
' Paul, check later: "colorkey=0x00fe00:0.5" -c:v png


' for FFMPEG
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

iSlideCount = ActivePresentation.Slides.Count
' We assume 1st slide is our notes
iSlideStart = 2

' Have we set our global single slide number, if so only export that
If gSlideNumber > 2 Then
iSlideStart = gSlideNumber
iSlideEnd = gSlideNumber
' turn off
gSlideNumber = 0
    Else
    iSlideStart = 3
    iSlideEnd = ActivePresentation.Slides.Count
End If

    ' *** Export to mp4 Loop ***
    For p = iSlideStart To iSlideEnd '2 To iSlideCount
    
        ' Loop in a loop to take the item in our first look, say 2 in 4
        ' and HIDE all the other slides, so in this loop slides 1,3 and 4
        ' would be hidden leaving only slide 2 that will get exported as movie
        For iExport = 1 To iSlideCount
            If p <> iExport Then
            ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoTrue
                Else
                ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoFalse
            End If
        Next ' slide to hide/unhide
        
        ' Get background color and use that to make transparent in ffmpeg
        tmp = Hex(ActivePresentation.Slides(p).Background.Fill.ForeColor.RGB)
        If tmp = "FF0000" Then sHEXColor = "0xFF0000"
        If tmp = "00FF00" Then sHEXColor = "0x00FF00"
        If tmp = "0000FF" Then sHEXColor = "0x0000FF"
        
        'Generate ffmpeg command
        sFFMPEG_config_1 = " -b:v 2M -vf " & Chr(34) & _
         "chromakey=" & sHEXColor & ":" & _
         sChromaSimilarity & ":" & _
         sChromaBlend & ",format=rgba,scale=1920x1080" & Chr(34) & _
         " -c:v libvpx-vp9 -auto-alt-ref 0 -y "
        
        ' If we have notes, use it as the file name
        sNotes = ActivePresentation.Slides(p).NotesPage. _
            Shapes.Placeholders(2).TextFrame.TextRange.Text
        
        If Trim(sNotes) > 1 Then
        sMP4FilNameAndPath = sMP4Path & sNotes & ".mp4"
        sWEBMFileNameAndPath = sWEBMPath & sNotes & sFileExt '".webm" or ".mkv"
            Else
            sMP4FilNameAndPath = sMP4Path & "slide_" & Trim(Str(p)) & ".mp4"
            sWEBMFileNameAndPath = sWEBMPath & "slide_" & Trim(Str(p)) & sFileExt '".webm" or ".mkv"
            
        End If
        
        ' Now create movie of the not hidden slide '1080 or '2160 (4K)
        If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
            ActivePresentation.CreateVideo FileName:=sMP4FilNameAndPath, _
            UseTimingsAndNarrations:=True, _
            VertResolution:=2160, _
            FramesPerSecond:=30, _
            Quality:=100
        Else
            MsgBox "There is another conversion to video in progress"
        End If
        
        ' VBA won't wait for above, so we have to use kludge wait loop
        Do
        ' Don't tie up the user interface; add DoEvents
        ' to give the mouse and keyboard time to keep up.
        DoEvents
        Select Case ActivePresentation.CreateVideoStatus
            Case PpMediaTaskStatus.ppMediaTaskStatusDone
                'MsgBox "Conversion complete!"
                Exit Do
            Case PpMediaTaskStatus.ppMediaTaskStatusFailed
                'MsgBox "Conversion failed!"
                Exit Do
            Case PpMediaTaskStatus.ppMediaTaskStatusInProgress
                'Debug.Print "Conversion in progress"
            Case PpMediaTaskStatus.ppMediaTaskStatusNone
                ' You'll get this value when you ask for the status
                ' and no conversion is happening or has completed.
            Case PpMediaTaskStatus.ppMediaTaskStatusQueued
                'Debug.Print "Conversion queued"
        End Select
    Loop
        
        ' Okay let's send MP4 to ffmpeg to create WEBM file!
        sCMD_ffmpeg = sFFMPEG_WithPath & sMP4FilNameAndPath & " " & sFFMPEG_config_1 & sWEBMFileNameAndPath
               
        errorCode = wsh.Run(sCMD_ffmpeg, windowStyle, waitOnReturn)

            If errorCode = 0 Then
                'Insert your code here
            Else
                'MsgBox "Program exited with error code " & errorCode & "."
            End If
        
        Debug.Print "processed... " & p

Next ' p slide to make movie
    
    ' now unhide all, we'll start again with next p
    For iExport = 1 To iSlideCount
    ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoFalse
    Next ' slide to hide/unhide

'MsgBox "done"

End Sub

Sub a__Create_Transparent_MOV_For_Each_Slide()
' Assume WEBMs were created first (the source MP4 files)
' NOTE!!!
' We use this for video that will work in Davinci Resolve


Dim Slide As Object
Dim sPPPath, sMP4Path, sWEBMPath, sSlide As String
Dim sMP4FilNameAndPath, sWEBMFileNameAndPath As String
Dim sFFMPEG_config_1, sFFMPEG_WithPath, sFileExt As String
Dim sSlideName, sNotes As String
Dim iSlideCount, iSlideExport As Integer

' Create paths to folders
' Assume we have sub-folders off powerpoint folder \ffmpeg \mp4 and \webm
sPPPath = ActivePresentation.Path
sFFMPEG_WithPath = sPPPath & "\ffmpeg\ffmpeg -i "
sMP4Path = sPPPath & "\mp4\"
sWEBMPath = sPPPath & "\mov\"
sFileExt = ".mov"

' ffmpeg notes
' https://trac.ffmpeg.org/wiki/Encode/VP9
' -y = force overwrite
' -pix_fmt yuva420p ' don't need
' -metadata:s:v:0
' "-lossless 1 ' looked worse
' '  -b:v 0 -crf 30  ' constant quality
' or libvpx
' alpha_mode=" & _
'    Chr(34) & "1" & Chr(34) & " -auto-alt-ref 0
' chromakey=:similarity:blend
' v1: "chromakey=0x00FF00:.20:.15,scale=1920x1080" & Chr(34) & _
' v2: "chromakey=0x00Fe00:.5.10,scale=1920x1080" & Chr(34) & _ ' didn't work
' Paul, check later: "colorkey=0x00fe00:0.5" -c:v png
sFFMPEG_config_1 = " -c:v qtrle -vf " & Chr(34) & _
    "chromakey=0x00FF00:.2:.15,format=rgba,scale=1920x1080" & Chr(34) & _
    " -y "

' for FFMPEG
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

iSlideCount = ActivePresentation.Slides.Count
' We assume 1st slide is our notes
iSlideStart = 2

    ' *** Export to mp4 Loop ***
    For p = 2 To iSlideCount
    
        ' Loop in a loop to take the item in our first look, say 2 in 4
        ' and HIDE all the other slides, so in this loop slides 1,3 and 4
        ' would be hidden leaving only slide 2 that will get exported as movie
        For iExport = 1 To iSlideCount
            If p <> iExport Then
            ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoTrue
                Else
                ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoFalse
            End If
        Next ' slide to hide/unhide
        
        ' If we have notes, use it as the file name
        sNotes = ActivePresentation.Slides(p).NotesPage. _
            Shapes.Placeholders(2).TextFrame.TextRange.Text
        
        If Trim(sNotes) > 1 Then
        sMP4FilNameAndPath = sMP4Path & sNotes & ".mp4"
        sWEBMFileNameAndPath = sWEBMPath & sNotes & sFileExt '".webm" or ".mkv"
            Else
            sMP4FilNameAndPath = sMP4Path & "slide_" & Trim(Str(p)) & ".mp4"
            sWEBMFileNameAndPath = sWEBMPath & "slide_" & Trim(Str(p)) & sFileExt '".webm" or ".mkv"
            
        End If
        
        ' Now create movie of the not hidden slide '1080 or '2160 (4K)
        If False Then 'create MP4
        
        If ActivePresentation.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
            ActivePresentation.CreateVideo FileName:=sMP4FilNameAndPath, _
            UseTimingsAndNarrations:=True, _
            VertResolution:=2160, _
            FramesPerSecond:=30, _
            Quality:=100
        Else
            MsgBox "There is another conversion to video in progress"
        End If
        
        
        ' VBA won't wait for above, so we have to use kludge wait loop
        Do
        ' Don't tie up the user interface; add DoEvents
        ' to give the mouse and keyboard time to keep up.
        DoEvents
        Select Case ActivePresentation.CreateVideoStatus
            Case PpMediaTaskStatus.ppMediaTaskStatusDone
                'MsgBox "Conversion complete!"
                Exit Do
            Case PpMediaTaskStatus.ppMediaTaskStatusFailed
                'MsgBox "Conversion failed!"
                Exit Do
            Case PpMediaTaskStatus.ppMediaTaskStatusInProgress
                'Debug.Print "Conversion in progress"
            Case PpMediaTaskStatus.ppMediaTaskStatusNone
                ' You'll get this value when you ask for the status
                ' and no conversion is happening or has completed.
            Case PpMediaTaskStatus.ppMediaTaskStatusQueued
                'Debug.Print "Conversion queued"
        End Select
        Loop
        
        End If ' create mp4
        
        ' Okay let's send MP4 to ffmpeg to create WEBM file!
        sCMD_ffmpeg = sFFMPEG_WithPath & sMP4FilNameAndPath & " " & sFFMPEG_config_1 & sWEBMFileNameAndPath
               
        errorCode = wsh.Run(sCMD_ffmpeg, windowStyle, waitOnReturn)

            If errorCode = 0 Then
                'Insert your code here
            Else
                'MsgBox "Program exited with error code " & errorCode & "."
            End If
        
        Debug.Print "processed... " & p

Next ' p slide to make movie
    
    ' now unhide all, we'll start again with next p
    For iExport = 1 To iSlideCount
    ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoFalse
    Next ' slide to hide/unhide

'MsgBox "done"

End Sub

Public Sub To_default_txt_ffmpeg_folder()
'txt_mp4_export_folder

ActivePresentation.Slides(1).Shapes("txt_ffmpeg_folder").TextFrame.TextRange.Text = ActivePresentation.Path & "\ffmpeg"
ActivePresentation.Slides(1).Shapes("txt_mp4_export_folder").TextFrame.TextRange.Text = ActivePresentation.Path & "\mp4"
ActivePresentation.Slides(1).Shapes("txt_webm_export_folder").TextFrame.TextRange.Text = ActivePresentation.Path & "\webm"

End Sub

