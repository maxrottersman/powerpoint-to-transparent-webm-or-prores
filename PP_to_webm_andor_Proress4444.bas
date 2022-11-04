Attribute VB_Name = "PP_to_webm_andor_Proress4444"
Sub Generate_Transparent_ProRes4444_AndOr_Webm()

Dim Slide As Object
Dim sPPPath, sMP4Path, sWEBMPath, sSlide As String
Dim sMP4FilNameAndPath, sWEBMFileNameAndPath As String
Dim sFFMPEG_config_1, sFFMPEG_WithPath, sFileExt As String
Dim sSlideName, sNotes As String
Dim iSlideCount, iSlideExport, iSlideNumberIfOnlyProcessOne As Integer

' *********** BEGIN CONFIT *************
' Get settings from CONFIG page
    ' If we only want to process one page, which one is it.  Otherwise set to 0
    If Val(ActivePresentation.Slides(1).Shapes("txt_export_only_one_slide").TextFrame.TextRange.Text) > 0 Then
    iSlideNumberIfOnlyProcessOne = Val(ActivePresentation.Slides(1).Shapes("txt_export_only_one_slide").TextFrame.TextRange.Text)
        Else
        iSlideNumberIfOnlyProcessOne = 0
    End If
    
' folder to write files
txt_ffmpeg_folder = ActivePresentation.Slides(1).Shapes("txt_ffmpeg_folder").TextFrame.TextRange.Text
txt_mp4_folder = ActivePresentation.Slides(1).Shapes("txt_mp4_folder").TextFrame.TextRange.Text
txt_webm_folder = ActivePresentation.Slides(1).Shapes("txt_webm_folder").TextFrame.TextRange.Text
txt_mov_folder = ActivePresentation.Slides(1).Shapes("txt_mov_folder").TextFrame.TextRange.Text
txt_output_types = ActivePresentation.Slides(1).Shapes("txt_output_types").TextFrame.TextRange.Text
txt_output_scale = ActivePresentation.Slides(1).Shapes("txt_output_scale").TextFrame.TextRange.Text
txt_jpg_folder = ActivePresentation.Slides(1).Shapes("txt_jpg_folder").TextFrame.TextRange.Text

    If Len(txt_output_scale) < 4 Then
    txt_output_scale = "1920x1080' 'defalt"
    End If
    
' *********** END CONFIG ***************

' Create paths to folders
' Assume we have sub-folders off powerpoint folder \ffmpeg \mp4 and \webm
'sPPPath = ActivePresentation.Path
sFFMPEG_WithPath = txt_ffmpeg_folder & "\ffmpeg -i "
sMP4Path = txt_mp4_folder & "\"
sWEBMPath = txt_webm_folder & "\"
sMOVPath = txt_mov_folder & "\"
sJPGPath = txt_jpg_folder & "\"

Dim sHEXColor As String
sHEXColor = "0x00FF00" ' default
sChromaSimilarity = ".2"
sChromaBlend = ".15"

'Generate ffmpeg command *** WEBM ***
sFFMPEG_config_webm = " -b:v 2M -vf " & Chr(34) & _
 "chromakey=" & sHEXColor & ":" & _
 sChromaSimilarity & ":" & _
 sChromaBlend & ",format=rgba,scale=" & txt_output_scale & Chr(34) & _
 " -c:v libvpx-vp9 -auto-alt-ref 0 -y "

'Generate ffmpeg command *** MOV 4444 with Alpha ***
sFFMPEG_config_mov4444 = " -c:v prores_ks -profile:v 4 -vendor apl0 -bits_per_mb 8000  -vf " & Chr(34) & _
 "chromakey=" & sHEXColor & ":" & _
 sChromaSimilarity & ":" & _
 sChromaBlend & ",format=rgba,scale=" & txt_output_scale & Chr(34) & _
 " -pix_fmt yuva444p10le "

' for FFMPEG
Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1
Dim errorCode As Integer

iSlideCount = ActivePresentation.Slides.Count
' We assume 1st slide is our CONFIG, and 2nd our notes, so we start on 3
iSlideStart = 3

' Make sure first slides are hidden
    For h = 1 To iSlideStart - 1
    ActivePresentation.Slides(h).SlideShowTransition.Hidden = msoTrue
    Next

    ' Start processing for each slide (start at 3 in this case)
    For p = iSlideStart To ActivePresentation.Slides.Count '2 To iSlideCount
    
        ' Only if we want one slide or we want all
        If iSlideNumberIfOnlyProcessOne = p Or iSlideNumberIfOnlyProcessOne = 0 Then
    
        ' Loop in a loop to take the item in our first look, say 2 in 4
        ' and HIDE all the other slides, so in this loop slides 1,3 and 4
        ' would be hidden leaving only slide 2 that will get exported as movie
            For iExport = iSlideStart To ActivePresentation.Slides.Count
                ' If p of master loop,doesn't each iExport of this look, then hide it
                ' end results, hides everything but current p.  Is there a smart way
                ' to do this probably, but I'm a dummy
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
        
        ' If we have notes, use it as the file name
        sNotes = ActivePresentation.Slides(p).NotesPage. _
            Shapes.Placeholders(2).TextFrame.TextRange.Text
        
            If Trim(sNotes) > 1 Then
            sMP4FilNameAndPath = sMP4Path & sNotes & ".mp4"
            sWEBMFileNameAndPath = sWEBMPath & sNotes & ".webm" '".webm" or ".mkv"
            sMOVFileNameAndPath = sMOVPath & sNotes & ".mov"
            sJPGFileNameAndPath = sJPGPath & sNotes & ".jpg"
                Else
                sMP4FilNameAndPath = sMP4Path & "slide_" & Trim(Str(p)) & ".mp4"
                sWEBMFileNameAndPath = sWEBMPath & "slide_" & Trim(Str(p)) & ".webm" '".webm" or ".mkv"
                sMOVFileNameAndPath = sMOVPath & "slide_" & Trim(Str(p)) & ".mov"
                sJPGFileNameAndPath = sJPGPath & "slide_" & Trim(Str(p)) & ".jpg" ' prob never fire
                
            End If
        
                ' Now create movie of the not hidden slide '1080 or '2160 (4K)
                ' MP4 !
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
            
                ' Okay let's send MP4 to ffmpeg to create MOV file!
                If InStr(UCase(txt_output_types), "MOV") > 0 Then
                
                sCMD_ffmpeg = sFFMPEG_WithPath & sMP4FilNameAndPath & " " & sFFMPEG_config_mov4444 & sMOVFileNameAndPath
                    
                    ' Delete previous vid file if exists
                    If Len(Dir$(sMOVFileNameAndPath)) > 0 Then
                    Kill sMOVFileNameAndPath
                    End If
                       
                errorCode = wsh.Run(sCMD_ffmpeg, windowStyle, waitOnReturn)
                    If errorCode = 0 Then
                        'Insert your code here
                    Else
                        'MsgBox "Program exited with error code " & errorCode & "."
                    End If
                Debug.Print "processed mov4444... " & p
                End If
                
                ' Okay let's send MP4 to ffmpeg to create MOV file!
                If InStr(UCase(txt_output_types), "WEBM") > 0 Then
                sCMD_ffmpeg = sFFMPEG_WithPath & sMP4FilNameAndPath & " " & sFFMPEG_config_webm & sWEBMFileNameAndPath
                
                    If Len(Dir$(sWEBMFileNameAndPath)) > 0 Then
                    Kill sWEBMFileNameAndPath
                    End If
                       
                errorCode = wsh.Run(sCMD_ffmpeg, windowStyle, waitOnReturn)
                    If errorCode = 0 Then
                        'Insert your code here
                    Else
                        'MsgBox "Program exited with error code " & errorCode & "."
                    End If
                Debug.Print "processed webm... " & p
                End If
          
        
        End If ' single slide output or all

Next ' p slide to make movie
    
    ' now unhide all, we'll start again with next p
    For iExport = 1 To iSlideCount
    ActivePresentation.Slides(iExport).SlideShowTransition.Hidden = msoFalse
    Next ' slide to hide/unhide
    

  ' Start processing for each slide (start at 3 in this case)
   
    '*************************
    ' Save JPG (thumbnail) if any
    ' ************************
    For p = iSlideStart To ActivePresentation.Slides.Count '2 To iSlideCount
    
    sNotes = ActivePresentation.Slides(p).NotesPage. _
            Shapes.Placeholders(2).TextFrame.TextRange.Text
    
        If InStr(UCase(sNotes), "JPG") > 0 Then
        
        sJPGFileNameAndPath = sJPGPath & sNotes & ".jpg"
            
            If Len(Dir$(sJPGFileNameAndPath)) > 0 Then
            Kill sJPGFileNameAndPath
            End If
            
            With Application.ActivePresentation.Slides(p)
            .Export sJPGFileNameAndPath, "JPG", 1920, 1080
            End With
        End If
    
    Next

MsgBox "Done Processing"

End Sub




