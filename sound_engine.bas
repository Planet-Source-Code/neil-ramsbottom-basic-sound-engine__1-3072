Attribute VB_Name = "Sound_Engine"

'==============================================
'|PlaySound Engine Copyright ©1999 UnpreXisten|
'==============================================
'|This code is FREEWARE and may be distributed|
'|so long as no charge is made for it and     |
'|the UnpreXisten name is mentioned.          |
'==============================================

'Use this sub, call it like so:
    
    'To play a pre-defined event
    'Playsound 1
    
    'To play a file not specified in code
    'Playsound 0,"C:\server\sound\candy.wav" 'For local access
    'Playsound 0,"\\office_Server\sound\candy.wav" 'For LAN access



Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public SoundActive As Boolean 'Its public, so the rest of the prog can interrogate it.

Public Sub PlaySound(sndEvent As Integer, Optional sndFilename As String)

Dim filename As String

'SoundActive=True 'Use this if you want sound on always


If SoundActive And sndFilename = "" Then 'SoundActive should be used to allow
                                         'the user to de-activate sound(s)
Select Case sndEvent
    
    Case 1: filename = "C&C Beep.wav"
    Case 2: filename = "C&C Bleep.wav"
    Case 3: filename = "C&C Cloaking.wav"
    Case 4: filename = "C&C critical stop.wav"
    Case 5: filename = "C&C default sound.wav"
    Case 6: filename = "C&C empty recycle (GDI).wav"
    Case 7: filename = "C&C EVA atmosphere.wav"
    Case 8: filename = "C&C EVA Base under attk.wav"
    Case 9: filename = "C&C EVA Canceled.wav"
    Case 10: filename = "C&C EVA Low power.wav"
    Case 11: filename = "C&C EVA Nuke approaching.wav"
    Case 12: filename = "C&C EVA Nuke available.wav"
    Case 13: filename = "C&C EVA Reinforcements.wav"
    Case 14: filename = "C&C EVA Select target.wav"
    Case 15: filename = "C&C EVA Silos needed.wav"
    Case 16: filename = "C&C exclamation.wav"
    Case 17: filename = "C&C exit windows (GDI).wav"
    Case 18: filename = "C&C Got a present 4 u.wav"
    Case 19: filename = "C&C Keystroke.wav"
    Case 20: filename = "C&C Map up.wav"
    
Case Else
    MsgBox "Error, invalid sound event passed to PlaySound function!", vbExclamation, "Error"
    Exit Sub
End Select
    
    If Right(App.Path, 2) = ":\" Then 'Is a drive (root like off a cd)
        'filename = App.Path & "data\sfx\" & filename
        filename = App.Path & filename 'Miss the '\' because the
                                       'App.Path has it
        
    Else
        'filename = App.Path & "\data\sfx\" & filename
        filename = App.Path & "\" & filename
                                        'Add the '\' because the
                                        'App.Path doesnt have it
    End If

End If

If filename = "" Then
    
    filename = sndFilename 'Set the filename to the Optional one

End If

On Error GoTo err 'Only error here can be a Sound System problem.

sndPlaySound filename, 3 'The 3 prevents the system from freezing during playback

Exit Sub

err:
    'Be silent for the error, its probably a windows thing (i.e. no S/C)

End Sub

