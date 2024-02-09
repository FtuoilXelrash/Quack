Attribute VB_Name = "Hub"
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, Optional ByVal hModule As Long = 0&, Optional ByVal dwFlags As Long = &H1) As Long
Private Declare Function sndPlaySoundA Lib "winmm.dll" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Function SendText(strOutput As String, intErrNo As Integer)
On Error GoTo err_SendText

    PluginSite.Hooks.AddChatText "QUACK> " & strOutput, 13, 0
    
    If (intErrNo = 1) Then
    
    Open App.Path & "\DeepErrorLog.txt" For Append Access Write As #1
    
    Print #1, " Ver: " & Hub.AppVersion & " (" & Now & ")"
    Print #1, strOutput
    Print #1, String(40, "-")
    Close #1
    End If
    
    Exit Function
err_SendText:
    PluginSite.Hooks.AddChatText "err_SendText: " & Err.Description, 13, 0
End Function

Public Function AppVersion() As String
  AppVersion = App.Major & "." & App.Minor & ".0." & App.Revision
End Function

' OLD WRITETOCHAT FUNCTION!
Public Sub WriteToChat(ByVal message As String, Color As Integer)

PluginSite.Hooks.AddChatText "QUACK> " & message, Color, 0

End Sub

Public Sub PlayExtWavFile(ByVal pFilename As String)
  On Error GoTo Hell
      Dim Ret As Long
'    WriteToChat "The External Wav file: " & pFilename & "", 6
  
  Dim check_file_exists As String
  check_file_exists = Dir$(pFilename)
  
  If check_file_exists = "" Then
    WriteToChat "EXTERANAL SOUND: " & pFilename & " WAS NOT FOUND!", 10
'    pFilename = App.Path & "\default.wav"

  Else
    Ret = sndPlaySoundA(pFilename, SND_ASYNC Or SND_NODEFAULT)
  End If
  Exit Sub
Hell:
    Hub.SendText "err_PlayExtWavFile: " & Err.Description, 1
End Sub
