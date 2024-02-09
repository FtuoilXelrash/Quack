VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents CSFilter As DecalFilters.CharacterStats
Attribute CSFilter.VB_VarHelpID = -1

Private Sub CSFilter_LoginComplete()
On Error GoTo err_dead

' Load locations, wand guid
'   SET THE DEBUG FLAG
'    DEBUG_MODE = True
    DEBUG_MODE = False
 
'  SET LOGINCOMPLETED FLAG
    bLoginCompleted = True
      
'   WELCOME KIT
WriteToChat "QUACK Ver. " & Hub.AppVersion & " By: Ftuoil Xelrash of Darktide!", 4
WriteToChat "Welcome to the world of " & Csf.CSFilter.Server & "! There is currently " & Csf.CSFilter.ServerPopulation & " players connected at this time.", 1
 
PlayerGUID = Csf.CSFilter.Character

If control.chkAniIdle.Checked = True Then
  Call hook.achooks.SetIdleTime(31536000)
WriteToChat "Client Idle Time Set to 3 Years!", 10
End If

    Exit Sub
    
err_dead:
    Hub.SendText "ERR_CSFilter_LoginComplete", 1
End Sub

Private Sub CSFilter_Death(ByVal bstrText As String)

'WriteToChat "I JUST DIED! - " & bstrText, 10

    If control.chkLogOnDeath.Checked = True Then
    WriteToChat "You just died and Log Off on Death is on!", 10
    
    Timers.LogOffDeath.Enabled = True
    
    End If
    
    If control.chkPlayWavOnDeath.Checked = True Then
    Hub.PlayExtWavFile (App.Path & "\" & "quack.wav")
    End If

    Exit Sub
    
err_dead:
    Hub.SendText "err_CSFilter_Death", 1
End Sub

