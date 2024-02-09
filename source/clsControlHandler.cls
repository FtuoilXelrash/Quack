VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControlHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'   ABOUT
Private WithEvents lblAboutBuildDate As DecalControls.StaticText
Attribute lblAboutBuildDate.VB_VarHelpID = -1
'Public imgLogo As DecalControls.FixedLayout
'Public imgBGPic As DecalControls.FixedLayout

Public WithEvents chkEatRedMessages As DecalControls.CheckBox
Attribute chkEatRedMessages.VB_VarHelpID = -1
Public WithEvents chkAniIdle As DecalControls.CheckBox
Attribute chkAniIdle.VB_VarHelpID = -1
Public WithEvents chkLogOnDeath As DecalControls.CheckBox
Attribute chkLogOnDeath.VB_VarHelpID = -1
Public WithEvents chkPlayWavOnDeath As DecalControls.CheckBox
Attribute chkPlayWavOnDeath.VB_VarHelpID = -1

' Initialize the controls used in the Plug-in GUI
Public Function InitControls() As Boolean
On Error GoTo err_dead
    With view.MainView
       
    Set lblAboutBuildDate = .control("lblAboutBuildDate")
'    Set imgLogo = .control("imgLogo")
'    Set imgBGPic = .control("imgBGPic")
    
    Set chkEatRedMessages = .control("chkEatRedMessages")
    Set chkAniIdle = .control("chkAniIdle")
    Set chkLogOnDeath = .control("chkLogOnDeath")
    Set chkPlayWavOnDeath = .control("chkPlayWavOnDeath")
        
    End With
    Call SetRandomLogo
    Call LoadSettings
    
    Exit Function
    
err_dead:
    Hub.SendText "err_InitControls: " & Err.Description, 1
End Function

Public Sub SetRandomLogo()
       Dim TheImage As IImageCache
       
'       Set imgLogo.Background = pluginSiteOld.LoadBitmapPortal(&H600118A)
'       Set imgBGPic.Background = pluginSiteOld.LoadBitmapPortal(&H6006067)
       Set TheImage = Nothing
       Exit Sub
End Sub

Public Function LoadSettings()
On Error GoTo err_dead

    Call LoadRegSetting
    Call ChangeStaticText(view.MainView.control("lblAboutBuildDate"), "" & _
    "Plug-in Version: " & Hub.AppVersion)
    Exit Function
    
err_dead:
    Hub.SendText "err_LoadSettings: " & Err.Description, 1
End Function

Public Sub LoadRegSetting()
On Error GoTo err_dead

  With reg
    .Root = HKEY_LOCAL_MACHINE
    .Key = "Software\Decal Plugins\QUACK"
    
    control.chkEatRedMessages.Checked = CBool(.Value("chkEatRedMessages"))
    control.chkAniIdle.Checked = CBool(.Value("chkAniIdle"))
    control.chkLogOnDeath.Checked = CBool(.Value("chkLogOnDeath"))
    control.chkPlayWavOnDeath.Checked = CBool(.Value("chkPlayWavOnDeath"))
    
      End With
      Exit Sub
err_dead:
    Hub.SendText "LoadRegSetting: " & Err.Description, 1
End Sub

Public Sub ChangeStaticText(ByRef TheControl As DecalControls.StaticText, ByVal TheText As String)
    TheControl.Text = TheText
End Sub

'   CHANGE REG SETTINGS AND ALERT USER WITH MESSAGE
Public Sub CheckSaved(ByVal WhatChecked As String, Optional ByVal OnOff As String)
    WriteToChat "The Change to " & WhatChecked & " setting was saved " & OnOff & "", 13
End Sub

Private Sub chkEatRedMessages_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkEatRedMessages") = CInt(bChecked)
  
    If chkEatRedMessages.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Option Eat Red Messages", Saved)
End Sub

Private Sub chkAniIdle_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkAniIdle") = CInt(bChecked)
  
    If chkAniIdle.Checked = True Then
    Saved = "ON"
    Call hook.achooks.SetIdleTime(31536000)
    WriteToChat "Client Idle Time Set to 3 Years!", 10
    
    Else
    Saved = "OFF"
    Call hook.achooks.SetIdleTime(900)
    WriteToChat "Client Idle Reset to 15 minutes", 10
    End If
  
  Call CheckSaved("Option Anti Idle", Saved)
End Sub

Private Sub chkLogOnDeath_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkLogOnDeath") = CInt(bChecked)
  
    If chkLogOnDeath.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Option Log Off On Death", Saved)
End Sub

Private Sub chkPlayWavOnDeath_Change(ByVal nID As Long, ByVal bChecked As Boolean)
  reg.Value("chkPlayWavOnDeath") = CInt(bChecked)
  
    If chkPlayWavOnDeath.Checked = True Then
    Saved = "ON"
    Else
    Saved = "OFF"
    End If
  
  Call CheckSaved("Option Play Sound On Death", Saved)
End Sub







