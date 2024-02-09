VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPlugin2

' Initialize the plug-in
Private Sub IPlugin2_Initialize(ByVal Site As Decal.IPluginSite2)
On Error GoTo err_dead
   
    boolSettingsLoaded = False
    boolLoggedIn = False
    Set PluginSite = Site
    Set pluginSiteOld = Site.PluginSite
    Set view = New clsViewHandler
    Set control = New clsControlHandler
    Set hook.achooks = Site.Hooks
    Set reg = New RegOp
   
    Set Csf.CSFilter = Site.Object("services\DecalNet.NetService\DecalFilters.CharacterStats")
'    Set world.wfilter = Site.Object("services\DecalNet.NetService\DecalFilters.world")
'    Set echo.NetEcho = Site.Object("services\DecalNet.NetService\DecalFilters.EchoFilter2")
    
    Call control.InitControls
    Exit Sub

err_dead:
    Hub.SendText "err_IPlugin2_Initialize: " & Err.Description, 1
End Sub

' Terminate the plug-in
Private Sub IPlugin2_Terminate()
On Error GoTo err_dead
   
'   KILL THE FORMS
    Set PluginSite = Nothing
    Set view.MainView = Nothing
    Set view = Nothing
    Set Csf.CSFilter = Nothing
'    Set world.wfilter = Nothing
    Set reg = Nothing
'    Set world = Nothing
    Set hook.achooks = Nothing
    Set control = Nothing
    Timers.LogOffDeath.Enabled = False
            
    Exit Sub
    
err_dead:
    Hub.SendText "err_IPlugin2_Terminate", 1
End Sub
