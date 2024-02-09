VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents MainView As DecalPlugins.view
Attribute MainView.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Dim Schema As String
      Dim i%
        Schema = ViewSchemaOld()
   
   ' Initialize Decal Controls
   Set MainView = pluginSiteOld.LoadView(Schema)
End Sub

Public Function ViewSchemaOld() As String
On Error GoTo err_dead

    ViewSchemaOld = "<?xml version=""1.0""?>"
    ViewSchemaOld = ViewSchemaOld & "<view icon=""25162"" title=""QUACK!"" top=""55"" left=""3"" width=""215"" height=""150"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Notebook"" name=""QUACK"">"
               
    ViewSchemaOld = ViewSchemaOld & "<page label=""OPTIONS"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""QuackOpts"" clipped="""">"
                      
'   ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""imgBGPic"" id=""24679"" left=""0"" top=""5"" width=""215"" height=""190"">"
        
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkEatRedMessages"" left=""5"" top=""10"" width=""110"" height=""20"" text=""Eat Red Text Messages"" fontsize=""12"" textcolor=""255""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkAniIdle"" left=""5"" top=""30"" width=""110"" height=""20"" text=""Anit Idle"" fontsize=""12""/>"
              
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkLogOnDeath"" left=""5"" top=""50"" width=""110"" height=""20"" text=""Log Off On Death"" fontsize=""12""/>"
                         
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.Checkbox"" name=""chkPlayWavOnDeath"" left=""5"" top=""70"" width=""110"" height=""20"" text=""Play Sound On Death"" fontsize=""12""/>"
                         
'    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
    
    ViewSchemaOld = ViewSchemaOld & "<page label=""ABOUT"">"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""FsAbout"" clipped="""">"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName"" text=""QUACK"" top=""5"" left=""75"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutName2"" text=""For Decal v3"" top=""18"" left=""65"" width=""225"" height=""30"" fontsize=""14"" fontstyle=""bold"" textcolor=""255""/>"
    
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutContact"" text=""Created By: Ftuoil Xelrash of Darktide"" top=""35"" left=""5"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.StaticText"" name=""lblAboutBuildDate"" text=""Plug-in Version: "" top=""50"" left=""30"" width=""210"" height=""20"" fontsize=""14"" fontstyle=""bold""/>"
'    ViewSchemaOld = ViewSchemaOld & "<control progid=""DecalControls.FixedLayout"" name=""imgLogo"" id=""4567"" left=""60"" top=""75"" width=""80"" height=""120"">"
     
'    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</page>"
       
    ViewSchemaOld = ViewSchemaOld & "</control>"
    ViewSchemaOld = ViewSchemaOld & "</view>"
    
    Exit Function

err_dead:
    Hub.SendText "err_ViewSchemaOld: " & Err.Description, 1
End Function

' Clear the MainView
Private Sub Class_Terminate()
On Error GoTo err_dead

  Set MainView = Nothing
  Exit Sub
  
err_dead:
    Hub.SendText "err_Class_terminate: " & Err.Description, 1
End Sub
