VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents achooks As Decal.achooks
Attribute achooks.VB_VarHelpID = -1


Private Sub AChooks_StatusTextIntercept(ByVal bstrText As String, bEat As Boolean)
On Error GoTo err_dead
     
'   LOOKING FOR THE DAMN RED TEXT CRAP
            
    If control.chkEatRedMessages.Checked = True Then
'    WriteToChat "RED TEXT SPAM: " & bstrText, 2
    bEat = True
    End If
       
    Exit Sub
    
err_dead:
    Hub.SendText "err_StatusTextIntercept: " & Err.Description, 1
End Sub
