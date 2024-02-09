VERSION 5.00
Begin VB.Form Timers 
   Caption         =   "Timers"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Timers"
   ScaleHeight     =   1080
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer LogOffDeath 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   480
      Top             =   240
   End
End
Attribute VB_Name = "Timers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LogOffDeath_Timer()

Timers.LogOffDeath.Enabled = False
hook.achooks.Logout

End Sub
