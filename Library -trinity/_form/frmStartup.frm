VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStartup.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   840
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrStartup_Timer()

End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
