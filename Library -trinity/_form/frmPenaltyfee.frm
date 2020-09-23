VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPenaltyfee 
   Caption         =   "Penalty Fee"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8970
   Begin VB.Frame fraList 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   9255
      Begin MSComctlLib.ListView lvList 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   0
      Picture         =   "frmPenaltyfee.frx":0000
      Top             =   0
      Width           =   8640
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmPenaltyfee.frx":8D17
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Fee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can view penalty fee per day."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   2475
   End
End
Attribute VB_Name = "frmPenaltyfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lvList.ColumnHeaders.Add , , " "
    lvList.ColumnHeaders.Add , , " "
    lvList.HideColumnHeaders = True
    lvList.ListItems.Add , , "Penalty Fee"
    lvList.ListItems(1).SubItems(1) = "P 5.00"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraList.Move 120, 2280, Me.ScaleWidth - 240, Me.ScaleHeight - (2280 + 120)
    lvList.Move 120, 240, fraList.Width - 240, fraList.Height - (240 + 120)
End Sub

