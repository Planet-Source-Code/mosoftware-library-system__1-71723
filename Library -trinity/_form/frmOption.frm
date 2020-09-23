VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame fraDR 
      Height          =   3015
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtNS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkStat 
         Caption         =   "On/Off Pop Up  Checker of Products in Critical Level."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can enabled and Disabled Pop Up check of Product on Critical Level."
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
         Left            =   840
         TabIndex        =   9
         Top             =   480
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmOption.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Option Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   " Ex: 3000 equivalent to 3 seconds. Value must not greather than 65000."
         Height          =   465
         Left            =   1560
         TabIndex        =   7
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "No. of Seconds to Pop-Up"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Please Select Date to view List of Yearly Sales Report and Auto Generate second Date."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   5055
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   3
         Left            =   5280
         Picture         =   "frmOption.frx":617A
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Get_PopUp_Val()
    Dim sSQL As String
    sSQL = "SELECT tbl_popup.popuptime, tbl_popup.popupstat " & _
        "FROM tbl_popup;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, 3, 3
            chkStat.Value = adoRes.Fields("popupstat")
            txtNS.Text = adoRes.Fields("popuptime")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub chkStat_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    UPDATE_DATA2 "[tbl_popup]", "popuptime", "", txtNS.Text
    UPDATE_DATA2 "[tbl_popup]", "popupstat", "", chkStat.Value
    If chkStat.Value = 1 Then
        frmMain.tmrPopup = True
    End If
    cmdApply.Enabled = False
End Sub

Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Get_PopUp_Val
End Sub

Private Sub txtNS_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtNS_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub


