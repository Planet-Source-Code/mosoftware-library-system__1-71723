VERSION 5.00
Begin VB.Form frmBLog 
   BorderStyle     =   0  'None
   Caption         =   "Borrowers Log"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   3255
      Begin VB.CommandButton cmdFind 
         Default         =   -1  'True
         Height          =   360
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Image imgStud 
         Height          =   795
         Left            =   120
         Picture         =   "frmBLog.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   3030
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Borrower ID."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   705
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmBLog.frx":55B32
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Identification  no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1860
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame fraRegInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   2535
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
      Begin VB.TextBox txtBtype 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtGender 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of Borrowers and their other Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   4230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowers Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   1875
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmBLog.frx":5BCAC
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         Caption         =   "B. TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   10
         Top             =   2055
         Width           =   1080
      End
      Begin VB.Image ImgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1410
         Left            =   240
         Picture         =   "frmBLog.frx":7265E
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808080&
         Caption         =   "ID NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   8
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   7
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1800
         TabIndex        =   6
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   0
      Picture         =   "frmBLog.frx":799EF
      Top             =   0
      Width           =   10590
   End
End
Attribute VB_Name = "frmBLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bIns As Boolean

Private Sub cmdFind_Click()
    GetBorrowerProfile
End Sub

Private Sub Form_Load()
    cmdFind.Picture = frmMain.i16x16.ListImages(11).ExtractIcon
End Sub

Public Function GetBorrowerProfile()
    On Error GoTo errHandler
    Dim sSQL As String, sValues As String
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    sSQL = "SELECT tbl_borrowers.B_id, tbl_borrower_type.b_type, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.add, tbl_borrowers.gender, tbl_borrowers.bday " & _
        "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
        "WHERE (((tbl_borrowers.B_id) Like '" & txtSearch.Text & "')) " & _
        "GROUP BY tbl_borrowers.B_id, tbl_borrower_type.b_type, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.add, tbl_borrowers.gender, tbl_borrowers.bday;"
    'MsgBox sSQL
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            bIns = True
            txtId.Text = adoRes.Fields("B_id")
            txtName.Text = adoRes.Fields("fn") & " " & Left(adoRes.Fields("mn"), 1) & ". " & adoRes.Fields("ln")
            txtGender.Text = adoRes.Fields("gender")
            txtBtype.Text = adoRes.Fields("b_type")
            ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & txtId.Text & ".jpg")
        Else
            bIns = False
            txtSearch.Text = ""
            txtId.Text = ""
            txtName.Text = ""
            txtGender.Text = ""
            txtBtype.Text = ""
            ImgPic.Picture = LoadPicture("")
            MsgBox "Invalid Borrower's ID.", vbExclamation, "Invalid ID"
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
errHandler:
    If err.Number = 53 Then ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & "nopic" & ".jpg")
    
    If bIns = True Then INSERTLOG
End Function

Public Function INSERTLOG()
    Dim sValues As String
    sValues = txtId.Text & "," & Time & "," & Date
    INSERT_DATA "tbl_borrower_log", "B_id,logtime,logdate", sValues, ",", False
End Function
