VERSION 5.00
Begin VB.Form frmLogin 
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   9960
      Top             =   2160
   End
   Begin VB.Frame fraLog 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   5760
      TabIndex        =   3
      Top             =   3720
      Width           =   5295
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3735
         Begin VB.ComboBox cboAcc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmLogin.frx":0000
            Left            =   1560
            List            =   "frmLogin.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtPass 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1560
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   1
            Text            =   "Admin"
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   0
            Text            =   "Administrator"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
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
            TabIndex        =   7
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
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
            TabIndex        =   6
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
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
            TabIndex        =   5
            Top             =   240
            Width           =   780
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmLogin.frx":0004
            Top             =   480
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4080
         ScaleHeight     =   855
         ScaleWidth      =   1215
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label lblTop 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your username and password then select what you want to access."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BB5900&
         Height          =   435
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   3660
      End
   End
   Begin VB.Image Image6 
      Height          =   1140
      Left            =   2400
      Picture         =   "frmLogin.frx":08CE
      Top             =   1560
      Width           =   10590
   End
   Begin VB.Image Image3 
      Height          =   1320
      Left            =   4440
      Picture         =   "frmLogin.frx":582E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Image Image4 
      Height          =   555
      Left            =   7440
      Picture         =   "frmLogin.frx":C878
      Top             =   6000
      Width           =   2475
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMsgBox As String
Dim i

Private Sub cmdCancel_Click()
    Unload frmMain
End Sub

Private Sub cmdOk_Click()
    Dim sLogSQL As String
    If txtUser.Text = "" And txtPass.Text = "" Then
        MsgBox "Please input valid Username and Password.", vbExclamation, "Login Error"
        txtUser.SetFocus
    ElseIf txtUser.Text = "" Then
        MsgBox "Please input valid Username.", vbExclamation, "Login Error"
        txtUser.SetFocus
    ElseIf txtPass.Text = "" Then
        MsgBox "Please input valid Password.", vbExclamation, "Login Error"
        txtPass.SetFocus
    ElseIf cboAcc.Text = "" Then
        MsgBox "You forgot to Select where you want to Access.", vbExclamation, "Login Error"
        cboAcc.SetFocus
    Else
        If cboAcc.ListIndex = 0 Then
            sLogSQL = "SELECT tbl_users.uid, tbl_users.usrnme, tbl_users.pass, tbl_users.admn " & _
                    "From tbl_users " & _
                    "WHERE (((tbl_users.usrnme) Like '" & txtUser.Text & "') AND ((tbl_users.pass) Like '" & ENCRYPT(txtPass.Text) & "')) " & _
                    "GROUP BY tbl_users.uid, tbl_users.usrnme, tbl_users.pass, tbl_users.admn;"
            iUSER = 1
            If UserPassMatch(sLogSQL, "admn") = True Then
                Unload Me
                frmMain.ToolMenuStatus True
            End If
        ElseIf cboAcc.ListIndex = 2 Then
            sLogSQL = "SELECT tbl_users.uid, tbl_users.usrnme, tbl_users.pass, tbl_users.opac " & _
                    "From tbl_users " & _
                    "WHERE (((tbl_users.usrnme) Like '" & txtUser.Text & "') AND ((tbl_users.pass) Like '" & ENCRYPT(txtPass.Text) & "')) " & _
                    "GROUP BY tbl_users.uid, tbl_users.usrnme, tbl_users.pass, tbl_users.opac;"
            iUSER = 2
            If UserPassMatch(sLogSQL, "opac") = True Then
                Unload Me
                LoadForm frmOPAC
            End If
        ElseIf cboAcc.ListIndex = 3 Then
            iUSER = 3
            Unload Me
            LoadForm frmBLog
        End If
    End If
End Sub

Public Function UserPassMatch(sSQL, sUserFields As String) As Boolean
    Dim sUser As String
    Dim sPass As String
    Dim iStat As Integer
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        sUser = adoRes.Fields("usrnme")
        sPass = adoRes.Fields("pass")
        iStat = adoRes.Fields(sUserFields)
        sUserId = adoRes.Fields("uid")
    adoRes.Close
    adoCon.Close
    If txtUser.Text = sUser And ENCRYPT(txtPass.Text) = sPass Then
        If iStat = 1 Then
            UserPassMatch = True
        Else
            UserPassMatch = False
            MsgBox "You are not Authorize to Enter this System.", vbCritical, "Unauthorize Personnel"
            txtUser.Text = ""
            txtPass.Text = ""
            txtUser.SetFocus
        End If
    Else
        UserPassMatch = False
        MsgBox "Invalid Username and Password.", vbExclamation, "Login"
        txtUser.Text = ""
        txtPass.Text = ""
        txtUser.SetFocus
    End If
    Set adoRes = Nothing
    Set adoCon = Nothing
End Function

Private Sub Form_Load()
    'Dim wXP As WindowsXPC
    Call addAccessItem 'Add Access Items
    cboAcc.ListIndex = 0
    i = 1
    'ActivateXPTheme wXP
End Sub

Private Sub Form_Resize()
'    fraLog.Move (Me.ScaleWidth / 2) - (5760 / 2.5), (Me.ScaleHeight / 2) - (3000 / 1.5)
End Sub

Public Function addAccessItem()
    cboAcc.Clear
    cboAcc.AddItem "Admin"
    cboAcc.AddItem "Transaction"
    cboAcc.AddItem "Opac"
    cboAcc.AddItem "Borrower's Log"
End Function

Private Sub Timer1_Timer()
        Select Case i
            Case 1: lblTop.Caption = lblTop.Caption & "P"
            Case 2: lblTop.Caption = lblTop.Caption & "l"
            Case 3: lblTop.Caption = lblTop.Caption & "e"
            Case 4: lblTop.Caption = lblTop.Caption & "a"
            Case 5: lblTop.Caption = lblTop.Caption & "s"
            Case 6: lblTop.Caption = lblTop.Caption & "e"
            Case 7: lblTop.Caption = lblTop.Caption & " "
            Case 8: lblTop.Caption = lblTop.Caption & "e"
            Case 9: lblTop.Caption = lblTop.Caption & "n"
            Case 10: lblTop.Caption = lblTop.Caption & "t"
            Case 11: lblTop.Caption = lblTop.Caption & "e"
            Case 12: lblTop.Caption = lblTop.Caption & "r"
            Case 13: lblTop.Caption = lblTop.Caption & " "
            Case 14: lblTop.Caption = lblTop.Caption & "v"
            Case 15: lblTop.Caption = lblTop.Caption & "a"
            Case 16: lblTop.Caption = lblTop.Caption & "l"
            Case 17: lblTop.Caption = lblTop.Caption & "i"
            Case 18: lblTop.Caption = lblTop.Caption & "d"
            Case 19: lblTop.Caption = lblTop.Caption & " "
            Case 20: lblTop.Caption = lblTop.Caption & "u"
            Case 21: lblTop.Caption = lblTop.Caption & "s"
            Case 22: lblTop.Caption = lblTop.Caption & "e"
            Case 23: lblTop.Caption = lblTop.Caption & "r"
            Case 24: lblTop.Caption = lblTop.Caption & "n"
            Case 25: lblTop.Caption = lblTop.Caption & "a"
            Case 26: lblTop.Caption = lblTop.Caption & "m"
            Case 27: lblTop.Caption = lblTop.Caption & "e"
            Case 28: lblTop.Caption = lblTop.Caption & " "
            Case 29: lblTop.Caption = lblTop.Caption & "a"
            Case 30: lblTop.Caption = lblTop.Caption & "n"
            Case 31: lblTop.Caption = lblTop.Caption & "d"
            Case 32: lblTop.Caption = lblTop.Caption & " "
            Case 33: lblTop.Caption = lblTop.Caption & "p"
            Case 34: lblTop.Caption = lblTop.Caption & "a"
            Case 35: lblTop.Caption = lblTop.Caption & "s"
            Case 36: lblTop.Caption = lblTop.Caption & "s"
            Case 37: lblTop.Caption = lblTop.Caption & "w"
            Case 38: lblTop.Caption = lblTop.Caption & "o"
            Case 39: lblTop.Caption = lblTop.Caption & "r"
            Case 40: lblTop.Caption = lblTop.Caption & "d"
            Case 41: lblTop.Caption = lblTop.Caption & "."
        End Select
    i = i + 1
    If i = 41 Then Timer1.Enabled = False
End Sub
