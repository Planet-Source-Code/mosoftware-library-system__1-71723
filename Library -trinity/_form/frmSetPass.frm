VERSION 5.00
Begin VB.Form frmSetPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set User Password"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   975
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame fraList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3615
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtConfirm 
         Alignment       =   2  'Center
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
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtNew 
         Alignment       =   2  'Center
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
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
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtUsr 
         Alignment       =   2  'Center
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
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1605
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1245
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   780
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set User Password"
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
      TabIndex        =   8
      Top             =   120
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSetPass.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can Set your User password."
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
      Top             =   360
      Width           =   2400
   End
End
Attribute VB_Name = "frmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Trim(txtUsr.Text) = "" Then
        Call MsgBox("You forgot to fill Username.", vbExclamation, "Null")
        txtUsr.SetFocus
    ElseIf Trim(txtOld.Text) = "" Then
        Call MsgBox("You forgot to fill Old Password.", vbExclamation, "Null")
        txtOld.SetFocus
    ElseIf Trim(txtNew.Text) = "" Then
        Call MsgBox("You forgot to fill New Password.", vbExclamation, "Null")
        txtNew.SetFocus
    ElseIf Trim(txtConfirm.Text) = "" Then
        Call MsgBox("You forgot to fill Confirm Password.", vbExclamation, "Null")
        txtConfirm.SetFocus
    Else
        If CHECK_USERNAME = True Then
            If CHECK_OLD_PASSWORD = True Then
                If txtNew.Text = txtConfirm.Text Then
                    Call UPDATE_NEW_PASS
                    Call MsgBox("Update new Password successfull.", vbInformation, "Update")
                    Unload Me
                Else
                    Call MsgBox("Confirm password not match.", vbInformation, "Confirm")
                    txtNew.Text = ""
                    txtConfirm.Text = ""
                    txtNew.SetFocus
                End If
                
            Else
                Call MsgBox("Username and Password not match.", vbExclamation, "Match")
                txtUsr.Text = ""
                txtOld.Text = ""
                txtNew.Text = ""
                txtConfirm.Text = ""
                txtUsr.SetFocus
            End If
        Else
            Call MsgBox("Username not match.", vbExclamation, "Match")
            txtUsr.Text = ""
            txtOld.Text = ""
            txtNew.Text = ""
            txtConfirm.Text = ""
            txtUsr.SetFocus
        End If
    End If
End Sub

Public Function CHECK_USERNAME() As Boolean
    Dim strUser As String
    Dim strSQL As String
    'sUserId
    strSQL = "SELECT tbl_users.usrnme " & _
        "From tbl_users " & _
        "WHERE (((tbl_users.uid) Like '" & sUserId & "')) " & _
        "GROUP BY tbl_users.usrnme;"
    strUser = FindFieldValue(strSQL, "usrnme")
    If strUser = txtUsr.Text Then
        CHECK_USERNAME = True
    Else
        CHECK_USERNAME = False
    End If
End Function

Private Function CHECK_OLD_PASSWORD() As Boolean
    Dim strPass As String
    Dim strSQL As String
    strSQL = "SELECT tbl_users.uid " & _
            "From tbl_users " & _
            "WHERE (((tbl_users.usrnme) Like '" & txtUsr.Text & "') AND ((tbl_users.pass) Like '" & ENCRYPT(txtOld.Text) & "')) " & _
            "GROUP BY tbl_users.uid;"
    If isRecordExist(strSQL) = True Then
        CHECK_OLD_PASSWORD = True
    Else
        CHECK_OLD_PASSWORD = False
    End If
End Function

Private Sub UPDATE_NEW_PASS()
    Dim strWhere As String
    strWhere = "WHERE tbl_users.usrnme Like '" & txtUsr.Text & "' AND tbl_users.pass Like '" & ENCRYPT(txtOld.Text) & "'"
    UPDATE_DATA2 "tbl_users", "pass", strWhere, ENCRYPT(txtNew.Text)
End Sub

