VERSION 5.00
Begin VB.Form frmTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
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
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "LOG OUT"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame fraInfo 
      Caption         =   "User Information"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5895
      Begin VB.Image ImgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NAME"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   960
         Width           =   2220
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblEmpNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USER ID"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   2220
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00808080&
         Caption         =   "USER ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "LOG TIME:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LOG TIME"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   1200
         Width           =   2220
      End
   End
   Begin VB.Frame fraTrans 
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "CHANGE PASSWORD"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PENALTY PAYMENT"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "LIST OF INFORMATION"
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
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "RETURN"
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdBorrow 
         Caption         =   "BORROW"
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   1770
         Left            =   2400
         Picture         =   "frmTransaction.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   120
      Picture         =   "frmTransaction.frx":D0FE
      Top             =   120
      Width           =   2970
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetUserInfo()
    
End Function

Private Sub lblGender_Click()

End Sub

Private Sub Command7_Click()
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
