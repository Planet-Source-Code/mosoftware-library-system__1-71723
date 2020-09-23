VERSION 5.00
Begin VB.Form frmBcodeInputer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode Inputer"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Enter"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   1560
      Width           =   3950
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   1560
      Width           =   600
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtBcode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Barcode"
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
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You can Barcode it or type the Barcode Manually."
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
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   600
         Index           =   0
         Left            =   120
         Picture         =   "frmBcodeInputer.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   600
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmBcodeInputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sBarcode As String

Private Sub cmdCancel_Click()
    sBarcode = ""
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If isNull(txtBcode.Text) = False Then
        sBarcode = txtBcode.Text
        Unload Me
    Else
        MsgBox "Please Input Barcode first then Press the Button Enter.", vbExclamation, "isNull"
        txtBcode.SetFocus
    End If
End Sub

Private Sub txtBcode_Change()
    If isNull(txtBcode.Text) = True Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub
