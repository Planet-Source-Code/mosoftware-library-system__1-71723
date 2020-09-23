VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5520
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame fraQuery 
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
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6855
      Begin VB.Frame Frame2 
         Caption         =   " Condition "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   6615
         Begin VB.TextBox txtFilter 
            Height          =   285
            Index           =   1
            Left            =   3120
            TabIndex        =   8
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtFilter 
            Height          =   285
            Index           =   0
            Left            =   3120
            TabIndex        =   2
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox cmbOperation 
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
            Index           =   0
            ItemData        =   "frmSearch.frx":0000
            Left            =   240
            List            =   "frmSearch.frx":001C
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   2470
         End
         Begin VB.ComboBox cmbOperation 
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
            Index           =   1
            ItemData        =   "frmSearch.frx":00A4
            Left            =   240
            List            =   "frmSearch.frx":00C0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   2470
         End
         Begin VB.OptionButton Option1 
            Caption         =   "And"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   720
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Or"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   285
            Index           =   0
            Left            =   3120
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   71958531
            CurrentDate     =   38207
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   285
            Index           =   1
            Left            =   5040
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   71958531
            CurrentDate     =   38207
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   285
            Index           =   2
            Left            =   3120
            TabIndex        =   9
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   71958531
            CurrentDate     =   38207
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   285
            Index           =   3
            Left            =   5040
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "MMM-dd-yyyy"
            Format          =   71958531
            CurrentDate     =   38207
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   360
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "And"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   17
            Top             =   390
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "And"
            Height          =   255
            Left            =   4560
            TabIndex        =   16
            Top             =   1110
            Width           =   375
         End
      End
      Begin VB.ComboBox cmbFields 
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
         ItemData        =   "frmSearch.frx":0148
         Left            =   1800
         List            =   "frmSearch.frx":014A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   4875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Fields then fill the following the click Search."
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
         Index           =   2
         Left            =   960
         TabIndex        =   19
         Top             =   480
         Width           =   3630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Records Where?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Query Analyzer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   1410
      End
      Begin VB.Image Image3 
         Height          =   600
         Index           =   1
         Left            =   120
         Picture         =   "frmSearch.frx":014C
         Stretch         =   -1  'True
         Top             =   180
         Width           =   600
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

Option Explicit


Public srcColumnHeaders As String 'Source column headers
Public srcNoOfCol As Long
Public srcForm As Form 'Source form
Public strFilter As String

Private Sub cmbOperation_Click(Index As Integer)
    If Index = 0 Then
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(0).Visible = True
            dtpDate(1).Visible = True
            txtFilter(0).Visible = False
        Else
            txtFilter(0).Visible = True
            dtpDate(0).Visible = False
            dtpDate(1).Visible = False
        End If
    Else
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(2).Visible = True
            dtpDate(3).Visible = True
            txtFilter(1).Visible = False
        Else
            txtFilter(1).Visible = True
            dtpDate(2).Visible = False
            dtpDate(3).Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Verify
    Dim vSearchFields As Variant
    vSearchFields = Split(srcForm.sSearchFields, ",")
    If cmbOperation(0).ListIndex <> 7 Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    
    On Error GoTo err
    'Initialize the fields
    strFilter = Replace(cmbFields.Text, "/", "") 'ex. City/Town for tblCustomer
    strFilter = Replace(cmbFields.Text, " ", "")
    strFilter = "[" & vSearchFields(cmbFields.ListIndex) & "]"
    'Initialize the operation used
    'First operation
    Select Case cmbOperation(0).ListIndex
        Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
        Case 1: strFilter = strFilter & " = '" & txtFilter(0).Text & "'"
        Case 2: strFilter = strFilter & " <> '" & txtFilter(0).Text & "'"
        Case 3: strFilter = strFilter & " > '" & txtFilter(0).Text & "'"
        Case 4: strFilter = strFilter & " >= '" & txtFilter(0).Text & "'"
        Case 5: strFilter = strFilter & " < '" & txtFilter(0).Text & "'"
        Case 6: strFilter = strFilter & " <= '" & txtFilter(0).Text & "'"
        Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(0).Value & "# AND #" & dtpDate(1).Value & "#"
    End Select
    If cmbOperation(1).Text <> "" Then
        '-Second operation
        If Option1.Value = True Then
            strFilter = strFilter & " AND "
        Else
            strFilter = strFilter & " OR "
        End If
        
        strFilter = strFilter & "[" & vSearchFields(cmbFields.ListIndex) & "]"
        Select Case cmbOperation(1).ListIndex
            Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(1).Text & "%'"
            Case 1: strFilter = strFilter & " = '" & txtFilter(1).Text & "'"
            Case 2: strFilter = strFilter & " <> '" & txtFilter(1).Text & "'"
            Case 3: strFilter = strFilter & " > '" & txtFilter(1).Text & "'"
            Case 4: strFilter = strFilter & " >= '" & txtFilter(1).Text & "'"
            Case 5: strFilter = strFilter & " < '" & txtFilter(1).Text & "'"
            Case 6: strFilter = strFilter & " <= '" & txtFilter(1).Text & "'"
            Case 7: strFilter = strFilter & " BETWEEN #" & dtpDate(2).Value & "# AND #" & dtpDate(3).Value & "#"
        End Select
    End If
    'MsgBox strFilter
    srcForm.Execute_SearchItem strFilter
    'InputBox "", , strFilter
    'Pass the condition to filtered records
    'srcForm.FilterRecord strFilter
    'Clear used variables
    'strFilter = vbNullString
    
    Unload Me
    Exit Sub
err:
        If err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf err.Number = 3001 Then
            Resume Next
        Else
            'prompt_err err, "frmFilter", "cmdOk_Click"
        End If
End Sub

Private Sub Form_Load()
    Dim vColHead As Variant
    'Initialize values
    dtpDate(0).Value = Date
    dtpDate(1).Value = Date
    dtpDate(2).Value = Date
    dtpDate(3).Value = Date
    'Set the images for the controls
    With frmMain
        Image1.Picture = .i16x16.ListImages(15).Picture
        Image2.Picture = .i16x16.ListImages(15).Picture
    End With
    
    Dim i As Integer
    srcNoOfCol = CountSplitItem(srcColumnHeaders, ",")
    vColHead = Split(srcColumnHeaders, ",")

    For i = 0 To srcNoOfCol
        If vColHead(i) <> "" Then cmbFields.AddItem vColHead(i)
    Next i
    i = 0
    
    cmbFields.ListIndex = 0
    cmbOperation(0).ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearch = Nothing
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    HLText txtFilter(Index)
End Sub
