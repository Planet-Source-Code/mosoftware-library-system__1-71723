VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBorrower_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   5535
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   4815
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   5160
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   17498112
         CurrentDate     =   39492
         MinDate         =   -109205.020833333
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   7
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4800
         Width           =   2985
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   6
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   4440
         Width           =   2985
      End
      Begin VB.ComboBox cboGen 
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
         ItemData        =   "frmBorrower_AE.frx":0000
         Left            =   1320
         List            =   "frmBorrower_AE.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   5
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3720
         Width           =   2985
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   1320
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3360
         Width           =   2985
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00C0FFFF&
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
         Index           =   0
         Left            =   1320
         MaxLength       =   14
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1440
         Width           =   1905
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   2985
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   3360
         Width           =   300
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2280
         Width           =   2985
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2640
         Width           =   2985
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "LAST NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   2670
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "BIRTHDATE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   5190
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "CELL NO."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   4845
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "TEL. NO."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   4485
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   3765
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "B.TYPE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   3390
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "FIRST NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Image imgExc 
         Height          =   360
         Index           =   0
         Left            =   4275
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   480
         Picture         =   "frmBorrower_AE.frx":001C
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "BORROWER ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill the following information then click Save."
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
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   20
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "MIDDLE NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   2310
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   3050
         Width           =   1305
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3240
         Top             =   1440
         Width           =   360
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save"
      Top             =   5640
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5655
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancel"
      Top             =   5655
      Width           =   315
   End
End
Attribute VB_Name = "frmBorrower_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bStat As Boolean
Public fCur As Form

Dim sSQL As String
Dim ifraIndex As Integer
Dim SelectedItem As String
'Dim bSFocus As Boolean
Dim sMsgBox As String
Dim vType As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_borrowers.B_id " & _
            "From tbl_borrowers " & _
            "WHERE (((tbl_borrowers.B_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrowers.B_id;"
    'txtInput(1).Text = Year(Date) & Month(Date) & Day(Date) & Format(Time, "hhmmss")
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_borrowers.B_id " & _
            "From tbl_borrowers " & _
            "WHERE (((tbl_borrowers.B_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrowers.B_id;"
    txtInput(0).Text = GenerateID
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Dim newFormSelect As New frmSelect
    Dim gSQL As String, gLblHead As String, gLblDef As String
    Dim gXY As String, gTitle As String, gColumns As String
    Dim gColWidth As String, gFields As String, gLoop As Integer
    Dim gIcon As Integer, gLvIcon As Integer, sNoRec As String
    
    If Index = 0 Then
        gSQL = "SELECT tbl_borrower_type.bt_id, tbl_borrower_type.b_type, tbl_borrower_type.penaltystat, tbl_borrower_type.maxdaysborrow, tbl_borrower_type.maxnoborrow " & _
            "FROM tbl_borrower_type;"
        gIcon = 11
        gLblHead = "Borrower Type"
        gLblDef = "Choose Borrower Type then click Select button."
        gXY = "10020,4785"
        gTitle = "Select Borrower Type"
        gColumns = "B.Type ID,Type Name,Penalty(on/off),Max Days / Book,Max No. Borrowed"
        gColWidth = "1700,2500,1500,1550,1500"
        gFields = "bt_id,b_type,penaltystat,maxdaysborrow,maxnoborrow"
        gLoop = CountSplitItem(gColumns, ",")
        gLvIcon = 13
        sNoRec = "No Current Borrower Type Record."
    End If
    
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    
    If Index = 0 And Len(SelectedItem) > 0 Then
        sSQL = "SELECT tbl_borrower_type.b_type " & _
            "From tbl_borrower_type " & _
            "WHERE (((tbl_borrower_type.bt_id) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_borrower_type.b_type;"
        txtInput(4).Text = SelectedItem & "-" & FindField(sSQL, "b_type")
    End If
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_borrowers.DateAdded, tbl_borrowers.AddedByFK, tbl_borrowers.DateModified, tbl_borrowers.LastUserFK " & _
            "From tbl_borrowers " & _
            "WHERE (((tbl_borrowers.B_id) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_borrowers.DateAdded, tbl_borrowers.AddedByFK, tbl_borrowers.DateModified, tbl_borrowers.LastUserFK;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        tDate1 = adoRes.Fields("DateAdded")
        tUser1 = adoRes.Fields("AddedByFK")
        tDate2 = adoRes.Fields("DateModified")
        tUser2 = adoRes.Fields("LastUserFK")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    tUser1 = getValueAt(tUser1)
    tUser2 = getValueAt(tUser2)
    strMess = "Date Added: " & tDate1 & vbCrLf & _
            "Added By: " & tUser1 & vbCrLf & _
            "" & vbCrLf & _
            "Last Modified: " & tDate2 & vbCrLf & _
            "Modified By: " & tUser2
    cmdSave.ToolTipText = strMess
    Call MsgBox(strMess, vbInformation, "Modification History")
End Sub

Private Sub cmdSave_Click()
    If bStat = False Then
        ModifyData
    Else
        InsertData
    End If
End Sub

Public Function InsertData()
    Dim sValues As String
    Dim sChkSql As String
    Dim i As Integer
    sChkSql = "SELECT tbl_borrowers.B_id " & _
            "From tbl_borrowers " & _
            "WHERE (((tbl_borrowers.B_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrowers.B_id;"
    If isFill(1) = True Then
            vType = Split(txtInput(4).Text, "-")
            sValues = txtInput(0).Text & "," & txtInput(1).Text & "," & txtInput(2).Text & "," & txtInput(3).Text & "," & cboGen.Text & "," & vType(0) & "," & txtInput(5).Text & "," & txtInput(6).Text & "," & txtInput(7).Text & "," & txtDate.Value _
                  & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                INSERT_DATA "tbl_borrowers", "B_id,fn,mn,ln,gender,bt_id,[add],tel,cel,bday,DateAdded,AddedByFK", sValues, ",", True
                frmBorrower.btnNew_Load 0
                LvSearchItem frmBorrower.lvList(0), txtInput(0).Text
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Borrower ID already exist. Please change it!", vbExclamation
                txtInput(1).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    Dim cName As String
    sChkSql = "SELECT tbl_borrowers.B_id " & _
            "From tbl_borrowers " & _
            "WHERE (((tbl_borrowers.B_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrowers.B_id;"
    If isFill(1) = True Then
        vType = Split(txtInput(4).Text, "-")
        sWhere = "B_id like '" & frmBorrower.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(1).Text & "," & txtInput(2).Text & "," & txtInput(3).Text & "," & cboGen.Text & "," & vType(0) & "," & txtInput(5).Text & "," & txtInput(6).Text & "," & txtInput(7).Text & "," & txtDate.Value _
                    & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmBorrower.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_borrowers", "fn,mn,ln,gender,bt_id,[add],tel,cel,bday,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmBorrower.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(3) = txtInput(2).Text
                    .SelectedItem.SubItems(4) = txtInput(3).Text
                    .SelectedItem.SubItems(5) = cboGen.Text
                    .SelectedItem.SubItems(6) = vType(1)
                    .SelectedItem.SubItems(7) = txtInput(5).Text
                    .SelectedItem.SubItems(8) = txtInput(6).Text
                    .SelectedItem.SubItems(9) = txtInput(7).Text
                    .SelectedItem.SubItems(10) = txtDate.Value
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Borrower ID already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub DTPicker1_Click()
    MsgBox DTPicker1.Value
End Sub

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Text = frmBorrower.lvList(0).SelectedItem.SubItems(1)
        txtInput(0).Text = GenerateID
        cmdCheck.Visible = True
        txtDate.Value = (Date - (360 * 15))
        'cmdGen
    Else
        With frmBorrower
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(2)
            txtInput(2).Text = .lvList(0).SelectedItem.SubItems(3)
            txtInput(3).Text = .lvList(0).SelectedItem.SubItems(4)
            cboGen.Text = .lvList(0).SelectedItem.SubItems(5)
            sSQL = "SELECT tbl_borrower_type.bt_id " & _
                "From tbl_borrower_type " & _
                "WHERE (((tbl_borrower_type.b_type) Like '" & .lvList(0).SelectedItem.SubItems(6) & "')) " & _
                "GROUP BY tbl_borrower_type.bt_id;"
            txtInput(4).Text = FindField(sSQL, "bt_id") & "-" & .lvList(0).SelectedItem.SubItems(6)
            txtInput(5).Text = .lvList(0).SelectedItem.SubItems(7)
            txtInput(6).Text = .lvList(0).SelectedItem.SubItems(8)
            txtInput(7).Text = .lvList(0).SelectedItem.SubItems(9)
            txtDate.Value = .lvList(0).SelectedItem.SubItems(10)
        End With
        cmdCheck.Visible = False
        cmdGen.Visible = False
        Me.Caption = "Edit Existing"
    End If
    SetButtonPicture
End Sub



Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdCheck.Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen.Picture = .i16x16.ListImages(17).ExtractIcon
        cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
    End With
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    If Index = 4 Then
        txtInput(5).SetFocus
    End If
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 Then
        MsgBox "Unabled to Input B.Type. You have to click the button in the upper right corner of the textbox to Insert a Barcode." _
        , vbExclamation, "Unabled Input"
    ElseIf Index = 0 Then
        KeyAscii = isNumber(KeyAscii)
    ElseIf Index >= 6 And Index <= 7 Then
        KeyAscii = isNumber(KeyAscii)
    ElseIf Index >= 1 And Index <= 3 Then
        KeyAscii = isChar(KeyAscii)
    End If
End Sub

Private Function isFill(iStep As Integer) As Boolean
    Dim sBeginMsg As String
    Dim sEndMsg As String
    Dim iNotFill As Integer
    Dim i As Integer
    Dim iStart As Integer
    Dim iSplitItem As Integer
    Dim sSplitedItem As Variant
    iNotFill = 0
    sMsgBox = ""
    
    sBeginMsg = "You forgot to fill the following "
    sEndMsg = "."
    If iStep = 1 Then
        For i = 0 To 7
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "Borrower ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "First Name"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Middle Name"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Last Name"
        End If
        If isNull(txtInput(4).Text) = True Then
            FillMsgBox "B.Type"
        End If
        If isNull(txtInput(5).Text) = True Then
            FillMsgBox "Address"
        End If
        If isNull(txtInput(6).Text) = True Then
            FillMsgBox "Tel. No."
        End If
        If isNull(txtInput(7).Text) = True Then
            FillMsgBox "Cell No."
        End If
        
        If isNull(cboGen.Text) = True Then
            FillMsgBox "Gender"
        End If
        
        iSplitItem = CountSplitItem(sMsgBox, ",")
        sSplitedItem = Split(sMsgBox, ",")
        If Len(sMsgBox) = 0 Then isFill = True Else isFill = False
        If isFill = False Then
            If iSplitItem > 0 Then
                sMsgBox = ""
                For i = 0 To iSplitItem
                    If i = iSplitItem Then sMsgBox = sMsgBox & " and" & sSplitedItem(i): Exit For
                    If i = 0 Then sMsgBox = sSplitedItem(i) Else sMsgBox = sMsgBox & "," & sSplitedItem(i)
                Next
            End If
            sMsgBox = sBeginMsg & sMsgBox & sEndMsg
            MsgBox sMsgBox, vbExclamation, "isFill"
            txtInput(iStart).SetFocus
        End If
    End If
End Function

Private Function FillMsgBox(sMsg As String)
    If Len(sMsgBox) = 0 Then
        sMsgBox = sMsg
    Else
        sMsgBox = sMsgBox & ", " & sMsg
    End If
End Function

