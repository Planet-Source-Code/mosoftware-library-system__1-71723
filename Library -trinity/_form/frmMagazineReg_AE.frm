VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMagazineReg_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdRmv 
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3400
         Width           =   300
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   3120
         Width           =   300
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
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1440
         Width           =   2145
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
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1800
         Width           =   2145
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
         Height          =   765
         Index           =   2
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2280
         Width           =   2985
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   4575
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1560
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   3120
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2752
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "AUTHOR"
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
         TabIndex        =   16
         Top             =   3150
         Width           =   1290
      End
      Begin VB.Image imgExc 
         Height          =   480
         Index           =   0
         Left            =   4080
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   240
         Picture         =   "frmMagazineReg_AE.frx":0000
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "REGISTERED ID"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You have to fill the following Information then click Save Button."
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
         Left            =   1080
         TabIndex        =   13
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodic Article Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   3390
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "ARTICLE"
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
         TabIndex        =   11
         Top             =   2325
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "ISSN"
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
         TabIndex        =   10
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3480
         Top             =   1800
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save"
      Top             =   4920
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4935
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   4935
      Width           =   315
   End
End
Attribute VB_Name = "frmMagazineReg_AE"
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_reg_magazines.rm_id " & _
            "From tbl_reg_magazines " & _
            "WHERE (((tbl_reg_magazines.rm_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_magazines.rm_id;"
    'txtInput(1).Text = Year(Date) & Month(Date) & Day(Date) & Format(Time, "hhmmss")
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_reg_magazines.rm_id " & _
            "From tbl_reg_magazines " & _
            "WHERE (((tbl_reg_magazines.rm_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_magazines.rm_id;"
    txtInput(1).Text = GenerateID
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
    Dim mRow As ListItem
    
    If Index = 0 Then
        gSQL = "SELECT tbl_authors.auid, tbl_authors.author, " & _
            "tbl_authors.yrborn FROM tbl_authors;"
        gIcon = 3
        gLblHead = "Book Author(s)"
        gLblDef = "Choose Book Author(s) then click Select button."
        gXY = "8010,4785"
        gTitle = "Select Author"
        gColumns = "Author ID,Author Name,Year Born"
        gColWidth = "1154,2954,1484"
        gFields = "auid,author,yrborn"
        gLoop = 2
        gLvIcon = 4
        sNoRec = "No Current Author(s) Info ."
    End If
    
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    
    If Index = 0 Then
        sSQL = "SELECT tbl_authors.author From tbl_authors WHERE (((tbl_authors.auid) Like '" & SelectedItem & "')) GROUP BY tbl_authors.author;"
        If isOnList(FindField(sSQL, "author")) = False Or lvList(0).ListItems.Count = 0 Then
            Set mRow = lvList(0).ListItems.Add(, , , , 4)
            mRow.SubItems(1) = FindField(sSQL, "author")
        Else
            MsgBox "You already add this Author.", vbExclamation, "Record Exist"
        End If
    End If
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_reg_magazines.DateAdded, tbl_reg_magazines.AddedByFK, tbl_reg_magazines.DateModified, tbl_reg_magazines.LastUserFK " & _
            "From tbl_reg_magazines " & _
            "WHERE (((tbl_reg_magazines.issn) Like '" & fCur.lvList(1).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_reg_magazines.DateAdded, tbl_reg_magazines.AddedByFK, tbl_reg_magazines.DateModified, tbl_reg_magazines.LastUserFK;"
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

Private Sub cmdRmv_Click()
    If lvList(0).ListItems.Count > 0 Then
        lvList(0).ListItems.Remove lvList(0).SelectedItem.Index
    End If
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
    sChkSql = "SELECT tbl_reg_magazines.rm_id " & _
            "From tbl_reg_magazines " & _
            "WHERE (((tbl_reg_magazines.rm_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_magazines.rm_id;"
    If isFill(1) = True Then
            sValues = txtInput(1).Text & "^" & txtInput(0).Text & "^" & txtInput(2).Text & "^" & GetAuthors _
                  & "^" & Date & "^" & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                INSERT_DATA "tbl_reg_magazines", "rm_id,issn,article,authors,DateAdded,AddedByFK", sValues, "^", True
                frmMagazine.btnNew_Load 1
                'LvSearchItem frmMagazine.lvList(1), txtInput(1).Text
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Registered ID already exist. Please change it!", vbExclamation
                txtInput(1).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_reg_magazines.rm_id " & _
            "From tbl_reg_magazines " & _
            "WHERE (((tbl_reg_magazines.rm_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_magazines.rm_id;"
    If isFill(1) = True Then
        sWhere = "rm_id like '" & frmMagazine.lvList(1).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(2).Text & "^" & GetAuthors _
                    & "^" & Date & "^" & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(1).Text = frmMagazine.lvList(1).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                'MsgBox GetAuthors
                UPDATE_DATA "tbl_reg_magazines", "article^Authors^DateModified^LastUserFK", sValues, sWhere, "^", True
                With frmMagazine.lvList(1)
                    .SelectedItem.SubItems(2) = txtInput(2).Text
                    .SelectedItem.SubItems(3) = GetAuthors
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Reg. Magazine ID already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    On Error Resume Next
    ifraIndex = 0
    Set lvList(0).SmallIcons = frmMain.iLv
    Set lvList(0).Icons = frmMain.iLv
    lvList(0).ColumnHeaders.Add , , , 300
    lvList(0).ColumnHeaders.Add , , "Author(s)", lvList(0).Width - 300
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Text = frmMagazine.lvList(0).SelectedItem.SubItems(1)
        txtInput(1).Text = GenerateID
        cmdCheck.Visible = True
        'cmdGen
    Else
        With frmMagazine
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(1).SelectedItem.SubItems(1)
            txtInput(2).Text = .lvList(1).SelectedItem.SubItems(2)
            InsAuthors .lvList(1).SelectedItem.SubItems(3)
        End With
        cmdCheck.Visible = False
        Me.Caption = "Edit Existing"
    End If
    SetButtonPicture
End Sub

Private Function isOnList(sAuthor As String) As Boolean
    Dim i As Integer
    isOnList = False
    For i = 1 To lvList(0).ListItems.Count
        If sAuthor = lvList(0).ListItems(i).Text Then
            isOnList = True
            Exit For
        End If
    Next
End Function

Public Function InsAuthors(sAuthors As String)
    Dim vAut As Variant, i As Integer
    Dim mRow As ListItem
    vAut = Split(sAuthors, "-")
    lvList(0).ListItems.Clear
    For i = 0 To CountSplitItem(sAuthors, "-")
        Set mRow = lvList(0).ListItems.Add(, , , , 4)
        mRow.SubItems(1) = vAut(i)
    Next
End Function

Public Function GetAuthors() As String
    Dim i As Integer
    GetAuthors = ""
    For i = 1 To lvList(0).ListItems.Count
        GetAuthors = GetAuthors & lvList(0).ListItems(i).SubItems(1)
        If Len(GetAuthors) > 0 And Not i = lvList(0).ListItems.Count Then
            GetAuthors = GetAuthors & "-" & lvList(0).ListItems(i).SubItems(1)
        End If
    Next
End Function

Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdCheck.Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen.Picture = .i16x16.ListImages(17).ExtractIcon

        cmdGet(0).Picture = .i16x16.ListImages(2).ExtractIcon
        cmdRmv.Picture = .i16x16.ListImages(4).ExtractIcon
    End With
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
        For i = 1 To 2
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Reg. ID"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Article"
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


