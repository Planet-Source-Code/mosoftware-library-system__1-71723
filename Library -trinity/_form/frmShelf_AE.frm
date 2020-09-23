VERSION 5.00
Begin VB.Form frmShelf_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4815
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
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3120
         Width           =   1905
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
         Top             =   1560
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
         Top             =   2040
         Width           =   3100
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   2400
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
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2400
         Width           =   3100
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
         TabIndex        =   6
         Top             =   2760
         Width           =   1905
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   4620
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "MAX QUANTITY"
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
         TabIndex        =   19
         Top             =   3165
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "SECTION"
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
         TabIndex        =   18
         Top             =   2445
         Width           =   1215
      End
      Begin VB.Image imgExc 
         Height          =   600
         Index           =   0
         Left            =   4275
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   480
         Picture         =   "frmShelf_AE.frx":0000
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "SHELF NAME"
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
         TabIndex        =   17
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill the following information the click save."
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
         TabIndex        =   16
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHELF"
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
         TabIndex        =   15
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "ACRONYM"
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
         TabIndex        =   14
         Top             =   2805
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "SHELF ID"
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
         TabIndex        =   13
         Top             =   1590
         Width           =   1305
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3240
         Top             =   1560
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
      TabIndex        =   9
      ToolTipText     =   "Save"
      Top             =   3720
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   3720
      Width           =   315
   End
End
Attribute VB_Name = "frmShelf_AE"
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
    sChkSql = "SELECT tbl_shelfs.sh_id " & _
            "From tbl_shelfs " & _
            "WHERE (((tbl_shelfs.sh_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_shelfs.sh_id;"
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_shelfs.sh_id " & _
            "From tbl_shelfs " & _
            "WHERE (((tbl_shelfs.sh_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_shelfs.sh_id;"
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
        gSQL = "SELECT tbl_sections.sc_id, tbl_sections.sction, tbl_sections.dscrpt " & _
            "FROM tbl_sections;"
        gIcon = 10
        gLblHead = "Shelf Sections"
        gLblDef = "You can choose Section of Shelf that you wanted then click Select button."
        gXY = "8000,4785"
        gTitle = "Select Section"
        gColumns = "Section ID,Section Name,Description"
        gColWidth = "1700,2500,2000"
        gFields = "sc_id,sction,dscrpt"
        gLoop = CountSplitItem(gColumns, ",")
        gLvIcon = 8
        sNoRec = "No Current Shelf(s) Info ."
    End If
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    If Index = 0 And Len(SelectedItem) > 0 Then
        sSQL = "SELECT tbl_sections.sction " & _
            "From tbl_sections " & _
            "WHERE (((tbl_sections.sc_id) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_sections.sction;"
        txtInput(2).Text = SelectedItem & "-" & FindField(sSQL, "sction")
    End If
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_shelfs.DateAdded, tbl_shelfs.AddedByFK, tbl_shelfs.DateModified, tbl_shelfs.LastUserFK " & _
            "From tbl_shelfs " & _
            "WHERE (((tbl_shelfs.sh_id) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_shelfs.DateAdded, tbl_shelfs.AddedByFK, tbl_shelfs.DateModified, tbl_shelfs.LastUserFK;"
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
    Dim sValues As String, vSectionid As Variant
    Dim sChkSql As String
    Dim i As Integer
    sChkSql = "SELECT tbl_shelfs.sh_id " & _
            "From tbl_shelfs " & _
            "WHERE (((tbl_shelfs.sh_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_shelfs.sh_id;"
    If isFill(1) = True Then
        vSectionid = Split(txtInput(2).Text, "-")
        sValues = txtInput(0).Text & "," & vSectionid(0) & "," & txtInput(1).Text & "," & txtInput(3).Text & "," & txtInput(4).Text _
                  & "," & Date & "," & sUserId
        If isRecordExist(sChkSql) = False Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
            INSERT_DATA "tbl_shelfs", "sh_id,sc_id,shelfname,acronym,maxqty,DateAdded,AddedByFK", sValues, ",", True
            frmShelf.btnNew_Load 0
            LvSearchItem frmShelf.lvList(0), txtInput(0).Text
            Unload Me
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
            MsgBox "Shelf ID already exist. Please change it!", vbExclamation
            txtInput(1).SetFocus
        End If
    End If
End Function

Public Function ModifyData()
    Dim sValues As String, vSectionid As Variant
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_shelfs.sh_id " & _
            "From tbl_shelfs " & _
            "WHERE (((tbl_shelfs.sh_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_shelfs.sh_id;"
    If isFill(1) = True Then
        vSectionid = Split(txtInput(2).Text, "-")
        sWhere = "sh_id like '" & frmShelf.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = vSectionid(0) & "," & txtInput(1).Text & "," & txtInput(3).Text & "," & txtInput(4).Text _
                  & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmShelf.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_shelfs", "sc_id,shelfname,acronym,maxqty,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmShelf.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(3) = vSectionid(1)
                    .SelectedItem.SubItems(4) = txtInput(3).Text
                    .SelectedItem.SubItems(5) = txtInput(4).Text
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Shelf already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Text = GenerateID
        cmdCheck.Visible = True
    Else
        With frmShelf
            txtInput(0).Text = frmShelf.lvList(0).SelectedItem.SubItems(1)
            sSQL = "SELECT tbl_sections.sc_id " & _
                "From tbl_sections " & _
                "WHERE (((tbl_sections.sction) Like '" & .lvList(0).SelectedItem.SubItems(3) & "')) " & _
                "GROUP BY tbl_sections.sc_id;"
            txtInput(2).Text = FindField(sSQL, "sc_id") & "-" & .lvList(0).SelectedItem.SubItems(3)
            txtInput(1).Text = frmShelf.lvList(0).SelectedItem.SubItems(2)
            txtInput(3).Text = frmShelf.lvList(0).SelectedItem.SubItems(4)
            txtInput(4).Text = frmShelf.lvList(0).SelectedItem.SubItems(5)
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

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        MsgBox "Unabled to Input Section. You have to click the button in the upper right corner of the textbox to Insert a Section." _
        , vbExclamation, "InputPointerException"
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
        For i = 0 To 4
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "Shelf ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Shelf Name"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Section"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Acronym"
        End If
        If isNull(txtInput(4).Text) = True Then
            FillMsgBox "Max Quantity"
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
