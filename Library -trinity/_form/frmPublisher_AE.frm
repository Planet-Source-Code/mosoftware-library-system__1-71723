VERSION 5.00
Begin VB.Form frmPublisher_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   5175
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4220
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   2520
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
         Index           =   6
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   4680
         Width           =   3105
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
         Index           =   5
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4320
         Width           =   3105
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
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3960
         Width           =   3105
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
         Height          =   1005
         Index           =   3
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2880
         Width           =   3105
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   1
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
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
         Left            =   1440
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1680
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
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2160
         Width           =   2385
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
         Left            =   1440
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2520
         Width           =   2745
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   0
         Left            =   3885
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   4575
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   2925
         Width           =   1320
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   1
         Left            =   3840
         Top             =   2160
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "FAX"
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
         TabIndex        =   23
         Top             =   4320
         Width           =   1320
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "WEBSITE"
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
         TabIndex        =   22
         Top             =   4680
         Width           =   1320
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "TELEPHONE"
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
         TabIndex        =   21
         Top             =   3960
         Width           =   1320
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "COUNTRY"
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
         TabIndex        =   20
         Top             =   2565
         Width           =   1320
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
         Left            =   600
         Picture         =   "frmPublisher_AE.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "COMPANY NAME"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You have to fill the following information then click Save Button to Finish Adding New Publisher."
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
         Height          =   795
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
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
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PUBLISHER ID"
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
         TabIndex        =   16
         Top             =   1710
         Width           =   1320
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3360
         Top             =   1680
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
      TabIndex        =   12
      ToolTipText     =   "Save"
      Top             =   5280
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancel"
      Top             =   5295
      Width           =   315
   End
End
Attribute VB_Name = "frmPublisher_AE"
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

Private Sub cmdCheck_Clicks()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim sChkSql As String
    Dim sISBN As String
    If Index = 0 Then
        sChkSql = "SELECT tbl_publishers.pubid " & _
                "From tbl_publishers " & _
                "WHERE (((tbl_publishers.pubid) Like '" & txtInput(0).Text & "')) " & _
                "GROUP BY tbl_publishers.pubid;"
        If isRecordExist(sChkSql) = True Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    ElseIf Index = 1 Then
        sChkSql = "SELECT tbl_publishers.pubid " & _
                "From tbl_publishers " & _
                "WHERE (((tbl_publishers.cmpny) Like '" & txtInput(1).Text & "')) " & _
                "GROUP BY tbl_publishers.pubid;"
        If isRecordExist(sChkSql) = True Then
            imgChk(1).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(1).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_publishers.pubid " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.pubid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_publishers.pubid;"
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
        gSQL = "SELECT tbl_country.country " & _
            "From tbl_country " & _
            "GROUP BY tbl_country.country " & _
            "ORDER BY tbl_country.country;"
        gIcon = 7
        gLblHead = "Countries"
        gLblDef = "Choose Book Publisher then click Select button."
        gXY = "6000,5000"
        gTitle = "Select Country"
        gColumns = "Country"
        gColWidth = "2500"
        gFields = "country"
        gLoop = CountSplitItem(gColumns, ",")
        gLvIcon = 12
        sNoRec = "No Current Country Record."
    End If
    
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    
    If Index = 0 And Len(SelectedItem) > 0 Then
        txtInput(2).Text = SelectedItem
    End If
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_publishers.DateAdded, tbl_publishers.AddedByFK, tbl_publishers.DateModified, tbl_publishers.LastUserFK " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.isbn) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_publishers.DateAdded, tbl_publishers.AddedByFK, tbl_publishers.DateModified, tbl_publishers.LastUserFK;"
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
    sChkSql = "SELECT tbl_publishers.pubid " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.pubid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_publishers.pubid;"
    If isFill(1) = True Then
        sValues = txtInput(0).Text & "," & txtInput(1).Text & "," & txtInput(2).Text & "," & txtInput(3).Text & _
         "," & txtInput(4).Text & "," & txtInput(5).Text & "," & txtInput(6).Text & "," & Date & "," & sUserId
        If isRecordExist(sChkSql) = False Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
            INSERT_DATA "tbl_publishers", "pubid,cmpny,cntry,[add],tel,fx,wbst,DateAdded,AddedByFK", sValues, ",", True
            frmPublisher.btnNew_Load 0
            LvSearchItem frmPublisher.lvList(0), txtInput(0).Text
            Unload Me
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
            MsgBox "Publisher ID already exist. Please change it!", vbExclamation
            txtInput(0).SetFocus
        End If
    End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_publishers.pubid " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.pubid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_publishers.pubid;"
    If isFill(1) = True Then
        sWhere = "pubid like '" & frmPublisher.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(1).Text & "," & txtInput(2).Text & "," & txtInput(3).Text & "," & txtInput(4).Text & "," & txtInput(5).Text & "," & txtInput(6).Text _
                    & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmPublisher.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_publishers", "cmpny,cntry,[add],tel,fx,wbst,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmPublisher.lvList(0)
                    '.selecteditem.SubItems(1) = txtInput(1).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(3) = txtInput(2).Text
                    .SelectedItem.SubItems(4) = txtInput(3).Text
                    .SelectedItem.SubItems(5) = txtInput(4).Text
                    .SelectedItem.SubItems(6) = txtInput(5).Text
                    .SelectedItem.SubItems(7) = txtInput(6).Text
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "ISBN already exist. Please change it!", vbExclamation
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
        cmdCheck(0).Visible = True
        cmdGen.Visible = True
    Else
        With frmPublisher
            txtInput(0).Text = frmPublisher.lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = frmPublisher.lvList(0).SelectedItem.SubItems(2)
            txtInput(2).Text = frmPublisher.lvList(0).SelectedItem.SubItems(3)
            txtInput(3).Text = frmPublisher.lvList(0).SelectedItem.SubItems(4)
            txtInput(4).Text = frmPublisher.lvList(0).SelectedItem.SubItems(5)
            txtInput(5).Text = frmPublisher.lvList(0).SelectedItem.SubItems(6)
            txtInput(6).Text = frmPublisher.lvList(0).SelectedItem.SubItems(7)
        End With
        cmdCheck(0).Visible = False
        cmdGen.Visible = False
        Me.Caption = "Edit Existing"
    End If
    SetButtonPicture
End Sub

Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        cmdCheck(1).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdCheck(0).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen.Picture = .i16x16.ListImages(17).ExtractIcon
    End With
End Sub

Private Function isFill(iStep As Integer) As Boolean
    Dim sBeginMsg As String, sEndMsg As String, iNotFill As Integer, i As Integer
    Dim iStart As Integer, iSplitItem As Integer, sSplitedItem As Variant
    iNotFill = 0
    sMsgBox = ""
    sBeginMsg = "You forgot to fill the following "
    sEndMsg = "."
    If iStep = 1 Then
        For i = 0 To 6
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "Publisher ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Company Name"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Country"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Address"
        End If
        If isNull(txtInput(4).Text) = True Then
            FillMsgBox "Telephone"
        End If
        If isNull(txtInput(5).Text) = True Then
            FillMsgBox "Fax"
        End If
        If isNull(txtInput(6).Text) = True Then
            FillMsgBox "WebSite"
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

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 Or Index = 5 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub
