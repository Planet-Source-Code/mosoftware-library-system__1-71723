VERSION 5.00
Begin VB.Form frmBType_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5140
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
         Left            =   1920
         MaxLength       =   10
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3360
         Width           =   2385
      End
      Begin VB.CheckBox chkPenalty 
         Caption         =   "You can (on/off) of Penalty for this type only."
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
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2550
         Width           =   2415
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
         Left            =   1920
         MaxLength       =   10
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2280
         Width           =   2385
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
         Left            =   1920
         Locked          =   -1  'True
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
         Left            =   1920
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
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
         Index           =   3
         Left            =   1920
         MaxLength       =   10
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3000
         Width           =   2385
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Index           =   0
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   0
         Left            =   4365
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   4900
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PENALTY FEE / DAY"
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
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PENALTY "
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
         Top             =   2640
         Width           =   1785
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "MAX NO. BORROWED"
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
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Image imgExc 
         Height          =   360
         Index           =   0
         Left            =   4560
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   480
         Picture         =   "frmBType_AE.frx":0000
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "TYPE NAME"
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
         TabIndex        =   17
         Top             =   1965
         Width           =   1905
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
         TabIndex        =   16
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowers Type"
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
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "MAX DAYS BORROWED"
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
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "B.TYPE ID"
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
         TabIndex        =   13
         Top             =   1485
         Width           =   1875
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3840
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
         Width           =   5140
      End
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save"
      Top             =   3960
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   3975
      Width           =   315
   End
End
Attribute VB_Name = "frmBType_AE"
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

Private Sub cmdCheck_Click(Index As Integer)
    Dim sChkSql As String
    If Index = 0 Then
        sChkSql = "SELECT tbl_borrower_type.bt_id " & _
                "From tbl_borrower_type " & _
                "WHERE (((tbl_borrower_type.bt_id) Like '" & txtInput(0).Text & "')) " & _
                "GROUP BY tbl_borrower_type.bt_id;"
    End If
    If isRecordExist(sChkSql) = True Then
        imgChk(Index).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(Index).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub


Private Sub cmdGen_Click(Index As Integer)
    Dim sChkSql As String
    sChkSql = "SELECT tbl_borrower_type.bt_id " & _
            "From tbl_borrower_type " & _
            "WHERE (((tbl_borrower_type.bt_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrower_type.bt_id;"
    txtInput(0).Text = GenerateID
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdMod_Click()
    'On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_borrower_type.DateAdded, tbl_borrower_type.AddedByFK, tbl_borrower_type.DateModified, tbl_borrower_type.LastUserFK " & _
            "From tbl_borrower_type " & _
            "WHERE (((tbl_borrower_type.bt_id) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_borrower_type.DateAdded, tbl_borrower_type.AddedByFK, tbl_borrower_type.DateModified, tbl_borrower_type.LastUserFK;"
    'MsgBox sSql
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
    sChkSql = "SELECT tbl_borrower_type.bt_id " & _
            "From tbl_borrower_type " & _
            "WHERE (((tbl_borrower_type.bt_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrower_type.bt_id;"
    If isFill(1) = True Then
            sValues = txtInput(0).Text & "," & txtInput(1).Text & "," & chkPenalty.Value & "," & txtInput(3).Text & "," & txtInput(2).Text & "," & txtInput(4).Text _
                 & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                sChkSql = "SELECT tbl_borrower_type.bt_id " & _
                        "From tbl_borrower_type " & _
                        "WHERE (((tbl_borrower_type.b_type) Like '" & txtInput(1).Text & "')) " & _
                        "GROUP BY tbl_borrower_type.bt_id;"
                If isRecordExist(sChkSql) = False Then
                    imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                    INSERT_DATA "tbl_borrower_type", "bt_id,b_type,penaltystat,maxdaysborrow,maxnoborrow,p_fee,DateAdded,AddedByFK", sValues, ",", True
                    frmBType.btnNew_Load 0
                    LvSearchItem frmBType.lvList(0), txtInput(0).Text
                    Unload Me
                Else
                    MsgBox "Type Name already exist. Please change it!", vbExclamation, "PointerException"
                    txtInput(2).SetFocus
                End If
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "B.Type ID already exist. Please change it!", vbExclamation, "PointerException"
                txtInput(1).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_borrower_type.bt_id " & _
            "From tbl_borrower_type " & _
            "WHERE (((tbl_borrower_type.bt_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_borrower_type.bt_id;"
    If isFill(1) = True Then
        If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmBType.lvList(0).SelectedItem.SubItems(1) Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
            sChkSql = "SELECT tbl_borrower_type.bt_id " & _
                    "From tbl_borrower_type " & _
                    "WHERE (((tbl_borrower_type.b_type) Like '" & txtInput(1).Text & "')) " & _
                    "GROUP BY tbl_borrower_type.bt_id;"
            If isRecordExist(sChkSql) = False Or txtInput(1).Text = frmBType.lvList(0).SelectedItem.SubItems(2) Then
                sWhere = "bt_id like '" & frmBType.lvList(0).SelectedItem.SubItems(1) & "'"
                sValues = txtInput(1).Text & "," & chkPenalty.Value & "," & txtInput(3).Text & "," & txtInput(2).Text & "," & txtInput(4).Text _
                      & "," & Date & "," & sUserId
                UPDATE_DATA "tbl_borrower_type", "b_type,penaltystat,maxdaysborrow,maxnoborrow,p_fee,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmBType.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(3) = chkPenalty.Value
                    .SelectedItem.SubItems(4) = txtInput(3).Text
                    .SelectedItem.SubItems(5) = txtInput(2).Text
                    .SelectedItem.SubItems(6) = txtInput(4).Text
                End With
                Unload Me
            Else
                MsgBox "Type Name already exist. Please change it!", vbExclamation, "PointerException"
                txtInput(2).SetFocus
            End If
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
            MsgBox "BType ID already exist. Please change it!", vbExclamation
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
        cmdGen(0).Visible = True
    Else
        With frmBType
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(2)
            chkPenalty.Value = .lvList(0).SelectedItem.SubItems(3)
            txtInput(3).Text = .lvList(0).SelectedItem.SubItems(4)
            txtInput(2).Text = .lvList(0).SelectedItem.SubItems(5)
            txtInput(4).Text = .lvList(0).SelectedItem.SubItems(6)
        End With
        cmdCheck(0).Visible = False
        cmdGen(0).Visible = False
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
        cmdCheck(0).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen(0).Picture = .i16x16.ListImages(17).ExtractIcon
        'cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
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
            FillMsgBox "B.Type ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Type Name"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Max No. Borrowed"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Max Days Borrowed"
        End If
        If isNull(txtInput(4).Text) = True Then
            FillMsgBox "Penalty Fee/Day"
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
            'MsgBox sMsgBox, vbExclamation, "isFill"
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
    If Index = 2 Or Index = 3 Then
        KeyAscii = isNumber(KeyAscii)
    ElseIf Index = 4 Then
        KeyAscii = isCurrency(KeyAscii, txtInput(4).Text)
    End If
End Sub
