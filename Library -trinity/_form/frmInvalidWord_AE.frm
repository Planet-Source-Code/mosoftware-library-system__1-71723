VERSION 5.00
Begin VB.Form frmInvalidWord_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save"
      Top             =   2280
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   2295
      Width           =   315
   End
   Begin VB.Frame fraList 
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   5
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
         Index           =   0
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1560
         Width           =   2265
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   300
      End
      Begin VB.Image imgExc 
         Height          =   240
         Index           =   0
         Left            =   4275
         Top             =   960
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   705
         Index           =   0
         Left            =   480
         Picture         =   "frmInvalidWord_AE.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   720
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
         TabIndex        =   8
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Word"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Top             =   315
         Width           =   2025
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "INVALID WORD"
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
         TabIndex        =   6
         Top             =   1590
         Width           =   1305
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3600
         Top             =   1560
         Width           =   360
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   315
   End
End
Attribute VB_Name = "frmInvalidWord_AE"
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
    sChkSql = "SELECT tbl_invalid_words.InvalidWord " & _
            "From tbl_invalid_words  " & _
            "WHERE (((tbl_invalid_words.InvalidWord) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_invalid_words.InvalidWord;"
    'txtInput(1).Text = Year(Date) & Month(Date) & Day(Date) & Format(Time, "hhmmss")
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_invalid_words.sc_id " & _
            "From tbl_invalid_words " & _
            "WHERE (((tbl_invalid_words.sc_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_invalid_words.sc_id;"
    txtInput(0).Text = GenerateID
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
    txtInput(2).Text = BarcodeValue
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_invalid_words.DateAdded, tbl_invalid_words.AddedByFK, tbl_invalid_words.DateModified, tbl_invalid_words.LastUserFK " & _
            "From tbl_invalid_words " & _
            "WHERE (((tbl_invalid_words.InvalidWord) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_invalid_words.DateAdded, tbl_invalid_words.AddedByFK, tbl_invalid_words.DateModified, tbl_invalid_words.LastUserFK;"
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
    Dim i As Integer, mRow As ListItem
    sChkSql = "SELECT tbl_invalid_words.InvalidWord " & _
            "From tbl_invalid_words  " & _
            "WHERE (((tbl_invalid_words.InvalidWord) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_invalid_words.InvalidWord;"
    If isFill(1) = True Then
            sValues = txtInput(0).Text _
                  & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                INSERT_DATA "tbl_invalid_words", "InvalidWord,DateAdded,AddedByFK", sValues, ",", True
                frmInvalidWord.btnNew_Load 0
                LvSearchItem frmInvalidWord.lvList(0), txtInput(0).Text
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Invalid Word already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_invalid_words.InvalidWord " & _
            "From tbl_invalid_words  " & _
            "WHERE (((tbl_invalid_words.InvalidWord) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_invalid_words.InvalidWord;"
    If isFill(1) = True Then
        sWhere = "InvalidWord like '" & frmInvalidWord.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(0).Text _
                  & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmInvalidWord.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_invalid_words", "InvalidWord,DateModified,AddedByFK", sValues, sWhere, ",", True
                With frmInvalidWord.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Invalid Word already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
    Else
        With frmInvalidWord
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
        End With
        Me.Caption = "Edit Existing"
    End If
    SetButtonPicture
End Sub

Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        cmdCheck.Picture = .i16x16.ListImages(12).ExtractIcon
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
        For i = 0 To 0
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "Invalid Word"
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

