VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAuthor_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   2880
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save"
      Top             =   2880
      Width           =   315
   End
   Begin VB.Frame fraList 
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4815
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   57671683
         CurrentDate     =   39496
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   300
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
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
         Index           =   1
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   2745
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
         Width           =   1905
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   4620
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3240
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "AUTHOR ID"
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
         TabIndex        =   14
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHOR"
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
         TabIndex        =   13
         Top             =   300
         Width           =   1215
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
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "AUTHOR NAME"
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   480
         Picture         =   "frmAuthor_AE.frx":0000
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgExc 
         Height          =   360
         Index           =   0
         Left            =   4275
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "YEAR BORN"
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
         TabIndex        =   10
         Top             =   2325
         Width           =   1215
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
End
Attribute VB_Name = "frmAuthor_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bStat As Boolean
Public fCur As Form

Dim sSQL As String
Dim SelectedItem As String
'Dim bSFocus As Boolean
Dim sMsgBox As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_authors.auid " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_authors.auid;"
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_authors.auid " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_authors.auid;"
    txtInput(0).Text = GenerateID
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_authors.DateAdded, tbl_authors.AddedByFK, tbl_authors.DateModified, tbl_authors.LastUserFK " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_authors.DateAdded, tbl_authors.AddedByFK, tbl_authors.DateModified, tbl_authors.LastUserFK;"
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
    sChkSql = "SELECT tbl_authors.auid " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_authors.auid;"
    If isFill(1) = True Then
        sValues = txtInput(0).Text & "," & txtInput(1) & "," & dtDate.Year _
                  & "," & Date & "," & sUserId
        If isRecordExist(sChkSql) = False Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
            INSERT_DATA "tbl_authors", "auid,author,yrborn,DateAdded,AddedByFK", sValues, ",", True
            frmAuthor.btnNew_Load 0
            LvSearchItem frmAuthor.lvList(0), txtInput(0).Text
            Unload Me
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
            MsgBox "Author ID already exist. Please change it!", vbExclamation
            txtInput(0).SetFocus
        End If
    End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_authors.auid " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_authors.auid;"
    If isFill(1) = True Then
        sWhere = "auid like '" & frmAuthor.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(1).Text & "-" & dtDate.Year _
                  & "-" & Date & "-" & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmAuthor.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_authors", "author-yrborn-DateModified-LastUserFK", sValues, sWhere, "-", True
                With frmAuthor.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(3) = dtDate.Year
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Author ID exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    On Error Resume Next
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Text = GenerateID
        cmdCheck.Visible = True
    Else
        With frmAuthor
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(2)
            dtDate.Year = .lvList(0).SelectedItem.SubItems(3)
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
        For i = 0 To 1
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "Author ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Author Name"
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


