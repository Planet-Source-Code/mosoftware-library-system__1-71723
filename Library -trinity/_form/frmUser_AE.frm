VERSION 5.00
Begin VB.Form frmUser_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   315
   End
   Begin VB.Frame fraList 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   1
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2620
         Width           =   300
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1635
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
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1680
         PasswordChar    =   "*"
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3360
         Width           =   2265
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
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1680
         PasswordChar    =   "*"
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3000
         Width           =   2265
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
         Left            =   1680
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2640
         Width           =   2265
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
         Left            =   1680
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2160
         Width           =   3105
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   4695
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   0
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1635
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
         Left            =   1680
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1680
         Width           =   2265
      End
      Begin VB.Frame fraChkBx 
         BorderStyle     =   0  'None
         Height          =   650
         Left            =   1560
         TabIndex        =   20
         Top             =   3720
         Width           =   2415
         Begin VB.CheckBox chkT 
            Caption         =   "Transaction"
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
            Left            =   1200
            TabIndex        =   24
            Top             =   0
            Width           =   1215
         End
         Begin VB.CheckBox chkC 
            Caption         =   "Cashier"
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
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkO 
            Caption         =   "Opac"
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
            Left            =   1200
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkA 
            Caption         =   "Admin"
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
            Left            =   120
            TabIndex        =   22
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "VERYFY PASSWORD"
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
         Top             =   3405
         Width           =   1575
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   2
         Left            =   3960
         Top             =   3360
         Width           =   240
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   1
         Left            =   3960
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1700
         Width           =   1665
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "USERNAME"
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
         Top             =   2670
         Width           =   1665
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "COMPLETE NAME"
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
         TabIndex        =   16
         Top             =   2190
         Width           =   1665
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You have to fill the following information then click Next Button or Save button to finish Adding New User."
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
         Height          =   555
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   2415
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3960
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PASSWORD"
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
         Top             =   3040
         Width           =   1665
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   720
         Picture         =   "frmUser_AE.frx":0000
         Top             =   360
         Width           =   720
      End
      Begin VB.Image imgExc 
         Height          =   360
         Index           =   0
         Left            =   4320
         Top             =   1080
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
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   4575
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save"
      Top             =   4560
      Width           =   315
   End
End
Attribute VB_Name = "frmUser_AE"
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
        sChkSql = "SELECT tbl_users.uid " & _
                "From tbl_users " & _
                "WHERE (((tbl_users.uid) Like '" & txtInput(0).Text & "')) " & _
                "GROUP BY tbl_users.uid;"
    ElseIf Index = 1 Then
        sChkSql = "SELECT tbl_users.uid " & _
                "From tbl_users " & _
                "WHERE (((tbl_users.usrnme) Like '" & txtInput(2).Text & "')) " & _
                "GROUP BY tbl_users.uid;"
    End If
    If isRecordExist(sChkSql) = True Then
        imgChk(Index).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(Index).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_users.uid " & _
            "From tbl_users " & _
            "WHERE (((tbl_users.uid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_users.uid;"
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
    sSQL = "SELECT tbl_users.DateAdded, tbl_users.AddedByFK, tbl_users.DateModified, tbl_users.LastUserFK " & _
            "From tbl_users " & _
            "WHERE (((tbl_users.uid) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_users.DateAdded, tbl_users.AddedByFK, tbl_users.DateModified, tbl_users.LastUserFK;"
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
    sChkSql = "SELECT tbl_users.uid " & _
            "From tbl_users " & _
            "WHERE (((tbl_users.uid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_users.uid;"
    If isFill(1) = True Then
            sValues = txtInput(0).Text & "," & txtInput(2).Text & "," & txtInput(1).Text & "," & ENCRYPT(txtInput(3).Text) & "," & chkA.Value & "," & chkC.Value & "," & chkT.Value & "," & chkO.Value & _
                   "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                sChkSql = "SELECT tbl_users.uid " & _
                    "From tbl_users " & _
                    "WHERE (((tbl_users.usrnme) Like '" & txtInput(2).Text & "')) " & _
                    "GROUP BY tbl_users.uid;"
                If isRecordExist(sChkSql) = False Then
                    imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                    INSERT_DATA "tbl_users", "uid,usrnme,completename,pass,admn,cashr,transction,opac,DateAdded,AddedByFK", sValues, ",", True
                    frmUser.btnNew_Load 0
                    LvSearchItem frmUser.lvList(0), txtInput(0).Text
                    Unload Me
                Else
                    MsgBox "Username already exist!", vbExclamation, "UsernamePointerException"
                    txtInput(2).SetFocus
                End If
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "User ID already exist. Please change it!", vbExclamation
                txtInput(1).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_users.uid " & _
            "From tbl_users " & _
            "WHERE (((tbl_users.uid) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_users.uid;"
    If isFill(1) = True Then
        sWhere = "uid like '" & frmUser.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(0).Text & "," & txtInput(2).Text & "," & txtInput(1).Text & "," & chkA.Value & "," & chkC.Value & "," & chkT.Value & "," & chkO.Value & _
                   "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(2).Text = frmUser.lvList(0).SelectedItem.SubItems(2) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_users", "uid,usrnme,completename,admn,cashr,transction,opac,DateModified,AddedByFK", sValues, sWhere, ",", True
                With frmUser.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(2).Text
                    .SelectedItem.SubItems(3) = txtInput(1).Text
                    .SelectedItem.SubItems(4) = chkA.Value
                    .SelectedItem.SubItems(5) = chkC.Value
                    .SelectedItem.SubItems(6) = chkT.Value
                    .SelectedItem.SubItems(7) = chkO.Value
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Username already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        cmdCheck(0).Visible = True
        cmdGen.Visible = True
        fraChkBx.Top = 3720
    Else
        With frmUser
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(3)
            txtInput(2).Text = .lvList(0).SelectedItem.SubItems(2)
            chkA.Value = .lvList(0).SelectedItem.SubItems(4)
            chkC.Value = .lvList(0).SelectedItem.SubItems(5)
            chkT.Value = .lvList(0).SelectedItem.SubItems(6)
            chkO.Value = .lvList(0).SelectedItem.SubItems(7)
        End With
        fraChkBx.Top = 3000
        lblCaption(3).Visible = False
        lblCaption(4).Visible = False
        txtInput(3).Visible = False
        txtInput(4).Visible = False
        cmdCheck(0).Visible = False
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
        cmdCheck(0).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdCheck(1).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen.Picture = .i16x16.ListImages(17).ExtractIcon
        'cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
    End With
End Sub

Private Sub txtInput_Change(Index As Integer)
    If Len(txtInput(4).Text) > 0 Then
        If Not txtInput(3).Text = txtInput(4).Text Then
            imgChk(2).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(2).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    Else
        imgChk(2).Picture = LoadPicture("")
    End If
End Sub

Private Function isFill(iStep As Integer) As Boolean
    Dim sBeginMsg As String, b As Integer
    Dim sEndMsg As String
    Dim iNotFill As Integer
    Dim i As Integer
    Dim iStart As Integer
    Dim iSplitItem As Integer
    Dim sSplitedItem As Variant
    iNotFill = 0
    sMsgBox = ""
    If bStat = True Then
        b = 4
    Else
        b = 2
    End If
    sBeginMsg = "You forgot to fill the following "
    sEndMsg = "."
    If iStep = 1 Then
        For i = 0 To b
            If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "User ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Complete Name"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Username"
        End If
        If bStat = True Then
            If isNull(txtInput(3).Text) = True Then
                FillMsgBox "Password"
            End If
            If isNull(txtInput(4).Text) = True Then
                FillMsgBox "Verefication Password"
            End If
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
