VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMagazine_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   5535
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
         Index           =   3
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3960
         Width           =   3225
      End
      Begin VB.ComboBox cboEdition 
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
         ItemData        =   "frmMagazine_AE.frx":0000
         Left            =   1440
         List            =   "frmMagazine_AE.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4680
         Width           =   3255
      End
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frmMagazine_AE.frx":008F
         Left            =   1440
         List            =   "frmMagazine_AE.frx":0099
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4320
         Width           =   3255
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
         Height          =   885
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2160
         Width           =   3225
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
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3120
         Width           =   3195
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   5040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
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
         CustomFormat    =   "ddddd"
         Format          =   20709376
         CurrentDate     =   39511
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "EDITION"
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
         TabIndex        =   21
         Top             =   4710
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "VOLUME"
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
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "DATE PUBLISH"
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
         TabIndex        =   19
         Top             =   5085
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PERIODIC TYPE"
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
         TabIndex        =   18
         Top             =   4350
         Width           =   1335
      End
      Begin VB.Image imgExc 
         Height          =   480
         Index           =   0
         Left            =   4275
         Top             =   840
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   480
         Picture         =   "frmMagazine_AE.frx":00B2
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PERIODIC NAME"
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
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You have to fill the following information then click Save Button to Finish."
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
         Height          =   675
         Index           =   0
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodical"
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
         TabIndex        =   15
         Top             =   270
         Width           =   1350
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "DESCRIPTION "
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
         Top             =   3165
         Width           =   1335
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
         TabIndex        =   13
         Top             =   1710
         Width           =   1425
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save"
      Top             =   5640
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   5655
      Width           =   315
   End
End
Attribute VB_Name = "frmMagazine_AE"
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
    sChkSql = "SELECT tbl_magazines.issn " & _
            "From tbl_magazines " & _
            "WHERE (((tbl_magazines.issn) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_magazines.issn;"
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
    sSQL = "SELECT tbl_magazines.DateAdded, tbl_magazines.AddedByFK, tbl_magazines.DateModified, tbl_magazines.LastUserFK " & _
            "From tbl_magazines " & _
            "WHERE (((tbl_magazines.issn) Like '" & fCur.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_magazines.DateAdded, tbl_magazines.AddedByFK, tbl_magazines.DateModified, tbl_magazines.LastUserFK;"
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
    sChkSql = "SELECT tbl_magazines.issn " & _
            "From tbl_magazines " & _
            "WHERE (((tbl_magazines.issn) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_magazines.issn;"
    If isFill(1) = True Then
            sValues = txtInput(0).Text & "," & txtInput(1).Text & "," & txtInput(2).Text & "," & cboType.Text & "," & txtInput(3).Text & "," & cboEdition.Text & "," & dtDate.Value _
                & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                INSERT_DATA "tbl_magazines", "issn,title,dsc,s_type,s_vol,s_edition,d_publish,DateAdded,AddedByFK", sValues, ",", True
                frmMagazine.btnNew_Load 0
                LvSearchItem frmMagazine.lvList(0), txtInput(0).Text
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "ISSN already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_magazines.issn " & _
            "From tbl_magazines " & _
            "WHERE (((tbl_magazines.issn) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_magazines.issn;"
    If isFill(1) = True Then
        sWhere = "issn like '" & frmMagazine.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(1).Text & "," & txtInput(2).Text & "," & cboType.Text & "," & txtInput(3).Text & "," & cboEdition.Text & "," & dtDate.Value _
                & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmMagazine.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_magazines", "title,dsc,s_type,s_vol,s_edition,d_publish,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmMagazine.lvList(0)
                    .SelectedItem.SubItems(1) = txtInput(0).Text
                    .SelectedItem.SubItems(2) = txtInput(1).Text
                    .SelectedItem.SubItems(4) = txtInput(3).Text
                    .SelectedItem.SubItems(7) = txtInput(2).Text
                    .SelectedItem.SubItems(3) = cboType.Text
                    .SelectedItem.SubItems(5) = cboEdition.Text
                    .SelectedItem.SubItems(6) = dtDate.Value
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Periodic already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    On Error Resume Next
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Locked = False
    Else
        With frmMagazine
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(2)
            txtInput(3).Text = .lvList(0).SelectedItem.SubItems(4)
            txtInput(2).Text = .lvList(0).SelectedItem.SubItems(7)
            cboType.Text = .lvList(0).SelectedItem.SubItems(3)
            cboEdition.Text = .lvList(0).SelectedItem.SubItems(5)
            dtDate.Value = .lvList(0).SelectedItem.SubItems(6)
        End With
        txtInput(0).Locked = True
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
        cmdCheck.Picture = .i16x16.ListImages(12).ExtractIcon
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
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
        For i = 0 To 3
        If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If isNull(txtInput(0).Text) = True Then
            FillMsgBox "ISSN"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Periodic name"
        End If
        If isNull(txtInput(2).Text) = True Then
           FillMsgBox "Description"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Volume"
        End If
        If isNull(cboType.Text) = True Then
            FillMsgBox "Periodic Type"
        End If
        If isNull(cboEdition.Text) = True Then
            FillMsgBox "Periodic Type"
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

