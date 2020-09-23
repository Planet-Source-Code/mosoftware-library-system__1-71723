VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHoliday_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   1
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1920
         Width           =   300
      End
      Begin VB.CheckBox chkStat 
         Caption         =   "Holiday Status"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
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
         Width           =   2505
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
         Index           =   0
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
         TabIndex        =   10
         Top             =   1680
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
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
         Format          =   62128128
         CurrentDate     =   39492
         MinDate         =   -109205.020833333
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   1
         Left            =   3840
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmHoliday_AE.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   600
         Picture         =   "frmHoliday_AE.frx":617A
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   2
         Left            =   840
         Picture         =   "frmHoliday_AE.frx":65BC
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "DATE"
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
         TabIndex        =   15
         Top             =   2325
         Width           =   1215
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
         Caption         =   "HOLIDAY NAME"
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
         Top             =   1960
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Fill the following information then click Save Button."
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
         Left            =   1440
         TabIndex        =   13
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "HOLIDAY ID"
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
         TabIndex        =   11
         Top             =   1480
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Save"
      Top             =   3360
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3375
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancel"
      Top             =   3375
      Width           =   315
   End
End
Attribute VB_Name = "frmHoliday_AE"
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
        sChkSql = "SELECT tbl_holiday.h_id " & _
                "From tbl_holiday " & _
                "WHERE (((tbl_holiday.h_id) Like '" & txtInput(0).Text & "')) " & _
                "GROUP BY tbl_holiday.h_id;"
        If isRecordExist(sChkSql) = True Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    ElseIf Index = 1 Then
        sChkSql = "SELECT tbl_holiday.h_id " & _
                "From tbl_holiday " & _
                "WHERE (((tbl_holiday.h_name) Like '" & txtInput(1).Text & "')) " & _
                "GROUP BY tbl_holiday.h_id;"
        If isRecordExist(sChkSql) = True Then
            imgChk(1).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(1).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_holiday.h_id " & _
            "From tbl_holiday " & _
            "WHERE (((tbl_holiday.h_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_holiday.h_id;"
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
    sSQL = "SELECT tbl_holiday.DateAdded, tbl_holiday.AddedByFK, tbl_holiday.DateModified, tbl_holiday.LastUserFK " & _
            "From tbl_holiday " & _
            "WHERE (((tbl_holiday.h_id) Like '" & fCur.lvList(0)..selecteditem.SubItems(1) & "')) " & _
            "GROUP BY tbl_holiday.DateAdded, tbl_holiday.AddedByFK, tbl_holiday.DateModified, tbl_holiday.LastUserFK;"
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
    sChkSql = "SELECT tbl_holiday.h_id " & _
            "From tbl_holiday " & _
            "WHERE (((tbl_holiday.h_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_holiday.h_id;"
    If isFill(1) = True Then
            sValues = txtInput(0).Text & "," & txtInput(1).Text & "," & txtDate.Value & "," & chkStat.Value _
                    & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                
                sChkSql = "SELECT tbl_holiday.h_id " & _
                    "From tbl_holiday " & _
                    "WHERE (((tbl_holiday.h_name) Like '" & txtInput(1).Text & "')) " & _
                    "GROUP BY tbl_holiday.h_id;"
                If isRecordExist(sChkSql) = False Then
                    INSERT_DATA "tbl_holiday", "h_id,h_name,h_date,h_status,DateAdded,AddedByFK", sValues, ",", True
                    frmHoliday.btnNew_Load 0
                    LvSearchItem frmHoliday.lvList(0), txtInput(0).Text
                    Unload Me
                Else
                    imgChk(1).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                    MsgBox "Holiday Name already exist. Please change it!", vbExclamation
                    txtInput(1).SetFocus
                End If
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
    sChkSql = "SELECT tbl_holiday.h_id " & _
            "From tbl_holiday " & _
            "WHERE (((tbl_holiday.h_id) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_holiday.h_id;"
    If isFill(1) = True Then
        sWhere = "h_id like '" & frmHoliday.lvList(0).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(1).Text & "," & txtDate.Value & "," & chkStat.Value _
                    & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(0).Text = frmHoliday.lvList(0).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                sChkSql = "SELECT tbl_holiday.h_id " & _
                    "From tbl_holiday " & _
                    "WHERE (((tbl_holiday.h_name) Like '" & txtInput(1).Text & "')) " & _
                    "GROUP BY tbl_holiday.h_id;"
                If isRecordExist(sChkSql) = False Or txtInput(1).Text = frmHoliday.lvList(0).SelectedItem.SubItems(2) Then
                    UPDATE_DATA "tbl_holiday", "h_name,h_date,h_status,DateModified,LastUserFK", sValues, sWhere, ",", True
                    With frmHoliday.lvList(0)
                        .SelectedItem.SubItems(1) = txtInput(0).Text
                        .SelectedItem.SubItems(2) = txtInput(1).Text
                        .SelectedItem.SubItems(3) = txtDate.Value
                        .SelectedItem.SubItems(4) = chkStat.Value
                    End With
                    Unload Me
                Else
                    imgChk(1).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                    MsgBox "Holiday Name already exist. Please change it!", vbExclamation
                    txtInput(1).SetFocus
                End If
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Holiday ID already exist. Please change it!", vbExclamation
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
        txtDate.Value = Date
        txtInput(0).Locked = False
    Else
        With frmHoliday
            txtInput(0).Text = .lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = .lvList(0).SelectedItem.SubItems(2)
            txtDate = .lvList(0).SelectedItem.SubItems(3)
            chkStat = .lvList(0).SelectedItem.SubItems(4)
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
        cmdGen.Picture = .i16x16.ListImages(17).ExtractIcon
        cmdCheck(0).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdCheck(1).Picture = .i16x16.ListImages(12).ExtractIcon
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
            FillMsgBox "Holiday ID"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Holiday Name"
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

