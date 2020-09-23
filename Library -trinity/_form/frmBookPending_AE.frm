VERSION 5.00
Begin VB.Form frmBookPending_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   2775
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2775
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Save"
      Top             =   2760
      Width           =   315
   End
   Begin VB.Frame fraList 
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   1320
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
         Height          =   765
         Index           =   1
         Left            =   1440
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1680
         Width           =   2505
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
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1320
         Width           =   2505
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pending"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1140
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
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "BARCODE"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1365
         Width           =   1305
      End
      Begin VB.Image imgExc 
         Height          =   360
         Index           =   0
         Left            =   4275
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "REMARKS"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   480
         Picture         =   "frmBookPending_AE.frx":0000
         Top             =   240
         Width           =   720
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
Attribute VB_Name = "frmBookPending_AE"
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
Dim sBarcode As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    'txtInput(1).Text = Year(Date) & Month(Date) & Day(Date) & Format(Time, "hhmmss")
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGen_Click()
    Dim sChkSql As String
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    txtInput(1).Text = GenerateID
    If isRecordExist(sChkSql) = True Then
        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    Else
        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
    End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Dim sWhere As String
    sBarcode = BarcodeValue
    'MsgBox sBarcode
    'sWhere = "barcode like '" & frmBookPending.lvList(0).SelectedItem.SubItems(1) & "'"
    If isBarcodeBookExist(sBarcode, False) = True And Len(sBarcode) > 0 Then
        If isBorrow(sBarcode) = True Then
            txtInput(0).Text = sBarcode
            'UPDATE_DATA tbl_reg_books, "pending,remakrs", "1," & txtInput(1).Text, sWhere, True
        Else
            MsgBox "You cannot Activate this Book in Pending. Book is Borrowed.", vbExclamation, "ErrPending"
            cmdGet(0).SetFocus
        End If
    Else
        If Len(sBarcode) > 0 Then
            MsgBox "Book not Registered. Please Re-enter Barcode again.", vbExclamation, "ErrPending"
            cmdGet(0).SetFocus
        End If
    End If
End Sub

Public Function isBorrow(sBarcodeItem As String) As Boolean
    Dim sIBBE As String
    sIBBE = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sBarcodeItem & "') AND ((tbl_reg_books.borrow) Like '0')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sIBBE, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            isBorrow = True
        Else
            isBorrow = False
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function isPending(sBarcodeItem As String) As Boolean
    Dim sIBBE As String
    sIBBE = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sBarcodeItem & "') AND ((tbl_reg_books.pending) Like '1')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sIBBE, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            isPending = True
        Else
            isPending = False
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub cmdMod_Click()
    On Error Resume Next
    Dim strMess As String
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    sSQL = "SELECT tbl_reg_books.DateAdded, tbl_reg_books.AddedByFK, tbl_reg_books.DateModified, tbl_reg_books.LastUserFK " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.rb_id) Like '" & fCur.lvList(1).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_reg_books.DateAdded, tbl_reg_books.AddedByFK, tbl_reg_books.DateModified, tbl_reg_books.LastUserFK;"
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
    Dim sChkSql As String, sWhere As String
    Dim i As Integer
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.barcode) Like '" & txtInput(0).Text & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    sWhere = "rb_id like '" & FindFieldValue(sChkSql, "rb_id") & "'"
    If isFill(1) = True Then
         If isPending(sBarcode) = False Then
            sValues = "1," & txtInput(1).Text
            UPDATE_DATA "tbl_reg_books", "pending,[remarks]", sValues, sWhere, ",", True
            Unload Me
            frmBookPending.Lv_MainInfo
        Else
            MsgBox "Book Status already in Pending.", vbExclamation, "ErrPending"
            cmdGet(0).SetFocus
        End If
    End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim sChkSql As String
    Dim sWhere As String
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    If isFill(1) = True Then
        sWhere = "rb_id like '" & frmBookPending.lvList(1).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(2).Text & "," & txtInput(3).Text _
                    & "," & Date & "," & sUserId
        If isPending(sBarcode) = True Then
            sValues = "1," & txtInput(1).Text
            UPDATE_DATA tbl_reg_books, "pending,remakrs", "0," & txtInput(1).Text, sWhere, True
            Unload Me
            frmBookPending.Lv_MainInfo
            Unload Me
        Else
            MsgBox "Book Status already in Pending.", vbExclamation, "ErrPending"
            cmdGet(0).SetFocus
        End If
    End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
    Else
        With frmBookPending
            txtInput(0).Text = frmBookPending.lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = frmBookPending.lvList(0).SelectedItem.SubItems(2)
            txtInput(2).Text = frmBookPending.lvList(0).SelectedItem.SubItems(3)
            txtInput(3).Text = frmBookPending.lvList(0).SelectedItem.SubItems(4)
        End With
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
        cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
    End With
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        MsgBox "Unabled to Input Barcode. You have to click the button in the upper right corner of the textbox to Insert a Barcode." _
        , vbExclamation, "Unabled Input"
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
            FillMsgBox "Barcode"
        End If
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Remarks"
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


