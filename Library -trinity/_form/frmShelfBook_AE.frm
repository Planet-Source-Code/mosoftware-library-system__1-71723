VERSION 5.00
Begin VB.Form frmShelfBook_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   1
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2640
         Width           =   300
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Index           =   1
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
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
         Index           =   3
         Left            =   1320
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2640
         Width           =   1875
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
         Width           =   1905
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   2280
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
         Top             =   2280
         Width           =   2985
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
         Index           =   4
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3000
         Width           =   2985
      End
      Begin VB.CommandButton cmdGen 
         Height          =   285
         Index           =   0
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Index           =   0
         Left            =   4000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1800
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   1
         Left            =   3240
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "CALL NO."
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
         TabIndex        =   20
         Top             =   2660
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "REG. BOOK ID"
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
         TabIndex        =   19
         Top             =   2310
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
         Picture         =   "frmShelfBook_AE.frx":0000
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "SHELF-BOOK ID"
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
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill the following information then click save."
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
         TabIndex        =   17
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHELF-BOOK"
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
         TabIndex        =   16
         Top             =   280
         Width           =   1830
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
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   3020
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1470
         Width           =   1305
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   3240
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
      TabIndex        =   10
      ToolTipText     =   "Save"
      Top             =   4200
      Width           =   315
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4215
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancel"
      Top             =   4215
      Width           =   315
   End
End
Attribute VB_Name = "frmShelfBook_AE"
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
Dim vRegBook As Variant

Private Sub cmdCheck_Click(Index As Integer)
    Dim sChkSql As String
    If Index = 0 Then
        sChkSql = "SELECT tbl_shelfbooks.sb_id " & _
                "From tbl_shelfbooks " & _
                "WHERE (((tbl_shelfbooks.sb_id) Like '" & txtInput(1).Text & "')) " & _
                "GROUP BY tbl_shelfbooks.sb_id;"
        'txtInput(1).Text = Year(Date) & Month(Date) & Day(Date) & Format(Time, "hhmmss")
        If isRecordExist(sChkSql) = True Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    ElseIf Index = 1 Then
        sChkSql = "SELECT tbl_shelfbooks.sb_id " & _
                "From tbl_shelfbooks " & _
                "WHERE (((tbl_shelfbooks.sh_id) Like '" & txtInput(0).Text & "') AND ((tbl_shelfbooks.callno) Like '" & txtInput(3).Text & "')) " & _
                "GROUP BY tbl_shelfbooks.sb_id;"
        If isRecordExist(sChkSql) = True Then
            imgChk(1).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(1).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGen_Click(Index As Integer)
    Dim sChkSql As String
    If Index = 0 Then
        sChkSql = "SELECT tbl_shelfbooks.sb_id " & _
                "From tbl_shelfbooks " & _
                "WHERE (((tbl_shelfbooks.sb_id) Like '" & txtInput(1).Text & "')) " & _
                "GROUP BY tbl_shelfbooks.sb_id;"
        txtInput(1).Text = GenerateID
        If isRecordExist(sChkSql) = True Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    ElseIf Index = 1 Then
        sChkSql = "SELECT Max(tbl_shelfbooks.callno) AS MaxOfcallno " & _
                "From tbl_shelfbooks " & _
                "WHERE (((tbl_shelfbooks.sh_id) Like '" & txtInput(0).Text & "'));"
        txtInput(3).Text = Val(FindFieldValue(sChkSql, "MaxOfcallno")) + 1
    End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
Dim newFormSelect As New frmSelect
    Dim gSQL As String, gLblHead As String, gLblDef As String
    Dim gXY As String, gTitle As String, gColumns As String
    Dim gColWidth As String, gFields As String, gLoop As Integer
    Dim gIcon As Integer, gLvIcon As Integer, sNoRec As String
    Dim sSQL As String
    Dim mRow As ListItem
    Dim SelectedItem As String
    Dim vmbrVal As VbMsgBoxResult
    
    If Index = 0 Then
        gSQL = "SELECT tbl_reg_books.rb_id, tbl_reg_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
        "FROM tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn;"
        gIcon = 2
        gLblHead = "Registered Book(s)"
        gLblDef = "Choose Registered Book(s) then click Select button."
        gXY = "11000,4785"
        gTitle = "Select Registered-Book"
        gColumns = "Registered ID,ISBN,Title,Year Publish,Description"
        gColWidth = "1600,1600,4000,2000,2300"
        gFields = "rb_id,isbn,title,yrpub,desc"
        gLoop = CountSplitItem(gColumns, ",")
        gLvIcon = 3
        sNoRec = "No Current Shelf(s) Info ."
    End If
    
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    
    If Index = 0 And Len(SelectedItem) > 0 Then
        sSQL = "SELECT tbl_shelfbooks.sb_id " & _
            "From tbl_shelfbooks " & _
            "WHERE (((tbl_shelfbooks.sh_id) Like '" & txtInput(0).Text & "') AND ((tbl_shelfbooks.rb_id) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_shelfbooks.sb_id;"
        If isRecordExist(sSQL) = False Then
            sSQL = "SELECT tbl_reg_books.isbn " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.rb_id) Like '" & SelectedItem & "')) " & _
                "GROUP BY tbl_reg_books.isbn;"
            txtInput(2).Text = SelectedItem & "/" & FindFieldValue(sSQL, "isbn")
            CheckBookIfHasCallNo
        Else
            MsgBox "Registered Book Already in this Shelf.", vbExclamation, "RegBookPointerException"
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
    sSQL = "SELECT tbl_shelfbooks.DateAdded, tbl_shelfbooks.AddedByFK, tbl_shelfbooks.DateModified, tbl_shelfbooks.LastUserFK " & _
            "From tbl_shelfbooks " & _
            "WHERE (((tbl_shelfbooks.sb_id) Like '" & fCur.lvList(1).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_shelfbooks.DateAdded, tbl_shelfbooks.AddedByFK, tbl_shelfbooks.DateModified, tbl_shelfbooks.LastUserFK;"
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
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    If isFill(1) = True Then
            vRegBook = Split(txtInput(2).Text, "/")
            sValues = txtInput(1).Text & "," & txtInput(0).Text & "," & vRegBook(0) & "," & txtInput(3).Text & "," & txtInput(4).Text _
                  & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                If cmdGen(1).Visible = False Then
                    imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                    INSERT_DATA "tbl_shelfbooks", "sb_id,sh_id,rb_id,callno,rmarks,DateAdded,AddedByFK", sValues, ",", True
                    frmShelf.btnNew_Load 1
                    frmShelf.lvList(1).FindItem(txtInput(1).Text).Selected = True
                    Unload Me
                Else
                    sSQL = "SELECT tbl_shelfbooks.sb_id " & _
                        "From tbl_shelfbooks " & _
                        "WHERE (((tbl_shelfbooks.sh_id) Like '" & txtInput(0).Text & "') AND ((tbl_shelfbooks.callno) Like '" & txtInput(3).Text & "')) " & _
                        "GROUP BY tbl_shelfbooks.sb_id;"
                    If isRecordExist(sSQL) = False Then
                        imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                        INSERT_DATA "tbl_shelfbooks", "sb_id,sh_id,rb_id,callno,rmarks,DateAdded,AddedByFK", sValues, ",", True
                        frmShelf.btnNew_Load 1
                        LvSearchItem frmShelf.lvList(1), txtInput(0).Text
                        Unload Me
                    Else
                        imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                        MsgBox "Call No. Already Exist in this Shelf. Please" & vbCrLf & " Change Call No. or Click Auto Generate.", vbExclamation, "CallNoExistException"
                    End If
                End If
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
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    If isFill(1) = True Then
        sWhere = "sb_id like '" & frmShelf.lvList(1).SelectedItem.SubItems(1) & "'"
        sValues = txtInput(4).Text _
                    & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Or txtInput(1).Text = frmShelf.lvList(1).SelectedItem.SubItems(1) Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                UPDATE_DATA "tbl_shelfbooks", "rmarks,DateModified,LastUserFK", sValues, sWhere, ",", True
                With frmShelf.lvList(1)
                    .SelectedItem.SubItems(7) = txtInput(4).Text
                End With
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "Shelf-Book already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    ifraIndex = 0
    If bStat = True Then
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtInput(0).Text = frmShelf.lvList(0).SelectedItem.SubItems(1)
        txtInput(1).Text = GenerateID
        cmdCheck(0).Visible = True
        'cmdGen
    Else
        With frmShelf
            txtInput(0).Text = frmShelf.lvList(0).SelectedItem.SubItems(1)
            txtInput(1).Text = frmShelf.lvList(1).SelectedItem.SubItems(2)
            txtInput(2).Text = frmShelf.lvList(1).SelectedItem.SubItems(4)
            txtInput(3).Text = frmShelf.lvList(1).SelectedItem.SubItems(3)
            txtInput(4).Text = frmShelf.lvList(1).SelectedItem.SubItems(9)
        End With
        txtInput(3).Locked = True
        cmdCheck(0).Visible = False
        cmdCheck(1).Visible = False
        cmdGen(0).Visible = False
        cmdGen(1).Visible = False
        cmdGet(0).Visible = False
        Me.Caption = "Edit Existing"
    End If
    SetButtonPicture
End Sub

Public Function CheckBookIfHasCallNo() As String
    Dim sSQL As String
    vRegBook = Split(txtInput(2).Text, "/")
        sSQL = "SELECT tbl_shelfbooks.callno " & _
        "FROM tbl_shelfbooks INNER JOIN tbl_reg_books ON tbl_shelfbooks.rb_id = tbl_reg_books.rb_id " & _
        "WHERE (((tbl_reg_books.isbn) Like '" & vRegBook(1) & "') AND ((tbl_shelfbooks.sh_id) Like '" & txtInput(0).Text & "')) " & _
        "GROUP BY tbl_shelfbooks.callno;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            txtInput(3).Text = adoRes.Fields("callno")
            txtInput(3).Locked = True
            cmdGen(1).Visible = False
            cmdCheck(1).Visible = False
        Else
            txtInput(3).Locked = False
            cmdGen(1).Visible = True
            cmdCheck(1).Visible = True
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdCheck(0).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdCheck(1).Picture = .i16x16.ListImages(12).ExtractIcon
        cmdGen(0).Picture = .i16x16.ListImages(17).ExtractIcon
        cmdGen(1).Picture = .i16x16.ListImages(17).ExtractIcon
        cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
    End With
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        MsgBox "Unabled to Input Barcode. You have to click the button in the upper right corner of the textbox to Insert a Barcode." _
        , vbExclamation, "Unabled Input"
    ElseIf Index = 3 Then
        If txtInput(3).Locked = True Then
            MsgBox "Unabled to Change Call No.", vbExclamation, "ChangePointerException"
        End If
        KeyAscii = isNumber(KeyAscii)
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
        If isNull(txtInput(1).Text) = True Then
            FillMsgBox "Shelf-Book ID"
        End If
        If isNull(txtInput(2).Text) = True Then
            FillMsgBox "Reg. Book ID"
        End If
        If isNull(txtInput(3).Text) = True Then
            FillMsgBox "Call no."
        End If
        If isNull(txtInput(4).Text) = True Then
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


