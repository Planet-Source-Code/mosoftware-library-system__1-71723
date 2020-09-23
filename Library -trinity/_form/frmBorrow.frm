VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBorrow 
   Caption         =   "Borrow/Return Boooks"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14220
   Begin VB.PictureBox picBtnP 
      Height          =   495
      Index           =   3
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   2880
      Width           =   495
      Begin VB.CommandButton cmdReturn 
         Height          =   435
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Return"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.PictureBox picBtnP 
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   2880
      Width           =   495
      Begin VB.CommandButton cmdBorrow 
         Height          =   435
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Borrow"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.Frame fraList 
      Height          =   2805
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdFind 
         Default         =   -1  'True
         Height          =   285
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   300
      End
      Begin VB.ComboBox cboBorrower 
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Text            =   "cboBorrower"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Frame fraProfile 
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
         Height          =   1470
         Left            =   1680
         TabIndex        =   7
         Top             =   1245
         Width           =   4215
         Begin VB.Label Label13 
            BackColor       =   &H00808080&
            Caption         =   "B. TYPE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label9 
            BackColor       =   &H00808080&
            Caption         =   "GENDER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808080&
            Caption         =   "NAME"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   765
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "ID NO."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblID 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   11
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label lblName 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   10
            Top             =   480
            Width           =   3045
         End
         Begin VB.Label lblGen 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   9
            Top             =   720
            Width           =   3045
         End
         Begin VB.Label lblBtype 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   8
            Top             =   960
            Width           =   3045
         End
      End
      Begin VB.Image imgStud 
         Height          =   1840
         Left            =   6120
         Picture         =   "frmBorrow.frx":0000
         Stretch         =   -1  'True
         Top             =   840
         Width           =   6270
      End
      Begin VB.Image ImgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1380
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Borrower ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   18
         Top             =   885
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can Borrow and Return Book in this module."
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
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   480
         Width           =   3450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrower Information"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   2040
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmBorrow.frx":55B32
         Top             =   120
         Width           =   720
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   4
         Left            =   0
         Top             =   240
         Width           =   12495
      End
   End
   Begin VB.Frame fraList 
      Height          =   6855
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   12495
      Begin MSComctlLib.ListView lvList 
         Height          =   5895
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Borrowers Book been borrowed"
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
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books Borrowed"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   2145
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   2
         Left            =   170
         Picture         =   "frmBorrow.frx":6C4E4
         Top             =   120
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
         Width           =   12375
      End
   End
End
Attribute VB_Name = "frmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iIndexBorrow As Integer
Dim sBarcode As String

Public Function SetButtonPic()
    cmdFind.Picture = frmMain.i16x16.ListImages(1).ExtractIcon
    cmdGet.Picture = frmMain.i16x16.ListImages(11).ExtractIcon
    cmdBorrow.Picture = frmMain.iPageEnabled.ListImages(8).ExtractIcon
    cmdReturn.Picture = frmMain.iLv.ListImages(14).ExtractIcon
End Function

Private Sub cboBorrower_Click()
    GetProfile Trim(cboBorrower.Text)
End Sub

Private Sub cmdBook_Click()
    iIndexBorrow = 0
    fraList(0).Visible = True
    fraList(1).Visible = False
    GetBookBorrowed
    Form_Resize
End Sub

Private Sub cmdBorrow_Click()
    If Len(lblID.Caption) > 0 Then
        If isCanBorrowed(lblID.Caption) = True Then
            If isHasOverdue = False Then
                frmBorrowList.Caption = "Borrow"
                frmBorrowList.bStat = True
                frmBorrowList.Show 1
            Else
                MsgBox "You cannot borrow another books. You have a Overdue book(s) to Return. Please Return First the book you borrowed with Overdue Date before you can borrow another Book(s).", vbExclamation, "ErrBorrow"
                cboBorrower.SetFocus
            End If
        Else
            MsgBox "Cannot Borrowed he meets his Max Capacity of Books to be Borrowed.", vbExclamation, "ErrCapacityBorowed"
            cboBorrower.SetFocus
        End If
    Else
        MsgBox "No Current Borrower's Selected.", vbExclamation, "NullBorrowers"
        cboBorrower.SetFocus
    End If
End Sub

Public Function isHasOverdue() As Boolean
    Dim sSQL As String
    sSQL = "SELECT tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date " & _
        "FROM (tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN (tbl_borrowers INNER JOIN tbl_borrow_record ON tbl_borrowers.B_id = tbl_borrow_record.B_id) ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (((tbl_borrow_record.s_return) Like '0') AND ((tbl_borrow_record.r_date)<Date()) AND ((tbl_borrowers.B_id) Like '" & lblID.Caption & "')) " & _
        "GROUP BY tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount < 1 Then
            isHasOverdue = False
        Else
            isHasOverdue = True
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub cmdFind_Click()
    GetProfile Trim(cboBorrower.Text)
End Sub

Private Sub cmdMagazine_Click()
    iIndexBorrow = 1
    fraList(1).Visible = True
    fraList(0).Visible = False
    GetMagazineBorrowed
    Form_Resize
End Sub

Private Sub cmdReturn_Click()
    Dim sSQL As String
    sSQL = "SELECT tbl_borrow_record.br_id " & _
        "From tbl_borrow_record " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & lblID.Caption & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id;"
    If Len(lblID.Caption) > 0 Then
        If isRecordExist(sSQL) = True Then
            frmBorrowList.Caption = "Return"
            frmBorrowList.bStat = False
            frmBorrowList.Show 1
        Else
            MsgBox "No Current Book's to return.", vbExclamation, "NoBorrowedBooks"
            cmdBorrow.SetFocus
        End If
    Else
        MsgBox "No Current Borrower's Selected.", vbExclamation, "NullBorrowers"
        cboBorrower.SetFocus
    End If
End Sub

Private Sub Form_Load()
    iIndexBorrow = 0
    SetButtonPic
    GetList
End Sub

Public Function LvClose(Index As Integer)
    If Index = 0 Then
        Unload Me
    End If
End Function

Public Function LvRefresh(Index As Integer)
    If Index = 0 And Len(cboBorrower.Text) > 0 Then
        GetProfile Trim(cboBorrower.Text)
    End If
End Function

Public Function GetProfile(sIDno As String)
    Dim sSQL As String, lCountRec As Long
    'On Error GoTo errHandler
    sSQL = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.gender, tbl_borrower_type.b_type " & _
        "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
        "WHERE (((tbl_borrowers.B_id) Like '" & sIDno & "'))  " & _
        "GROUP BY tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.gender, tbl_borrower_type.b_type;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        lCountRec = adoRes.RecordCount
        If adoRes.RecordCount > 0 Then
            lblID.Caption = adoRes.Fields("B_id")
            lblName.Caption = adoRes.Fields("fn") & " " & Left(adoRes.Fields("mn"), 1) & ". " & adoRes.Fields("ln")
            lblGen.Caption = adoRes.Fields("gender")
            lblBtype.Caption = adoRes.Fields("b_type")
            On Error GoTo errHandler
            ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & lblID.Caption & ".jpg")
errHandler:
            If err.Number = 53 Then ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & "nopic" & ".jpg")
        Else
            lblID.Caption = ""
            lblName.Caption = ""
            lblGen.Caption = ""
            lblBtype.Caption = ""
            ImgPic.Picture = LoadPicture("")
            lvList(0).ColumnHeaders.Clear
            lvList(0).ListItems.Clear
            MsgBox "Borrower ID does not Exist.", vbExclamation, "ExitPointerException"
            cboBorrower.Text = ""
            cboBorrower.SetFocus
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    If iIndexBorrow = 0 And lCountRec > 0 Then
        fraList(0).Visible = True
        GetBookBorrowed
    End If
    Form_Resize
End Function

Public Function GetList()
    Dim sSQL As String, i As Integer
    sSQL = "SELECT tbl_borrowers.B_id " & _
        "FROM tbl_borrowers;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        cboBorrower.Clear
        Do While Not adoRes.EOF
            cboBorrower.AddItem adoRes.Fields("B_id")
            adoRes.MoveNext
        Loop
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function GetBookBorrowed()
    Dim sSQL As String, i As Integer, mRow As ListItem
    Dim iRecItem As Long
    On Error Resume Next
    sSQL = "SELECT tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
        "FROM tbl_books INNER JOIN (tbl_borrow_record INNER JOIN tbl_reg_books ON tbl_borrow_record.rb_id = tbl_reg_books.rb_id) ON tbl_books.isbn = tbl_reg_books.isbn " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & lblID.Caption & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
    Set lvList(0).SmallIcons = frmMain.iLv
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    lvList(0).ColumnHeaders.Clear
    lvList(0).ListItems.Clear
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        iRecItem = adoRes.RecordCount
        If adoRes.RecordCount > 0 Then
            lvList(0).ColumnHeaders.Add , , , 300
            lvList(0).ColumnHeaders.Add , , "ID", 2000
            lvList(0).ColumnHeaders.Add , , "ISBN", 1400
            lvList(0).ColumnHeaders.Add , , "Title", 4000
            lvList(0).ColumnHeaders.Add , , "Borrow Date", 1200
            lvList(0).ColumnHeaders.Add , , "Return Date", 1200
            lvList(0).ColumnHeaders.Add , , "Date Returned", 1250
            lvList(0).ColumnHeaders.Add , , "Returned?", 1200
            lvList(0).ColumnHeaders.Add , , "Penalty Day(s)", 1250
            lvList(0).ForeColor = vbBlack
            Do While Not adoRes.EOF
                Set mRow = lvList(0).ListItems.Add(, , , , 2)
                mRow.SubItems(1) = adoRes.Fields("br_id")
                mRow.SubItems(2) = adoRes.Fields("isbn")
                mRow.SubItems(3) = adoRes.Fields("title")
                mRow.SubItems(4) = adoRes.Fields("b_date")
                mRow.SubItems(5) = adoRes.Fields("r_date")
                mRow.SubItems(6) = adoRes.Fields("datereturned")
                mRow.SubItems(7) = adoRes.Fields("s_return")
                mRow.SubItems(8) = adoRes.Fields("d_penalty")
                adoRes.MoveNext
            Loop
        Else
            lvList(0).ColumnHeaders.Add , , "", 8000
            lvList(0).ListItems.Add , , "No Current Book(s) Borrowed Record Found.", , 1
            lvList(0).SelectedItem.ForeColor = vbRed
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    If iRecItem > 0 Then
        CheckBorrowedPenalty
    End If
End Function

Public Function GetMagazineBorrowed()
    Dim sSQL As String, i As Integer, mRow As ListItem
    'On Error Resume Next
    sSQL = "SELECT tbl_borrow_rec_mag.br_id, tbl_magazines.issn, tbl_magazines.title, tbl_borrow_rec_mag.b_date, tbl_borrow_rec_mag.r_date, tbl_borrow_rec_mag.datereturned, tbl_borrow_rec_mag.s_return, tbl_borrow_rec_mag.d_penalty " & _
        "FROM tbl_magazines INNER JOIN (tbl_reg_magazines INNER JOIN tbl_borrow_rec_mag ON tbl_reg_magazines.rm_id = tbl_borrow_rec_mag.rm_id) ON tbl_magazines.issn = tbl_reg_magazines.issn " & _
        "WHERE (((tbl_borrow_rec_mag.s_return) Like '0') AND ((tbl_borrow_rec_mag.B_id) Like '" & lblID.Caption & "')) " & _
        "GROUP BY tbl_borrow_rec_mag.br_id, tbl_magazines.issn, tbl_magazines.title, tbl_borrow_rec_mag.b_date, tbl_borrow_rec_mag.r_date, tbl_borrow_rec_mag.datereturned, tbl_borrow_rec_mag.s_return, tbl_borrow_rec_mag.d_penalty;"
    Set lvList(1).SmallIcons = frmMain.iLv
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    lvList(1).ColumnHeaders.Clear
    lvList(1).ListItems.Clear
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            With lvList(1)
            .ColumnHeaders.Add , , "ID", 2000
            .ColumnHeaders.Add , , "ISSN", 2000
            .ColumnHeaders.Add , , "Title", 4000
            .ColumnHeaders.Add , , "Borrow Date", 2000
            .ColumnHeaders.Add , , "Return Date", 2000
            .ColumnHeaders.Add , , "Date Returned", 2000
            .ColumnHeaders.Add , , "Returned?", 2000
            .ColumnHeaders.Add , , "Penalty Day(s)", 2000
            .ForeColor = vbBlack
            End With
            Do While Not adoRes.EOF
                Set mRow = lvList(1).ListItems.Add(, , adoRes.Fields("br_id"), , 6)
                mRow.SubItems(1) = adoRes.Fields("issn")
                mRow.SubItems(2) = adoRes.Fields("title")
                mRow.SubItems(3) = adoRes.Fields("b_date")
                mRow.SubItems(4) = adoRes.Fields("r_date")
                mRow.SubItems(5) = adoRes.Fields("title")
                mRow.SubItems(6) = adoRes.Fields("datereturned")
                mRow.SubItems(7) = adoRes.Fields("s_return")
                mRow.SubItems(8) = adoRes.Fields("d_penalty")
                adoRes.MoveNext
            Loop
        Else
            lvList(1).ColumnHeaders.Add , , "", 8000
            lvList(1).ListItems.Add , , "No Current Magazine(s) Borrowed Record Found.", , 1
            lvList(1).SelectedItem.ForeColor = vbRed
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub Form_Resize()
    Dim i As Integer
    On Error Resume Next
    fraList(3).Move 120, 0, Me.ScaleWidth - 240
    spMag(4).Width = fraList(3).Width
    i = 0
    fraList(i).Move 120, 3360, Me.ScaleWidth - 240, Me.ScaleHeight - (3360 + 120)
    lvList(i).Move 120, 840, fraList(i).Width - 240, fraList(i).Height - (840 + 120)
    spMag(i).Move 0, 240, fraList(i).Width
    imgStud.Left = fraList(3).Width - imgStud.Width
End Sub

Public Function CheckBorrowedPenalty()
    Dim i As Integer, dReturnDate As Date, j As Integer
    For i = 1 To lvList(0).ListItems.Count
        dReturnDate = lvList(0).ListItems(i).SubItems(5)
        If Date > dReturnDate Then
            For j = 1 To lvList(0).ListItems(i).ListSubItems.Count
                lvList(0).ListItems(i).ListSubItems(j).ForeColor = &H80&
            Next
            lvList(0).ListItems(i).SmallIcon = 1
        End If
    Next
End Function
