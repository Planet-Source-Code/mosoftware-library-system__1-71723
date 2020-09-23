VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBorrowList 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraList 
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   3495
      Begin VB.Image Image2 
         Height          =   2655
         Left            =   120
         Picture         =   "frmBorrowList.frx":0000
         Stretch         =   -1  'True
         Top             =   840
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmBorrowList.frx":2CF02
         Top             =   50
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view Receipt Information."
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
         TabIndex        =   14
         Top             =   480
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Information"
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
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1920
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraList 
      Height          =   3615
      Index           =   0
      Left            =   3720
      TabIndex        =   5
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox picBtnP 
         Height          =   375
         Index           =   3
         Left            =   1320
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   840
         Width           =   375
         Begin VB.CommandButton cmdRemove 
            Height          =   315
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Remove Selected"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.PictureBox picBtnP 
         Height          =   375
         Index           =   2
         Left            =   7680
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   10
         Top             =   840
         Width           =   375
         Begin VB.CommandButton cmdCancel 
            Height          =   315
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Close"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.PictureBox picBtnP 
         Height          =   375
         Index           =   1
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   840
         Width           =   375
         Begin VB.CommandButton cmdSave 
            Height          =   315
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Save And Print"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.PictureBox picBtnP 
         Height          =   375
         Index           =   0
         Left            =   840
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   840
         Width           =   375
         Begin VB.CommandButton cmdAdd 
            Height          =   315
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Add Books"
            Top             =   0
            Width           =   315
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3836
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Books Information"
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
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books and their Current Information that you want to Borrow or Return."
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
         TabIndex        =   6
         Top             =   480
         Width           =   5685
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   120
         Picture         =   "frmBorrowList.frx":33754
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
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmBorrowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bStat As Boolean 'true is borrow || false is return
Dim sBarcode As String
Dim dReturnDate As String, sGeneraID As String
Dim iItemBorrowed As Integer, iMaxBorrowed As Integer
Dim dCurrentDate As String

Private Sub cmdAdd_Click()
    If bStat = True Then
        iMaxBorrowed = GetMaxBorrowed(frmBorrow.lblID.Caption)
        iItemBorrowed = itemCountBorrowed(frmBorrow.lblID.Caption)
        If (iItemBorrowed + Val(lvList(0).ListItems.Count)) < iMaxBorrowed Then
            BorrowItem
        Else
            MsgBox "Unabled to Borrow Another Books. Borrower meets his max capacity of Books to be Borrowed.", vbExclamation, "MaxCapacity"
            cmdAdd.SetFocus
        End If
    Else
        ReturnItem
    End If
End Sub

Public Function BorrowItem()
    Dim sSQL As String
    If Len(frmBorrow.lblID.Caption) > 0 Then
        sBarcode = BarcodeValue
        sSQL = "SELECT tbl_reg_books.rb_id " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
                "GROUP BY tbl_reg_books.rb_id;"
        If isRecordExist(sSQL) = True Then
            BorrowBook sBarcode
        Else
            If Len(sBarcode) > 0 Then
                MsgBox "Err: Barcode not Registered.", vbExclamation, "BarcodeExeption"
                cmdAdd.SetFocus
            End If
        End If
    End If
End Function

Public Function ReturnItem()
    Dim sSQL As String
    If Len(frmBorrow.lblID.Caption) > 0 Then
        sBarcode = BarcodeValue
        sSQL = "SELECT tbl_reg_books.rb_id " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
                "GROUP BY tbl_reg_books.rb_id;"
        If isRecordExist(sSQL) = True Then
            ReturnBook sBarcode
        Else
            If Len(sBarcode) > 0 Then
                MsgBox "Err: Barcode not Registered.", vbExclamation, "BarcodeExeption"
                cmdAdd.SetFocus
            End If
        End If
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    If lvList(0).ListItems.Count > 0 Then
        lvList(0).ListItems.Remove lvList(0).SelectedItem.Index
    End If
End Sub

Private Sub cmdSave_Click()
    Dim a As VbMsgBoxResult
    If lvList(0).ListItems.Count > 0 Then
        a = MsgBox("Do you want to save now?", vbQuestion + vbYesNo, "Save")
        If a = vbYes Then
            If bStat = True Then
                SaveItemBorrow
                GenerateBorrowReceipt
            Else
                GenerateReturnReceipt
                SaveItemReturn
            End If
            Unload Me
        End If
        frmBorrow.GetProfile frmBorrow.lblID.Caption
    Else
        MsgBox "Unabled to Save. No Current Record Detected.", vbExclamation, "No Record"
    End If
End Sub

Private Sub Form_Load()
    SetButton
    SetListview
End Sub

Public Function SetListview()
    Set lvList(0).SmallIcons = frmMain.iLv
    lvList(0).ColumnHeaders.Clear
    lvList(0).ListItems.Clear
    lvList(0).ColumnHeaders.Add , , , 300
    lvList(0).ColumnHeaders.Add , , "Reg. Book ID", 1800
    lvList(0).ColumnHeaders.Add , , "Barcode", 1500
    lvList(0).ColumnHeaders.Add , , "ISBN", 1500
    lvList(0).ColumnHeaders.Add , , "Title", 5000
End Function

Public Function SetButton()
    cmdSave.Picture = frmMain.i16x16.ListImages(9).ExtractIcon
    cmdAdd.Picture = frmMain.i16x16.ListImages(2).ExtractIcon
    cmdCancel.Picture = frmMain.i16x16.ListImages(7).ExtractIcon
    cmdRemove.Picture = frmMain.i16x16.ListImages(4).ExtractIcon
End Function

Public Function BorrowBook(sBarcode As String)
    Dim sSQL As String, iMaxBooks As Integer, iRecCount As Integer
    sSQL = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.pending) Like '0') AND ((tbl_reg_books.borrow) Like '0') AND ((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    If isRecordExist(sSQL) = True Then
        sSQL = "SELECT tbl_borrow_record.br_id " & _
            "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
            "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "') AND ((tbl_borrow_record.B_id) Like '" & frmBorrow.lblID.Caption & "') AND ((tbl_borrow_record.datereturned) Like '0')) " & _
            "GROUP BY tbl_borrow_record.br_id;"
        If isRecordExist(sSQL) = False Then
            If isRecordOnLV = False Then
                InsertItem
                'iMaxBorrowed = GetMaxBorrowed(frmBorrow.lblID.Caption)
                'iItemBorrowed = itemCountBorrowed(frmBorrow.lblID.Caption)
                If (iItemBorrowed + Val(lvList(0).ListItems.Count)) < iMaxBorrowed Then
                    cmdAdd_Click
                End If
            Else
                MsgBox "You already add this on cart.", vbExclamation, "AlreadyExist"
                cmdAdd.SetFocus
            End If
        Else
            MsgBox "You already borrowed this book.", vbExclamation, "BorrowException"
            cmdAdd.SetFocus
        End If
    Else
        MsgBox "Books are not available. Book status must be borrowed or pending.", vbExclamation, "BorrowException"
        cmdAdd.SetFocus
    End If
End Function

Public Function ReturnBook(sBarcode As String)
    Dim sSQL As String, iMaxBooks As Integer, iRecCount As Integer
    sSQL = "SELECT tbl_borrow_record.br_id " & _
        "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & frmBorrow.lblID.Caption & "') AND ((tbl_reg_books.barcode) Like '" & sBarcode & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id;"
    If isRecordExist(sSQL) = True Then
        If isRecordOnLV = False Then
            InsertItem
            ReturnItem
        Else
            MsgBox "You already add this on cart.", vbExclamation, "AlreadyExist"
            cmdAdd.SetFocus
        End If
    Else
        MsgBox "You don't borrowed this Book. Book has not Borrowed on your Record.", vbExclamation, "ReturnException"
        cmdAdd.SetFocus
    End If
End Function

Public Function isRecordOnLV() As Boolean
    Dim i As Integer
    isRecordOnLV = False
    For i = 1 To lvList(0).ListItems.Count
        If lvList(0).ListItems(i).SubItems(2) = sBarcode Then
            isRecordOnLV = True
            Exit For
        End If
    Next
End Function

Public Function InsertItem()
    Dim mRow As ListItem, sSQL As String
    On Error Resume Next
    sSQL = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_books.isbn, tbl_books.title " & _
        "FROM tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_books.isbn, tbl_books.title;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
         Set mRow = lvList(0).ListItems.Add(, , , , 2)
         mRow.SubItems(1) = adoRes.Fields("rb_id")
         mRow.SubItems(2) = adoRes.Fields("barcode")
         mRow.SubItems(3) = adoRes.Fields("isbn")
         mRow.SubItems(4) = adoRes.Fields("title")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function SaveItemReturn()
    Dim i As Integer
    dCurrentDate = Date
    For i = 1 To lvList(0).ListItems.Count
        Return_Book lvList(0).ListItems(i).SubItems(2)
    Next
    MsgBox "Record Save.", vbInformation, "Save Successful"
End Function

Public Function SaveItemBorrow()
    Dim i As Integer
    sGeneraID = GenerateID
    dReturnDate = GetReturnDate
    For i = 1 To lvList(0).ListItems.Count
        BookBorrowAdd lvList(0).ListItems(i).SubItems(2)
        sGeneraID = sGeneraID + 1
    Next
    MsgBox "Record Save.", vbInformation, "Save Successful"
End Function

Public Function BookBorrowAdd(sItemBarcode As String)
    Dim sValues As String, sSQL As String
    Dim sWhere As String
    sSQL = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sItemBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    sValues = sGeneraID & "," & frmBorrow.lblID.Caption & "," & FindFieldValue(sSQL, "rb_id") & "," & Date & "," & dReturnDate & "," & "0" _
                & "," & sUserId
    INSERT_DATA "tbl_borrow_record", "br_id,B_id,rb_id,b_date,r_date,s_return,AddedByFK", sValues, ",", False
    sWhere = "barcode like '" & sItemBarcode & "'"
    UPDATE_DATA "tbl_reg_books", "borrow", "1", sWhere, ",", False
End Function

Public Function GenerateBorrowReceipt()
    Dim i As Integer
    Dim sSQL As String
    Dim sReceipt As String
    On Error GoTo errHandler
    sReceipt = ""
    sReceipt = sReceipt & "TRINITY UNIVERSITY" & vbCrLf
    sReceipt = sReceipt & "QUEZON CITY" & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Borrow Date: " & Date & vbCrLf
    sReceipt = sReceipt & "Expected Return Date: " & dReturnDate & vbCrLf & vbCrLf
    sReceipt = sReceipt & "------------------------------------------" & vbCrLf
    For i = 1 To lvList(0).ListItems.Count
        sReceipt = sReceipt & lvList(0).ListItems(i).SubItems(1) & " - " & lvList(0).ListItems(i).SubItems(3) & vbCrLf
        sReceipt = sReceipt & vbTab & lvList(0).ListItems(i).SubItems(4) & vbCrLf
    Next
    sReceipt = sReceipt & "------------------------------------------" & vbCrLf
    sReceipt = sReceipt & "Total Items Borrowed: " & lvList(0).ListItems.Count & vbCrLf
    sReceipt = sReceipt & "Borrowed By: " & frmBorrow.lblName.Caption & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Valid By: " & GetLibrarian & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Library Official Receipt" & vbCrLf
    sReceipt = sReceipt & "of Transaction" & vbCrLf
    MsgBox sReceipt
    Printer.Print sReceipt
errHandler:
    If err.Number = 487 Then
        MsgBox "Error Printing. No Printer Detected.", vbExclamation, "PrinterException"
    End If
End Function

Public Function GetLibrarian() As String
    Dim sSQL As String
    sSQL = "SELECT tbl_users.usrnme " & _
        "From tbl_users " & _
        "WHERE (((tbl_users.uid) Like '" & sUserId & "')) " & _
        "GROUP BY tbl_users.usrnme;"
    GetLibrarian = FindFieldValue(sSQL, "usrnme")
End Function

Public Function GenerateReturnReceipt()
    Dim i As Integer
    Dim sSQL As String, lTotalPenalty As Long
    Dim sReceipt As String, iPenaltyDay As Integer
    On Error GoTo errHandler
    lTotalPenalty = 0
    sReceipt = ""
    sReceipt = sReceipt & "TRINITY UNIVERSITY" & vbCrLf
    sReceipt = sReceipt & "QUEZON CITY" & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Date Returned: " & dReturnDate & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Reg. Book ID" & vbTab & "ISBN" & vbTab & vbTab & "Penalty Days" & vbCrLf
    sReceipt = sReceipt & "-------------------------------------------------------------" & vbCrLf
    For i = 1 To lvList(0).ListItems.Count
        sSQL = "SELECT tbl_borrow_record.r_date " & _
            "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
            "WHERE (((tbl_borrow_record.B_id) Like '" & frmBorrow.lblID.Caption & "') AND ((tbl_reg_books.barcode) Like '" & lvList(0).ListItems(i).SubItems(2) & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
            "GROUP BY tbl_borrow_record.r_date;"
        iPenaltyDay = Val(GetPenaltyDay(FindField(sSQL, "r_date")))
        lTotalPenalty = lTotalPenalty + iPenaltyDay
        sReceipt = sReceipt & lvList(0).ListItems(i).SubItems(1) & vbTab & lvList(0).ListItems(i).SubItems(3) & vbTab & iPenaltyDay & vbCrLf
    Next
    sReceipt = sReceipt & "-------------------------------------------------------------" & vbCrLf
    sReceipt = sReceipt & "Total Items Borrowed: " & lvList(0).ListItems.Count & vbCrLf
    sReceipt = sReceipt & "Total Day's Penalty: " & lTotalPenalty & vbCrLf
    sReceipt = sReceipt & "Borrowed By: " & frmBorrow.lblName.Caption & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Valid By: " & GetLibrarian & vbCrLf & vbCrLf
    
    sReceipt = sReceipt & "Library Official Receipt" & vbCrLf
    sReceipt = sReceipt & "of Transaction" & vbCrLf
    MsgBox sReceipt
    Printer.Print sReceipt
errHandler:
    If err.Number = 487 Then
        MsgBox "Error Printing. No Printer Detected.", vbExclamation, "PrinterException"
    End If
End Function

'This FUNCTION is GENERATOR OF RETURN DATE
Public Function GetReturnDate() As String
    Dim iMaxDay As Integer, i As Integer, d As Integer
    Dim dNowDate As Date
    Dim dDateCheck As Date, sSQL As String
    Dim sNowDate As String
    'On Error Resume Next
    dNowDate = Date
    sSQL = "SELECT tbl_borrower_type.maxdaysborrow " & _
        "From tbl_borrower_type " & _
        "WHERE (((tbl_borrower_type.b_type) Like '" & frmBorrow.lblBtype.Caption & "')) " & _
        "GROUP BY tbl_borrower_type.maxdaysborrow;"
    iMaxDay = FindFieldValue(sSQL, "maxdaysborrow")
    dNowDate = dNowDate + iMaxDay
    d = 0
    Do While Not d = 0
        sNowDate = dNowDate
        If isWeekDay(sNowDate) = False Then
            If isHoliday(sNowDate) = False Then
                Exit Do
            Else
                dNowDate = dNowDate + 1
            End If
        Else
            dNowDate = dNowDate + 1
        End If
    Loop
    GetReturnDate = Str(dNowDate)
End Function

Public Function isHoliday(dDate As String) As Boolean
    Dim sSQL As String
    sSQL = "SELECT tbl_holiday.h_id " & _
        "From tbl_holiday " & _
        "WHERE (((tbl_holiday.h_date) Like '" & dDate & "') AND ((tbl_holiday.h_status) Like '0')) " & _
        "GROUP BY tbl_holiday.h_id;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            isHoliday = True
        Else
            isHoliday = False
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Public Function isWeekDay(dDate As String) As Boolean
    Dim vDay As Variant
    isWeekDay = False
    vDay = Split(Format(dDate, "dddddd"), ",")
    'MsgBox vDay(0)
    Select Case UCase(vDay(0))
        Case UCase("Saturday"): isWeekDay = True
        Case UCase("Sunday"): isWeekDay = True
    End Select
End Function

Public Function Return_Book(sItemBarcode As String)
    Dim sWhere As String, sValues As String, sSQL As String, iPenaltyDay As Long
    sSQL = "SELECT tbl_borrower_type.penaltystat " & _
        "From tbl_borrower_type  " & _
        "WHERE (((tbl_borrower_type.b_type) Like '" & frmBorrow.lblBtype.Caption & "'))  " & _
        "GROUP BY tbl_borrower_type.penaltystat;"
    If FindField(sSQL, "penaltystat") = "1" Then
        sSQL = "SELECT tbl_borrow_record.r_date " & _
            "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
            "WHERE (((tbl_borrow_record.B_id) Like '" & frmBorrow.lblID.Caption & "') AND ((tbl_reg_books.barcode) Like '" & sItemBarcode & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
            "GROUP BY tbl_borrow_record.r_date;"
        iPenaltyDay = Val(GetPenaltyDay(FindField(sSQL, "r_date")))
    Else
        iPenaltyDay = 0
    End If
    
    sSQL = "SELECT tbl_borrow_record.br_id " & _
        "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & frmBorrow.lblID.Caption & "') AND ((tbl_reg_books.barcode) Like '" & sItemBarcode & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id;"
    sWhere = "br_id like '" & FindFieldValue(sSQL, "br_id") & "'"
    sValues = dCurrentDate & ",1," & iPenaltyDay & "," & sUserId
    UPDATE_DATA "tbl_borrow_record", "datereturned,s_return,d_penalty,LastUserFK", sValues, sWhere, ",", False
    sWhere = "barcode like '" & sItemBarcode & "'"
    UPDATE_DATA "tbl_reg_books", "borrow", "0", sWhere, ",", False
End Function

Public Function GetPenaltyDay(dDate As Date) As Long
    If Date > dDate Then
        GetPenaltyDay = Date - dDate
    Else
        GetPenaltyDay = 0
    End If
End Function

Private Sub imgTop_Click()
    GenerateReturnReceipt
End Sub

