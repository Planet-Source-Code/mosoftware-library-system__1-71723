VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPenaltyPayment 
   Caption         =   "Penalty Payment"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame fraList 
      Height          =   2925
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   12495
      Begin VB.ComboBox cboBorrower 
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Text            =   "cboBorrower"
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdFind 
         Default         =   -1  'True
         Height          =   285
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   300
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   300
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
         Height          =   1575
         Left            =   1680
         TabIndex        =   5
         Top             =   1150
         Width           =   4150
         Begin VB.Label lblFee 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Top             =   1200
            Width           =   2205
         End
         Begin VB.Label Label7 
            BackColor       =   &H00808080&
            Caption         =   "PENALTY FEE / DAY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   1635
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
            TabIndex        =   13
            Top             =   960
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
            TabIndex        =   12
            Top             =   720
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
            TabIndex        =   11
            Top             =   480
            Width           =   3045
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
            TabIndex        =   10
            Top             =   240
            Width           =   3045
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
            TabIndex        =   9
            Top             =   240
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
            TabIndex        =   8
            Top             =   480
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
            TabIndex        =   7
            Top             =   720
            Width           =   765
         End
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
            TabIndex        =   6
            Top             =   960
            Width           =   765
         End
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmPenaltyPayment.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penalty Payment"
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
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Borrower ID to view list of Penalty."
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
         TabIndex        =   15
         Top             =   480
         Width           =   3465
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
         TabIndex        =   14
         Top             =   885
         Width           =   870
      End
      Begin VB.Image ImgPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1380
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1270
         Width           =   1485
      End
      Begin VB.Image imgStud 
         Height          =   1995
         Left            =   6120
         Picture         =   "frmPenaltyPayment.frx":169B2
         Stretch         =   -1  'True
         Top             =   840
         Width           =   6270
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
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   15015
      Begin MSComctlLib.ListView lvList 
         Height          =   2535
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4471
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
      Begin VB.Image Image3 
         Height          =   720
         Index           =   2
         Left            =   170
         Picture         =   "frmPenaltyPayment.frx":5D834
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books has penalty"
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
         Width           =   2325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of books that has penalty"
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
         TabIndex        =   19
         Top             =   480
         Width           =   3075
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   15015
      End
   End
   Begin VB.Frame fraRec 
      Height          =   3135
      Left            =   11280
      TabIndex        =   21
      Top             =   6380
      Width           =   3855
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Save Transaction"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Tendered"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   645
      End
      Begin VB.Image Image3 
         Height          =   600
         Index           =   3
         Left            =   120
         Picture         =   "frmPenaltyPayment.frx":741E6
         Stretch         =   -1  'True
         Top             =   165
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Receipt"
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
         Index           =   3
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   2
         Left            =   0
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Image imgBottom 
      Height          =   2970
      Left            =   120
      Picture         =   "frmPenaltyPayment.frx":7AA38
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   11085
   End
End
Attribute VB_Name = "frmPenaltyPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iIndexBorrow As Integer
Dim sBarcode As String
Dim dblPenaltyFee As Double
Dim sGeneratedID As String

Public Function SetButtonPic()
    cmdFind.Picture = frmMain.i16x16.ListImages(1).ExtractIcon
    cmdGet.Picture = frmMain.i16x16.ListImages(11).ExtractIcon
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
    Dim sSQL As String
    If Len(lblID.Caption) > 0 Then
        sBarcode = BarcodeValue
        If iIndexBorrow = 0 Then
            sSQL = "SELECT tbl_reg_books.rb_id " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
                "GROUP BY tbl_reg_books.rb_id;"
            If isRecordExist(sSQL) = True Then
                BorrowBook sBarcode
            Else
                If Len(sBarcode) > 0 Then
                    MsgBox "Err: Barcode not Registered.", vbExclamation, "BarcodeExeption"
                    cboBorrower.SetFocus
                End If
            End If
        ElseIf iIndexBorrow = 1 Then
            sSQL = "SELECT tbl_reg_magazines.rm_id " & _
                "From tbl_reg_magazines " & _
                "WHERE (((tbl_reg_magazines.barcode) Like '" & sBarcode & "')) " & _
                "GROUP BY tbl_reg_magazines.rm_id;"
            If isRecordExist(sSQL) = True Then
                BorrowMagazine sBarcode
            Else
                If Len(sBarcode) > 0 Then
                    MsgBox "Err: Barcode not Registered.", vbExclamation, "BarcodeExeption"
                    cboBorrower.SetFocus
                End If
            End If
        End If
    End If
End Sub

Public Function BorrowBook(sBarcode As String)
    Dim sSQL As String, iMaxBooks As Integer, iRecCount As Integer
    sSQL = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.pending) Like '0') AND ((tbl_reg_books.borrow) Like '0') AND ((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    If isRecordExist(sSQL) = True Then
        sSQL = "SELECT tbl_borrow_record.br_id " & _
            "FROM tbl_reg_books INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
            "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "') AND ((tbl_borrow_record.B_id) Like '" & lblID.Caption & "') AND ((tbl_borrow_record.datereturned) Like '0')) " & _
            "GROUP BY tbl_borrow_record.br_id;"
        If isRecordExist(sSQL) = False Then
            sSQL = "SELECT tbl_borrower_type.maxdaysborrow " & _
                "From tbl_borrower_type " & _
                "WHERE (((tbl_borrower_type.b_type) Like '" & lblBtype.Caption & "')) " & _
                "GROUP BY tbl_borrower_type.maxdaysborrow;"
            iMaxBooks = FindFieldValue(sSQL, "maxdaysborrow")
            sSQL = "SELECT Count(tbl_borrow_record.br_id) AS CountOfbr_id " & _
                "From tbl_borrow_record " & _
                "WHERE (((tbl_borrow_record.B_id) Like '" & lblID.Caption & "') AND ((tbl_borrow_record.s_return) Like '0'));"
            iRecCount = FindFieldValue(sSQL, "CountOfbr_id")
            If iRecCount < iMaxBooks Then
                BookBorrowAdd
                GetProfile lblID.Caption
            Else
                MsgBox "You have been meet the max limit of Books to be Borrowed.", vbExclamation, "BorrowException"
                cmdBorrow.SetFocus
            End If
        Else
            MsgBox "You already borrowed this book.", vbExclamation, "BorrowException"
            cmdBorrow.SetFocus
        End If
    Else
        MsgBox "Books are not available. Book status must be borrowed or pending.", vbExclamation, "BorrowException"
        cmdBorrow.SetFocus
    End If
End Function

Public Function BookBorrowAdd()
    Dim sValues As String, sSQL As String
    Dim sWhere As String
    sSQL = "SELECT tbl_reg_books.rb_id " & _
        "From tbl_reg_books " & _
        "WHERE (((tbl_reg_books.barcode) Like '" & sBarcode & "')) " & _
        "GROUP BY tbl_reg_books.rb_id;"
    sValues = GenerateID & "," & lblID.Caption & "," & FindFieldValue(sSQL, "rb_id") & "," & Date & "," & GetReturnDate & "," & "0" _
                & "," & sUserId
    INSERT_DATA "tbl_borrow_record", "br_id,B_id,rb_id,b_date,r_date,s_return,AddedByFK", sValues, ",", True
    sWhere = "barcode like '" & sBarcode & "'"
    UPDATE_DATA "tbl_reg_books", "borrow", "1", sWhere, ",", False
End Function

Public Function UpdateRegisteredBook(iBorrow As Integer)
    
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
    Dim sSQL As String, a As VbMsgBoxResult
    If Len(lblID.Caption) > 0 Then
        a = MsgBox("Do you want to Return that you Select?", vbQuestion + vbYesNo, "Return")
        If a = vbYes Then
            If Len(lblID.Caption) > 0 Then
                If iIndexBorrow = 0 Then
                    If lvList(0).ColumnHeaders.Count > 1 Then
                        Return_Book lvList(0).SelectedItem.SubItems(4)
                        GetProfile lblID.Caption
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    SaveTransaction
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

Public Function GetProfile(sIDno As String)
    Dim sSQL As String, lCountRec As Long
    'On Error GoTo errHandler
    sSQL = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.gender, tbl_borrower_type.b_type, tbl_borrower_type.p_fee " & _
        "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
        "WHERE (((tbl_borrowers.B_id) Like '" & sIDno & "'))  " & _
        "GROUP BY tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrowers.gender, tbl_borrower_type.b_type, tbl_borrower_type.p_fee;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        lCountRec = adoRes.RecordCount
        If adoRes.RecordCount > 0 Then
            lblID.Caption = adoRes.Fields("B_id")
            lblName.Caption = adoRes.Fields("fn") & " " & Left(adoRes.Fields("mn"), 1) & " " & adoRes.Fields("ln")
            lblGen.Caption = adoRes.Fields("gender")
            lblBtype.Caption = adoRes.Fields("b_type")
            lblFee.Caption = Format(adoRes.Fields("p_fee"), "0.00")
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
            lvList(1).ColumnHeaders.Clear
            lvList(1).ListItems.Clear
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
    If lCountRec > 0 Then
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
    On Error Resume Next
    sSQL = "SELECT tbl_borrow_record.br_id, tbl_reg_books.rb_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
        "FROM (tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN tbl_borrow_record ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (((tbl_borrow_record.B_id) Like '" & lblID.Caption & "') AND ((tbl_borrow_record.d_penalty)>0) AND ((tbl_borrow_record.s_paid) Like '0')) " & _
        "GROUP BY tbl_borrow_record.br_id, tbl_reg_books.rb_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
    Set lvList(0).SmallIcons = frmMain.iLv
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    lvList(0).ColumnHeaders.Clear
    lvList(0).ListItems.Clear
    'MsgBox sSQL
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        If adoRes.RecordCount > 0 Then
            lvList(0).ColumnHeaders.Add , , , 300
            lvList(0).ColumnHeaders.Add , , "Borrow ID", 1500
            lvList(0).ColumnHeaders.Add , , "Reg. BookID", 1500
            lvList(0).ColumnHeaders.Add , , "ISBN", 1400
            lvList(0).ColumnHeaders.Add , , "Title", 4000
            lvList(0).ColumnHeaders.Add , , "Borrow Date", 1200
            lvList(0).ColumnHeaders.Add , , "Return Date", 1150
            lvList(0).ColumnHeaders.Add , , "Date Returned", 1250
            lvList(0).ColumnHeaders.Add , , "Returned?", 1000
            lvList(0).ColumnHeaders.Add , , "Penalty Day(s)", 1250
            lvList(0).ForeColor = vbBlack
            Do While Not adoRes.EOF
                Set mRow = lvList(0).ListItems.Add(, , , , 2)
                mRow.SubItems(1) = adoRes.Fields("br_id")
                mRow.SubItems(2) = adoRes.Fields("rb_id")
                mRow.SubItems(3) = adoRes.Fields("isbn")
                mRow.SubItems(4) = adoRes.Fields("title")
                mRow.SubItems(5) = adoRes.Fields("b_date")
                mRow.SubItems(6) = adoRes.Fields("r_date")
                mRow.SubItems(7) = adoRes.Fields("datereturned")
                mRow.SubItems(8) = adoRes.Fields("s_return")
                mRow.SubItems(9) = adoRes.Fields("d_penalty")
                adoRes.MoveNext
            Loop
            txtTotal = Format(CalcTotalAmount, "0.00")
            txtCash.SetFocus
        Else
            lvList(0).ColumnHeaders.Add , , "", 8000
            lvList(0).ListItems.Add , , "No Current Book(s) Penalty Record Found.", , 1
            lvList(0).SelectedItem.ForeColor = vbRed
            txtTotal.Text = "0.00"
        End If
    'On Er GoTo errExit
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
    fraList(i).Move 120, 2880, Me.ScaleWidth - 240, Me.ScaleHeight - (2880 + fraRec.Height + 240)
    lvList(i).Move 120, 840, fraList(i).Width - 240, fraList(i).Height - (840 + 120)
    spMag(i).Move 0, 240, fraList(i).Width
    imgStud.Left = fraList(3).Width - imgStud.Width
    fraRec.Move Me.ScaleWidth - (120 + fraRec.Width), Me.ScaleHeight - (fraRec.Height + 180)
    imgBottom.Move 120, Me.ScaleHeight - (fraRec.Height + 180 - 120), Me.ScaleWidth - (fraRec.Width + 360)
End Sub

Public Function Return_Book(dDate As Date)
    Dim sWhere As String, sValues As String, sSQL As String, iPenaltyDay As Integer
    sSQL = "SELECT tbl_borrower_type.penaltystat " & _
        "From tbl_borrower_type  " & _
        "WHERE (((tbl_borrower_type.b_type) Like '" & lblBtype.Caption & "'))  " & _
        "GROUP BY tbl_borrower_type.penaltystat;"
    If FindField(sSQL, "penaltystat") = "1" Then
        iPenaltyDay = GetPenaltyDay(dDate)
    Else
        iPenaltyDay = 0
    End If
    sSQL = "SELECT tbl_borrow_record.rb_id " & _
        "From tbl_borrow_record " & _
        "WHERE (((tbl_borrow_record.br_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
        "GROUP BY tbl_borrow_record.rb_id;"
    sWhere = "br_id like '" & lvList(0).SelectedItem.SubItems(1) & "'"
    sValues = Date & ",1," & iPenaltyDay & "," & sUserId
    UPDATE_DATA "tbl_borrow_record", "datereturned,s_return,d_penalty,LastUserFK", sValues, sWhere, ",", True
    sWhere = "rb_id like '" & FindFieldValue(sSQL, "rb_id") & "'"
    UPDATE_DATA "tbl_reg_books", "borrow", "0", sWhere, ",", False
End Function

Public Function GetPenaltyDay(dDate As Date) As Integer
    If Date > dDate Then
        GetPenaltyDay = Date - dDate
    Else
        GetPenaltyDay = 0
    End If
End Function

Public Function CalcTotalAmount() As Double
    Dim i As Integer
    For i = 1 To lvList(0).ListItems.Count
        CalcTotalAmount = CalcTotalAmount + (lvList(0).ListItems(i).SubItems(9) * Val(lblFee.Caption))
    Next
End Function

Public Function CalcChange() As Double
    CalcChange = Val(txtCash.Text) - Val(txtTotal.Text)
End Function

Private Sub txtCash_Change()
    If Val(lvList(0).ColumnHeaders.Count) > 1 Then
        If Val(txtCash.Text) >= Val(txtTotal.Text) Then
            txtChange.Text = Format(CalcChange, "0.00")
            cmdPrint.Enabled = True
        Else
            cmdPrint.Enabled = False
        End If
    Else
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub txtCash_GotFocus()
    txtCash.Text = ""
End Sub

Private Sub txtCash_LostFocus()
    txtCash.Text = Format(Val(txtCash.Text), "0.00")
End Sub

Public Function SaveTransaction()
    Dim i As Integer, sSQL As String, sValues As String
    Dim sWhere As String
    sGeneratedID = GenerateID
    sValues = sGeneratedID & "," & txtCash.Text & "," & txtTotal.Text & "," & txtChange.Text & "," & Date _
                & "," & sUserId
    INSERT_DATA "tbl_penalty_payment", "brp_id,s_cash,s_total,s_change,p_date,validby", sValues, ",", True
    
    For i = 1 To lvList(0).ListItems.Count
        'INSERT RELATED RECORD 'TRANSACTION PENALTY PAYMENT' TO  'BORROW TRANSACTION'
        sValues = sGeneratedID & "," & lvList(0).ListItems(i).SubItems(1)
        INSERT_DATA "tbl_borrow_penalty", "brp_id,br_id", sValues, ",", False
        
        'SET BORROW TRANSATION SET STATUS TO PAID
        sWhere = "br_id like '" & lvList(0).ListItems(i).SubItems(1) & "'"
        UPDATE_DATA "tbl_borrow_record", "s_paid", "1", sWhere, ",", False
    Next
    CreateReceipt
    txtTotal.Text = "0.00"
    txtCash.Text = ""
    txtChange.Text = "0.00"
    cmdFind_Click
    cboBorrower.SetFocus
    cmdPrint.Enabled = False
End Function

Public Function CreateReceipt()
    Dim i As Integer
    Dim sSQL As String
    Dim sReceipt As String, iPenaltyDay As Integer
    On Error GoTo errHandler
    sReceipt = ""
    sReceipt = sReceipt & "TRINITY UNIVERSITY" & vbCrLf
    sReceipt = sReceipt & "QUEZON CITY" & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Transaction ID: " & sGeneratedID & vbCrLf
    sReceipt = sReceipt & "Current Date: " & Date & vbCrLf
    sReceipt = sReceipt & "Fee / Day: " & lblFee.Caption & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Reg. Book ID" & vbTab & "ISBN" & vbTab & vbTab & "Penalty" & vbCrLf
    sReceipt = sReceipt & "-------------------------------------------------------------" & vbCrLf
    For i = 1 To lvList(0).ListItems.Count
        sReceipt = sReceipt & lvList(0).ListItems(i).SubItems(2) & vbTab & lvList(0).ListItems(i).SubItems(3) & vbTab & lvList(0).ListItems(i).SubItems(9) & vbCrLf
    Next
    sReceipt = sReceipt & "-------------------------------------------------------------" & vbCrLf
    sReceipt = sReceipt & "TOTAL" & vbTab & vbTab & vbTab & txtTotal.Text & vbCrLf
    sReceipt = sReceipt & "CASH" & vbTab & vbTab & vbTab & txtCash.Text & vbCrLf
    sReceipt = sReceipt & "CHANGE" & vbTab & vbTab & vbTab & txtChange.Text & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Library Official Receipt" & vbCrLf
    sReceipt = sReceipt & "of Transaction" & vbCrLf
    MsgBox sReceipt
    Printer.Print sReceipt
errHandler:
    If err.Number = 487 Then
        MsgBox "Error Printing. No Printer Detected.", vbExclamation, "PrinterException"
    End If
End Function

Public Function PenaltFormat(sCurrency As String, iLength As Integer) As String
    Dim i As Integer, iAddLength As Integer
    iAddLength = iLength - Len(sCurrency)
    PenaltFormat = sCurrency
    For i = 1 To iAddLength
        PenaltFormat = " " & PenaltFormat
    Next
End Function

