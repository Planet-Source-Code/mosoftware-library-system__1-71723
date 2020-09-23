VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overdue Books"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraS 
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
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8295
      Begin MSComctlLib.ListView lvList 
         Height          =   2295
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
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
      Left            =   6480
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Timer tmrBlnk 
      Interval        =   500
      Left            =   120
      Top             =   240
   End
   Begin MSComctlLib.ImageList imgTmr 
      Left            =   9120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopUp.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkStat 
      Caption         =   "On/Off Pop Up  Checker of Products in Critical Level."
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
      TabIndex        =   6
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   120
      Picture         =   "frmPopUp.frx":618A
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000AA27C&
      Caption         =   "List Of Over Due Books"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "List of Products has meet their Reorder Level or Critcal Level. Please contact autorize personnel according to this matter."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   7410
   End
   Begin VB.Image imgAct 
      Height          =   720
      Left            =   120
      Picture         =   "frmPopUp.frx":C304
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "For more information please ask an authorize personnel."
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
      Left            =   600
      TabIndex        =   3
      Top             =   4020
      Width           =   4050
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim intBlnkTmr As Integer

Private Sub cmdOk_Click()
    UPDATE_DATA2 "[tbl_popup]", "popupstat", "", chkStat.Value
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    SetListview
    InsertItem
    sSQL = "SELECT tbl_popup.popuptime, tbl_popup.popupstat " & _
        "FROM tbl_popup;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, 3, 3
        chkStat.Value = adoRes.Fields("popupstat")
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Sub

Public Function SetListview()
    lvList.ColumnHeaders.Clear
    lvList.ListItems.Clear
    Set lvList.SmallIcons = frmMain.iLv
    Set lvList.Icons = frmMain.iLv
    lvList.ColumnHeaders.Add , , "", 300
    lvList.ColumnHeaders.Add , , "Borrow ID", 1700
    lvList.ColumnHeaders.Add , , "Borrower ID", 1700
    lvList.ColumnHeaders.Add , , "Name", 2000
    lvList.ColumnHeaders.Add , , "ISBN", 2000
    lvList.ColumnHeaders.Add , , "Title", 3000
    lvList.ColumnHeaders.Add , , "Borrow Date", 1500
    lvList.ColumnHeaders.Add , , "Expected Return Date", 1500
End Function

Public Function InsertItem()
    Dim sSQL As String
    Dim mRow As ListItem
    'On Error Resume Next
    lvList.ListItems.Clear
    sSQL = "SELECT tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date " & _
        "FROM (tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN (tbl_borrowers INNER JOIN tbl_borrow_record ON tbl_borrowers.B_id = tbl_borrow_record.B_id) ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (tbl_borrow_record.s_return Like '0') AND (tbl_borrow_record.r_date<date()) " & _
        "GROUP BY tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        Do While Not adoRes.EOF
            Set mRow = lvList.ListItems.Add(, , , , 19)
            mRow.SubItems(1) = adoRes.Fields("br_id")
            mRow.SubItems(2) = adoRes.Fields("B_id")
            mRow.SubItems(3) = adoRes.Fields("fn") & " " & Right(adoRes.Fields("mn"), 1) & ". " & adoRes.Fields("ln")
            mRow.SubItems(4) = adoRes.Fields("ISBN")
            mRow.SubItems(5) = adoRes.Fields("Title")
            mRow.SubItems(6) = adoRes.Fields("b_date")
            mRow.SubItems(7) = adoRes.Fields("r_date")
            adoRes.MoveNext
        Loop
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub tmrBlnk_Timer()
    If intBlnkTmr Mod 2 Then
        imgAct.Picture = imgTmr.ListImages(1).ExtractIcon
    Else
        Set imgAct.Picture = Nothing
    End If
    intBlnkTmr = intBlnkTmr + 1
End Sub
