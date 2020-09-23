VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSales_Rpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame fraDR 
      Height          =   2415
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdView 
         Caption         =   "Print"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   57540609
         CurrentDate     =   39494
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   57540609
         CurrentDate     =   39494
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   960
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmSales_Rpt.frx":0000
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view Sales Report according what you needed."
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
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   4005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Index           =   8
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Index           =   4
         Left            =   2400
         TabIndex        =   6
         Top             =   1560
         Width           =   180
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   3
         Left            =   4080
         Picture         =   "frmSales_Rpt.frx":27A2
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Select Date to view List of Sales  Report."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   9255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSales_Rpt.frx":891C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSales_Rpt.frx":1F2DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSales_Rpt.frx":25468
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSales_Rpt.frx":3C502
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSales_Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub cmdView_Click(Index As Integer)
    Dim rsView1 As ADODB.Recordset
    Dim sSQL As String
    On Error Resume Next
    If dtDate(0).Value <= dtDate(1).Value Then
        Set rsView1 = New ADODB.Recordset
        Set adoCon = New ADODB.Connection
        sSQL = "SELECT tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date " & _
            "FROM (tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN (tbl_borrowers INNER JOIN tbl_borrow_record ON tbl_borrowers.B_id = tbl_borrow_record.B_id) ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
            "WHERE (((tbl_borrow_record.s_return) Like '1') AND ((tbl_borrow_record.b_date) Between #" & dtDate(0).Value & "# And #" & dtDate(1).Value & "#)) " & _
            "GROUP BY tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date;"
        adoCon.Open sCon
        rsView1.Open sSQL, adoCon, 3, 3
        Set dtrSR.DataSource = rsView1
        sSQL = "SELECT Sum(tbl_orders.total) AS SumOftotal " & _
            "From tbl_orders " & _
            "WHERE (((tbl_orders.orderdate) Between #" & dtDate(0).Value & "# And #" & dtDate(1).Value & "#));"
        dtrSR.Sections("Section2").Controls("lblDate").Caption = "Date: " & dtDate(0).Value & " - " & dtDate(0).Value
        'Set adoCon = New ADODB.Connection
        'Set adoRes = New ADODB.Recordset
        Set adoCon = New ADODB.Connection
        Set adoRes = New ADODB.Recordset
        adoCon.Open sCon
        adoRes.Open sSQL, adoCon, 3, 3
            dtrSR.Sections("Section5").Controls("lblTotalSales").Caption = Val(adoRes.Fields("SumOftotal"))
        adoRes.Close
        adoCon.Close
        Set adoCon = Nothing
        Set adoRes = Nothing
        dtrSR.Show 1
    Else
        MsgBox "Date From must not greather than Date To.", vbExclamation, "DatePointerException"
    End If
End Sub

Private Sub Form_Load()
    dtDate(0).Value = Date
    dtDate(1).Value = Date
End Sub


