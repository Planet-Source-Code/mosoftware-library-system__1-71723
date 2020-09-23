VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBook_AE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save"
      Top             =   5780
      Width           =   315
   End
   Begin VB.PictureBox picBtn 
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   25
      Top             =   5760
      Width           =   1335
      Begin VB.CommandButton btnFirst 
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "First"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton btnPrev 
         Enabled         =   0   'False
         Height          =   315
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Previous"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton btnNext 
         Height          =   315
         Left            =   645
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Next"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton btnLast 
         Height          =   315
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Last"
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdMod 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5780
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancel"
      Top             =   5780
      Width           =   315
   End
   Begin VB.Frame fraList 
      Height          =   5655
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   1
         Left            =   4395
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   1920
         Width           =   300
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   4575
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   3600
         Index           =   0
         Left            =   1560
         TabIndex        =   17
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6350
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   5
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
      Begin VB.CommandButton cmdRmv 
         Height          =   285
         Left            =   4395
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2205
         Width           =   300
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Unleash the Information Technology on School. Faster and much more easier to use."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1035
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   1155
         Left            =   120
         Picture         =   "frmBook_AE.frx":0000
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Author(s)"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Insert Author(s) of this book then click &Save button."
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
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   840
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   1
         Left            =   720
         Picture         =   "frmBook_AE.frx":761D
         Top             =   240
         Width           =   720
      End
      Begin VB.Image imgExc 
         Height          =   600
         Index           =   1
         Left            =   4080
         Top             =   840
         Width           =   600
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame fraList 
      Height          =   5655
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdCheck 
         Height          =   285
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   300
      End
      Begin MSComCtl2.DTPicker txtYr 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   3840
         Width           =   2895
         _ExtentX        =   5106
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
         CalendarBackColor=   16777215
         CustomFormat    =   "yyyy"
         Format          =   19791875
         CurrentDate     =   39456
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   4575
      End
      Begin VB.CommandButton cmdGet 
         Height          =   285
         Index           =   0
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Select"
         Top             =   3000
         Width           =   300
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   6
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4200
         Width           =   2865
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
         Index           =   5
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3000
         Width           =   2865
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H80000018&
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
         Left            =   2595
         MaxLength       =   1
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H80000018&
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
         Left            =   2355
         MaxLength       =   1
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H80000018&
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
         MaxLength       =   1
         TabIndex        =   0
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H80000018&
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
         Left            =   1560
         MaxLength       =   7
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1680
         Width           =   770
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
         Index           =   4
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2160
         Width           =   2865
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   2880
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "DESCRIPTION"
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
         TabIndex        =   34
         Top             =   4245
         Width           =   1260
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "YEAR PUBLISH"
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
         TabIndex        =   33
         Top             =   3870
         Width           =   1290
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "PUBLISHER"
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
         TabIndex        =   32
         Top             =   3045
         Width           =   1290
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "TITLE"
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
         TabIndex        =   31
         Top             =   2160
         Width           =   1275
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
         Picture         =   "frmBook_AE.frx":1DFCF
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H00808080&
         Caption         =   "ISBN"
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
         TabIndex        =   23
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You have to fill the following information then click Next Button or Save button to finish Adding New Book."
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
         Height          =   615
         Index           =   0
         Left            =   1320
         TabIndex        =   22
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   2760
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
Attribute VB_Name = "frmBook_AE"
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

Private Sub btnNext_Click()
    If isFill(1) = True Then
        ifraIndex = ifraIndex + 1
        If ifraIndex >= 0 And ifraIndex <= 1 Then
            fraList(ifraIndex - 1).Visible = False
            fraList(ifraIndex).Visible = True
            If ifraIndex = 1 Then
                btnLast.Enabled = False
                btnNext.Enabled = False
                
                btnPrev.Enabled = True
                btnFirst.Enabled = True
            End If
        Else
            ifraIndex = 1
        End If
    End If
End Sub

Private Sub btnPrev_Click()
    ifraIndex = ifraIndex - 1
    If ifraIndex >= 0 And ifraIndex <= 1 Then
        fraList(ifraIndex + 1).Visible = False
        fraList(ifraIndex).Visible = True
        If ifraIndex = 0 Then
            btnPrev.Enabled = False
            btnFirst.Enabled = False
            
            btnLast.Enabled = True
            btnNext.Enabled = True
        End If
    Else
        ifraIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim sChkSql As String
    Dim sISBN As String
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.rb_id) Like '" & txtInput(1).Text & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    If isFillISBN = True Then
        If isRecordExist(sChkSql) = True Then
            imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
        Else
            imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
        End If
    End If
End Sub

Private Sub cmdGet_Click(Index As Integer)
    Dim newFormSelect As New frmSelect
    Dim gSQL As String, gLblHead As String, gLblDef As String
    Dim gXY As String, gTitle As String, gColumns As String
    Dim gColWidth As String, gFields As String, gLoop As Integer
    Dim gIcon As Integer, gLvIcon As Integer, sNoRec As String
    Dim mRow As ListItem
    
    If Index = 0 Then
        gSQL = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny " & _
            "FROM tbl_publishers;"
        gIcon = 4
        gLblHead = "Book Publishers"
        gLblDef = "Choose Book Publisher then click Select button."
        gXY = "10020,4785"
        gTitle = "Select Publisher"
        gColumns = "Publisher ID,Company Name"
        gColWidth = "1154,4770"
        gFields = "pubid,cmpny"
        gLoop = 1
        gLvIcon = 5
        sNoRec = "No Current Publisher Info ."
    ElseIf Index = 1 Then
        gSQL = "SELECT tbl_authors.auid, tbl_authors.author, " & _
            "tbl_authors.yrborn FROM tbl_authors;"
        gIcon = 3
        gLblHead = "Book Author(s)"
        gLblDef = "Choose Book Author(s) then click Select button."
        gXY = "8010,4785"
        gTitle = "Select Author"
        gColumns = "Author ID,Author Name,Year Born"
        gColWidth = "1154,2954,1484"
        gFields = "auid,author,yrborn"
        gLoop = 2
        gLvIcon = 4
        sNoRec = "No Current Author(s) Info ."
    End If
    
    SelectedItem = SelectItem(newFormSelect, gSQL, gIcon, gLblHead, _
                gLblDef, gXY, gTitle, gColumns, gLoop, gColWidth, _
                    gLvIcon, gFields, sNoRec)
    
    If Index = 0 And Len(SelectedItem) > 0 Then
        sSQL = "SELECT tbl_publishers.cmpny " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.pubid) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_publishers.cmpny;"
        txtInput(5).Text = SelectedItem & "-" & FindField(sSQL, "cmpny")
    ElseIf Index = 1 Then
        sSQL = "SELECT tbl_authors.author " & _
            "From tbl_authors " & _
            "WHERE (((tbl_authors.auid) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_authors.author;"
        If isOnList(SelectedItem & "-" & FindField(sSQL, "author")) = False Or lvList(0).ListItems.Count = 0 Then
            Set mRow = lvList(0).ListItems.Add(, , , , 4)
            mRow.SubItems(1) = SelectedItem & "-" & FindField(sSQL, "author")
        Else
            MsgBox "You already add this Author.", vbExclamation, "Record Exist"
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
    sSQL = "SELECT tbl_books.DateAdded, tbl_books.AddedByFK, tbl_books.DateModified, tbl_books.LastUserFK " & _
            "From tbl_books " & _
            "WHERE (((tbl_books.isbn) Like '" & frmBook.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_books.DateAdded, tbl_books.AddedByFK, tbl_books.DateModified, tbl_books.LastUserFK;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    'txtInput(6).Text = sSql
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        tDate1 = adoRes.Fields("DateAdded")
        tUser1 = adoRes.Fields("AddedByFK")
        tDate2 = adoRes.Fields("DateModified")
        tUser2 = adoRes.Fields("LastUserFK")
        'MsgBox tUser2
    adoRes.Close
    adoCon.Close
    Set adoRes = Nothing
    Set adoCon = Nothing
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

Private Sub cmdRmv_Click()
    If lvList(0).ListItems.Count > 0 Then
        lvList(0).ListItems.Remove lvList(0).SelectedItem.Index
    End If
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
    Dim vPub As Variant
    Dim sChkSql As String
    Dim sWhere As String
    Dim vAuthor As Variant
    Dim i As Integer
    sChkSql = "SELECT tbl_books.isbn " & _
            "From tbl_books " & _
            "WHERE (((tbl_books.isbn) Like '" & getISBN & "')) " & _
            "GROUP BY tbl_books.isbn;"
    If isFill(1) = True Then
            vPub = Split(txtInput(5).Text, "-")
            'sWhere = "isbn like '" & frmBook.lvList(0).SelectedItem.SubItems(1) & "'"
            sValues = getISBN & "," & txtInput(4).Text & "," & vPub(0) & "," & txtYr.Year & "," & txtInput(6).Text _
                & "," & Date & "," & sUserId
            If isRecordExist(sChkSql) = False Then
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                INSERT_DATA "tbl_books", "isbn,title,pubid,yrpub,[desc],DateAdded,AddedByFK", sValues, ",", True
                frmBook.btnNew_Load 0
                LvSearchItem frmBook.lvList(0), getISBN
                sSQL = "SELECT tbl_books.isbn " & _
                    "From tbl_books " & _
                    "WHERE (((tbl_books.isbn) Like '" & getISBN & "')) " & _
                    "GROUP BY tbl_books.isbn;"
                If isRecordIfExist(sSQL, 5) = True Then
                    For i = 1 To lvList(0).ListItems.Count
                        vAuthor = Split(lvList(0).ListItems(i).SubItems(1), "-")
                        sValues = getISBN & "," & vAuthor(0)
                        INSERT_DATA "tbl_bookauthor", "isbn,auid", sValues, ",", False
                    Next
                Else
                    MsgBox "Book Insert not found.", vbExclamation, "isRecordExist"
                End If
                Unload Me
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "ISBN already exist. Please change it!", vbExclamation
                'txtInput(0).SetFocus
            End If
        End If
End Function

Public Function ModifyData()
    Dim sValues As String
    Dim vPub As Variant
    Dim sChkSql As String
    Dim sWhere As String
    Dim vAuthor As Variant
    sChkSql = "SELECT tbl_books.isbn " & _
            "From tbl_books " & _
            "WHERE (((tbl_books.isbn) Like '" & getISBN & "')) " & _
            "GROUP BY tbl_books.isbn;"
    If isFill(1) = True Then
            vPub = Split(txtInput(5).Text, "-")
            sWhere = "isbn like '" & frmBook.lvList(0).SelectedItem.SubItems(1) & "'"
            If isRecordExist(sChkSql) = False Or getISBN = frmBook.lvList(0).SelectedItem.SubItems(1) Then
                sChkSql = "SELECT tbl_reg_books.rb_id " & _
                    "From tbl_reg_books " & _
                    "WHERE (((tbl_reg_books.isbn) Like '" & getISBN & "')) " & _
                    "GROUP BY tbl_reg_books.rb_id;"
                imgChk(0).Picture = frmMain.i16x16.ListImages(12).ExtractIcon
                If isRecordExist(sChkSql) = True Then
                    sValues = txtInput(4).Text & "," & vPub(0) & "," & txtYr.Year & "," & txtInput(6).Text _
                        & "," & Date & "," & sUserId
                    UPDATE_DATA "tbl_books", "title,pubid,yrpub,[desc],DateModified,LastUserFK", sValues, sWhere, ",", True
                    With frmBook.lvList(0)
                        .SelectedItem.SubItems(1) = getISBN
                        .SelectedItem.SubItems(2) = txtInput(4).Text
                        .SelectedItem.SubItems(3) = vPub(1)
                        .SelectedItem.SubItems(4) = txtYr.Year
                        .SelectedItem.SubItems(5) = txtInput(6).Text
                    End With
                    Unload Me
                Else
                    sValues = getISBN & "," & txtInput(4).Text & "," & vPub(0) & "," & txtYr.Year & "," & txtInput(6).Text _
                        & "," & Date & "," & sUserId
                    UPDATE_DATA "tbl_books", "isbn,title,pubid,yrpub,[desc],DateModified,LastUserFK", sValues, sWhere, ",", True
                    With frmBook.lvList(0)
                        '.selecteditem.SubItems(1) = getISBN
                        .SelectedItem.SubItems(2) = txtInput(4).Text
                        .SelectedItem.SubItems(3) = vPub(1)
                        .SelectedItem.SubItems(4) = txtYr.Year
                        .SelectedItem.SubItems(5) = txtInput(6).Text
                    End With
                    Unload Me
                End If
            Else
                imgChk(0).Picture = frmMain.i16x16.ListImages(7).ExtractIcon
                MsgBox "ISBN already exist. Please change it!", vbExclamation
                txtInput(0).SetFocus
            End If
        End If
End Function

Private Sub Form_Load()
    Dim vID As Variant, i As Integer, sChkSql As String
    On Error Resume Next
    sChkSql = "SELECT tbl_reg_books.rb_id " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.isbn) Like '" & frmBook.lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
    Set lvList(0).SmallIcons = frmMain.iLv
    Set lvList(0).Icons = frmMain.iLv
    lvList(0).ColumnHeaders.Add , , , 300
    lvList(0).ColumnHeaders.Add , , "Author(s)", lvList(0).Width
    ifraIndex = 0
    If bStat = True Then 'CREATE NEW
        cmdMod.Visible = False
        Me.Caption = "Create New Entry"
        txtYr.Year = Year(Date)
    Else 'MODIFY DATA
        With frmBook
            'bSFocus = False
            vID = Split(.lvList(0).SelectedItem.SubItems(1), "-")
            txtInput(0).Text = vID(0)
            txtInput(1).Text = vID(1)
            txtInput(2).Text = vID(2)
            txtInput(3).Text = vID(3)
            txtInput(4).Text = .lvList(0).SelectedItem.SubItems(2)
            sSQL = "SELECT tbl_publishers.pubid " & _
                "From tbl_publishers " & _
                "WHERE (((tbl_publishers.cmpny) Like '" & .lvList(0).SelectedItem.SubItems(3) & "')) " & _
                "GROUP BY tbl_publishers.pubid;"
            txtInput(5).Text = FindField(sSQL, "pubid") & "-" & .lvList(0).SelectedItem.SubItems(3)
            txtYr.Year = .lvList(0).SelectedItem.SubItems(4)
            txtInput(6).Text = .lvList(0).SelectedItem.SubItems(5)
            'bSFocus = True
        End With
        'MsgBox sChkSql
        If isRecordExist(sChkSql) = True Then
            For i = 0 To 3
                txtInput(i).Locked = True
            Next
        End If
        Me.Caption = "Edit Existing"
        btnNext.Enabled = False
        btnLast.Enabled = False
        picBtn.Enabled = False
    End If
    SetButtonPicture
    'ActivateXPTheme wXP
End Sub

Private Sub lvList_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim CURR_COL As Integer
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lvList(Index).SortOrder = 0
    Else
        lvList(Index).SortOrder = Abs(lvList(Index).SortOrder - 1)
    End If
    lvList(Index).SortKey = ColumnHeader.Index - 1
    
    lvList(Index).Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub txtInput_Change(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        If Len(txtInput(Index).Text) > 0 Then
            txtInput(1).SetFocus
        End If
    ElseIf Index = 1 Then
        If Len(txtInput(Index).Text) > 6 Then
            txtInput(2).SetFocus
        End If
    ElseIf Index = 2 Then
        If Len(txtInput(Index).Text) > 0 Then
            txtInput(3).SetFocus
        End If
    End If
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    HLText txtInput(Index)
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index >= 0 And Index <= 3 Then KeyAscii = isNumAndChar(KeyAscii)
    If KeyAscii = 8 Then
            If Index = 1 Then
                If Len(txtInput(Index).Text) = 0 Then
                    txtInput(0).SetFocus
                End If
            ElseIf Index = 2 Then
                If Len(txtInput(Index).Text) = 0 Then
                    txtInput(1).SetFocus
                End If
            ElseIf Index = 3 Then
                If Len(txtInput(Index).Text) = 0 Then
                    txtInput(2).SetFocus
            End If
        End If
    End If
    If Index = 5 Then
        MsgBox "Unabled to Input Publisher. You have to click the button in the upper right corner of the textbox to Insert a Publisher." _
        , vbExclamation, "Unabled Input"
    End If
End Sub

                                    ''''''''''''''''''''''''''''''
                                    'List of New Function Created'
                                    ''''''''''''''''''''''''''''''
                                    
'Check if is String is On the List of the Listview
Private Function isOnList(sAuthor As String) As Boolean
    Dim i As Integer
    isOnList = False
    For i = 1 To lvList(0).ListItems.Count
        If sAuthor = lvList(0).ListItems(i).Text Then
            isOnList = True
            Exit For
        End If
    Next
End Function

'Set List of Buttons Picture
Private Sub SetButtonPicture()
    Dim i As Integer
    With frmMain
        btnFirst.Picture = .iPageEnabled.ListImages(1).ExtractIcon
        btnFirst.DisabledPicture = .iPageDisabled.ListImages(1).ExtractIcon
        btnPrev.Picture = .iPageEnabled.ListImages(2).ExtractIcon
        btnPrev.DisabledPicture = .iPageDisabled.ListImages(2).ExtractIcon
        btnNext.Picture = .iPageEnabled.ListImages(3).ExtractIcon
        btnNext.DisabledPicture = .iPageDisabled.ListImages(3).ExtractIcon
        btnLast.Picture = .iPageEnabled.ListImages(4).ExtractIcon
        btnLast.DisabledPicture = .iPageDisabled.ListImages(4).ExtractIcon
        cmdSave.Picture = .i16x16.ListImages(9).ExtractIcon
        cmdCancel.Picture = .i16x16.ListImages(7).ExtractIcon
        cmdMod.Picture = .i16x16.ListImages(10).ExtractIcon
        
        imgExc(0).Picture = .iHead.ListImages(2).ExtractIcon
        cmdGet(0).Picture = .i16x16.ListImages(11).ExtractIcon
        
        imgExc(1).Picture = .iHead.ListImages(2).ExtractIcon
        cmdGet(1).Picture = .i16x16.ListImages(2).ExtractIcon
        cmdRmv.Picture = .i16x16.ListImages(4).ExtractIcon
        cmdCheck.Picture = .i16x16.ListImages(12).ExtractIcon
    End With
End Sub
Private Function SetColor(colorCons As ColorConstants) As ColorConstants
    SetColor = colorCons
End Function

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
        For i = 0 To 6
            If isNull(txtInput(i).Text) = True Then
               iNotFill = iNotFill + 1
            End If
            If iNotFill = 1 Then
                iStart = i
                Exit For
            End If
        Next
        If Not Len(txtInput(0).Text) = txtInput(0).MaxLength Then
            FillMsgBox "ISBN first layer"
        End If
        If Not Len(txtInput(1).Text) = txtInput(1).MaxLength Then
            FillMsgBox "ISBN Second layer"
        End If
        If Not Len(txtInput(2).Text) = txtInput(2).MaxLength Then
            FillMsgBox "ISBN Third layer"
        End If
        If Not Len(txtInput(3).Text) = txtInput(3).MaxLength Then
            FillMsgBox "ISBN fourth layer"
        End If
        If isNull(txtInput(4).Text) = True Then
            FillMsgBox "Title"
        End If
        If isNull(txtInput(5).Text) = True Then
            FillMsgBox "Publisher"
        End If
        'If isNull(txtInput(6).Text) = True Then
        '    FillMsgBox "Description"
        'End If
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

Private Function isFillISBN() As Boolean
    Dim sBeginMsg As String
    Dim sEndMsg As String
    Dim iNotFill As Integer
    Dim i As Integer
    Dim iStart As Integer
    Dim iSplitItem As Integer
    Dim sSplitedItem As Variant
    iNotFill = 0
    sMsgBox = ""
    
    For i = 0 To 6
        If isNull(txtInput(i).Text) = True Then
            iNotFill = iNotFill + 1
        End If
        If iNotFill = 1 Then
            iStart = i
        End If
    Next
    
    sBeginMsg = "You forgot to fill the following "
    sEndMsg = "."
    If Not Len(txtInput(0).Text) = txtInput(0).MaxLength Then
        FillMsgBox "ISBN first layer"
    End If
    If Not Len(txtInput(1).Text) = txtInput(1).MaxLength Then
        FillMsgBox "ISBN Second layer"
    End If
    If Not Len(txtInput(2).Text) = txtInput(2).MaxLength Then
        FillMsgBox "ISBN Third layer"
    End If
    If Not Len(txtInput(3).Text) = txtInput(3).MaxLength Then
        FillMsgBox "ISBN fourth layer"
    End If
    iSplitItem = CountSplitItem(sMsgBox, ",")
    sSplitedItem = Split(sMsgBox, ",")
    If Len(sMsgBox) = 0 Then isFillISBN = True Else isFillISBN = False
    If isFillISBN = False Then
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
End Function

Public Function getISBN() As String
getISBN = txtInput(0).Text & "-" & txtInput(1).Text & "-" & txtInput(2).Text & "-" & txtInput(3).Text
End Function

