VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPublisher 
   Caption         =   "Publishers"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11295
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   9255
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   1
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   43
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   46
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   1
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   47
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  ToolTipText     =   "Create New"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnES 
               Height          =   375
               Index           =   1
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   49
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   50
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   1
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   51
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   52
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnD 
               Height          =   375
               Index           =   1
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   53
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnR 
               Height          =   375
               Index           =   1
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   55
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   1
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   57
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   1
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   44
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   1
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   1
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   33
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   1
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Previous 250"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   1
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   1
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   1
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   1
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   1
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
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
            Index           =   1
            Left            =   -480
            TabIndex        =   41
            Top             =   45
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1931
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
         Caption         =   "You can view their books that they've been publish."
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
         Index           =   1
         Left            =   960
         TabIndex        =   61
         Top             =   480
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books been Published"
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
         Index           =   1
         Left            =   960
         TabIndex        =   60
         Top             =   240
         Width           =   2685
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   240
         Picture         =   "frmPublisher.frx":0000
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Record:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   19
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   0
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   0
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   0
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   0
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   0
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   0
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   0
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
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
            Index           =   0
            Left            =   -120
            TabIndex        =   27
            Top             =   45
            Width           =   2055
         End
      End
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   3
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   6
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   7
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  ToolTipText     =   "Create New"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnES 
               Height          =   375
               Index           =   0
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   9
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   0
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   11
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnD 
               Height          =   375
               Index           =   0
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   13
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnR 
               Height          =   375
               Index           =   0
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   15
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   16
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   0
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   17
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   0
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   4
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1905
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "frmPublisher.frx":27A2
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher Information"
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
         TabIndex        =   31
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view book information and change the information to become more accurate the system."
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
         TabIndex        =   30
         Top             =   480
         Width           =   6990
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Record:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "View List of Published Books"
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
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   2535
   End
   Begin VB.PictureBox picConn 
      Height          =   135
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   3120
      Width           =   9255
   End
End
Attribute VB_Name = "frmPublisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim view_other As Boolean
Dim INT_SIZE As Integer
Dim int_size_active As Integer
Dim CURR_COL As Integer
Dim iStartPage(2) As Long
Dim iNoPage(2) As Integer
Dim iRec(2) As Long
Dim sSQL(2) As String

Dim sColumns(2) As String, sColWidth(2) As String, sFields(2) As String
Dim iIcon(2) As Integer, iLoop(2) As Integer, sNoRec(2) As String
Public iLvIndex As Integer
Public sSearchFields As String

Private Sub btnAll_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, " & _
                "tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst " & _
                    "FROM tbl_publishers;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Publishers Info Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
                "From tbl_books " & _
                "WHERE (((tbl_books.pubid) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Books Published Records."
    End If
End Sub

Private Sub btnC_Click(Index As Integer)
    iLvIndex = Index
    LvClose Index
End Sub

Private Sub btnCN_Click(Index As Integer)
    iLvIndex = Index
    LvNew Index
End Sub

Private Sub btnD_Click(Index As Integer)
    iLvIndex = Index
    If iRec(Index) > 0 Then
        LvDelete Index
    End If
End Sub

Private Sub btnES_Click(Index As Integer)
    iLvIndex = Index
    LvEdit Index
End Sub

Private Sub btnFirst_Click(Index As Integer)
    If Index = 0 Then
        iStartPage(Index) = 1
        LvPageStat Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0)
    ElseIf Index = 1 Then
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    End If
End Sub

Private Sub btnLast_Click(Index As Integer)
    Dim iLastNoPage As Long
    iLastNoPage = 1
    If Index = 0 Then
        Do While iLastNoPage <= iRec(Index)
            iLastNoPage = iLastNoPage + iNoPage(Index)
        Loop
        iStartPage(Index) = iLastNoPage - iNoPage(Index)
        LvPageStat Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0)
    ElseIf Index = 1 Then
        Do While iLastNoPage <= iRec(Index)
            iLastNoPage = iLastNoPage + iNoPage(Index)
        Loop
        iStartPage(Index) = iLastNoPage - iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    End If
End Sub

Private Sub btnNew_Click(Index As Integer)
    iLvIndex = Index
    btnNew_Load Index
End Sub

Private Sub btnNext_Click(Index As Integer)
    If Index = 0 Then
        iStartPage(Index) = iStartPage(Index) + iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    ElseIf Index = 1 Then
        iStartPage(Index) = iStartPage(Index) + iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    End If
End Sub

Private Sub btnP_Click(Index As Integer)
    PRINT_RECORD Index
End Sub

Public Function PRINT_RECORD(Index As Integer)
    Dim rsView1 As ADODB.Recordset
    
    Set rsView1 = New ADODB.Recordset
    Set adoCon = New ADODB.Connection
    adoCon.Open sCon
    
    If Index = 0 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        Set dtrPub.DataSource = rsView1
        dtrPub.Show 1
    ElseIf Index = 1 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrPubBook.Sections("Section2")
            .Controls("lblPubid").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblComp").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblC").Caption = lvList(0).SelectedItem.SubItems(2)
            .Controls("lblAdd").Caption = lvList(0).SelectedItem.SubItems(3)
        End With
        Set dtrPubBook.DataSource = rsView1
        dtrPubBook.Show 1
    End If
End Function

Private Sub btnPrev_Click(Index As Integer)
    If Index = 0 Then
        iStartPage(Index) = iStartPage(Index) - iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    ElseIf Index = 1 Then
        iStartPage(Index) = iStartPage(Index) - iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    End If
End Sub

Private Sub btnR_Click(Index As Integer)
    iLvIndex = Index
    LvRefresh Index
End Sub

Private Sub btnS_Click(Index As Integer)
    iLvIndex = Index
    PopupMenu frmMain.mnuFS
End Sub

Private Sub cmdReg_Click()
    If view_other = True And lvList(1).Visible = True Then
        fraList(1).Visible = False
        picConn.Visible = False
        view_other = False
    Else
        sSQL(1) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
                "From tbl_books  " & _
                "WHERE (((tbl_books.pubid) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc;"
        Lv_OtherInfo
        fraList(1).Visible = True
        picConn.Visible = True
        view_other = True
    End If
    INT_SIZE = Me.ScaleHeight / 2
    Listview_Resize
End Sub

Private Sub Form_Load()
    view_other = False
    INT_SIZE = Me.ScaleHeight / 2
    Set_Icon_btn Me, 1
    fraList(1).Visible = False
    picConn.Visible = False
    sSQL(0) = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, " & _
            "tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst " & _
            "FROM tbl_publishers;"
    Lv_MainInfo
    lvList(0).Refresh
    frmMain.TabMainIni 2, "Publishers", 5
    ListOfDisabledBtn
End Sub

Public Function ListOfDisabledBtn()
    btnES(1).Visible = False
    btnCN(1).Visible = False
    btnES(1).Visible = False
    btnD(1).Visible = False
    btnNew(1).Visible = True
    btnEdited(1).Visible = True
    btnAll(1).Visible = True
End Function

Private Sub Form_LostFocus()
    If frmFind.Visible = True Then
        frmFind.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call Listview_Resize
    cmdReg.Top = Me.ScaleHeight - (cmdReg.Height + 120)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.TabMainIni 1, "Publishers", 5
End Sub

Private Sub lvList_Click(Index As Integer)
    iLvIndex = Index
End Sub

Private Sub lvList_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lvList(Index).SortOrder = 0
    Else
        lvList(Index).SortOrder = Abs(lvList(Index).SortOrder - 1)
    End If
    lvList(Index).SortKey = ColumnHeader.Index - 1
    
    lvList(Index).Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub lvList_DblClick(Index As Integer)
    If Index = 0 Then
        If iRec(Index) > 0 Then
            frmPublisher_AE.bStat = False
            Set frmPublisher_AE.fCur = Me
            frmPublisher_AE.Show 1
        End If
    End If
End Sub

Private Sub lvList_GotFocus(Index As Integer)
    iLvIndex = Index
End Sub

Private Sub lvList_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If iRec(Index) > 0 Then
        lblSelected(Index).Caption = "Selected Record: " & (iStartPage(Index) - 1) + lvList(Index).SelectedItem.Index
    Else
        lblSelected(Index).Caption = "Selected Record: None"
    End If
    If Index = 0 Then
        If lvList(1).Visible = True Then
            LvRefresh 1
        End If
    End If
End Sub

Private Sub lvList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.iLvIndex = Index
    If Button = 2 Then PopupMenu frmMain.mnuAct
End Sub

Private Sub picConn_DblClick()
    INT_SIZE = Me.ScaleHeight / 2
    Listview_Resize
End Sub

Private Sub picConn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    int_size_active = 1
End Sub

Private Sub picConn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If int_size_active = 1 Then
        If Y < 0 Then
            If picConn.Top > (fraList(0).Top + 2600) Then
                picConn.Top = picConn.Top - (-(Y))
            End If
        Else
            If fraList(1).Height >= 2600 Then
                picConn.Top = picConn.Top + Y
            End If
        End If
        INT_SIZE = picConn.Top
        Listview_Resize
    End If
End Sub

Private Sub picConn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    int_size_active = 0
End Sub

                                    ''''''''''''''''''''''''''''''''''''''
                                    'List of New Function\Methods Created'
                                    ''''''''''''''''''''''''''''''''''''''
Public Function FindText()
    If iRec(iLvIndex) > 0 Then
        frmFind.Refresh_Values lvList(iLvIndex)
    End If
End Function

'This function will resize the Frames
Private Sub Listview_Resize()
    On Error Resume Next
    Dim i As Integer
    If view_other = False Then
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, Me.ScaleHeight - (cmdReg.Height + 180)
        lvList(i).Move 120, 1320, fraList(i).Width - 240, fraList(i).Height - (1320 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
        lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
    Else
        picConn.Width = Me.ScaleWidth - 240
        picConn.Top = INT_SIZE
        'this will use for listview(0)
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, picConn.Top - (105)
        lvList(i).Move 120, 1320, fraList(i).Width - 240, fraList(i).Height - (1320 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 120)
        lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
        'this will use for listview(1)
        i = 1
        fraList(i).Move 120, (picConn.Top + picConn.Height) - 15, Me.ScaleWidth - 240, (Me.ScaleHeight - ((picConn.Top - 15) + 240 + 240 + 120 + 80))
        lvList(i).Move 120, 1320, fraList(i).Width - 240, fraList(i).Height - (1320 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
        'lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
    End If
End Sub

'This Function is use to Refresh the Info on Listview(0)
Public Sub Lv_MainInfo()
    sColWidth(0) = "1650,3990,2369,1235,1964,1000,1000"
    sColumns(0) = "Publisher ID,Company Name,Country,Address,Telephone,Fax,Website"
    iIcon(0) = 5
    iLoop(0) = 6
    sFields(0) = "pubid,cmpny,cntry,add,tel,fx,wbst"
    sNoRec(0) = "No Current Publisher Info Records."
    iStartPage(0) = 1
    iNoPage(0) = 75
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0))
    If iRec(0) > 0 Then
        cmdReg.Enabled = True
        Lv_OtherInfo
    Else
        view_other = False
        picConn.Visible = False
        Listview_Resize
        cmdReg.Enabled = False
    End If
End Sub

'This Function is used to Refresh the Info on Listview(1) Index
Public Sub Lv_OtherInfo()
    sColWidth(1) = "1650,7000,1235,2333"
    sColumns(1) = "ISBN,Title,Year Publish,Description"
    iIcon(1) = 2
    iLoop(1) = 3
    sFields(1) = "isbn,title,yrpub,desc"
    sNoRec(1) = "No Current Book Info Records."
    iStartPage(1) = 1
    iNoPage(1) = 75
    iRec(1) = LvPageStat(Me, 1, sSQL(1), iStartPage(1), iNoPage(1), iIcon(1), sColumns(1), iLoop(1), sColWidth(1), sFields(1), sNoRec(1))
    lvList(1).Refresh
End Sub

'This will be used to Control the Update of Information
Private Sub btnEdited_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.DateModified) between #" & Date & "# AND #" & Date & "#)) " & _
            "GROUP BY tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Publishers Updated Info Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
                "From tbl_books " & _
                "WHERE (((tbl_books.DateModified) Between #" & Date - 7 & "# And #" & Date & "#) AND ((tbl_books.pubid) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Books Updated Books."
    End If
End Sub

'This will used to Get New Create Item
Public Function btnNew_Load(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst " & _
            "From tbl_publishers " & _
            "WHERE (((tbl_publishers.DateAdded) between #" & Date & "# AND #" & Date & "#)) " & _
            "GROUP BY tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), "No Current New Publishers Added Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
        End If
        If lvList(1).Visible = True Then
            LvRefresh 1
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
                "From tbl_books " & _
                "WHERE (((tbl_books.DateAdded) Between #" & Date - 7 & "# And #" & Date & "#) AND ((tbl_books.pubid) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current New Books Added Records."
    End If
End Function

'This will be used to Where to Create New Item
Public Function LvNew(Index As Integer)
    If Index = 0 Then
        frmPublisher_AE.bStat = True
        Set frmPublisher_AE.fCur = Me
        frmPublisher_AE.Show 1
    End If
End Function

Private Function isOnList(Index As Integer, sAuthor As String) As Boolean
    Dim i As Integer
    isOnList = False
    For i = 1 To lvList(Index).ListItems.Count
        If sAuthor = lvList(Index).ListItems(i).Text Then
            isOnList = True
            Exit For
        End If
    Next
End Function

'This will used Where to Delete Records
Public Function LvDelete(Index As Integer)
    On Error GoTo errHandler
    Dim i
    Dim sMsgDel As String
    Dim sMsgFooter As String
    Dim sMsgId As String
    If iRec(Index) > 0 Then
        If Index = 0 Then
            sMsgDel = "You are about to delete this record?"
            sMsgId = "Publisher ID"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_publishers", "pubid", lvList(Index).SelectedItem.SubItems(1)
                lvList(Index).ListItems.Remove lvList(Index).SelectedItem.Index
                If lvList(Index).ListItems.Count = 0 Then
                    Call Lv_MainInfo
                End If
                If lvList(1).Visible = True Then
                    LvRefresh 1
                End If
                lblPageInfo(Index).Caption = iStartPage(Index) & " - " & (iStartPage(Index) + (iNoPage(Index) - 2)) & " of " & iRec(Index) - 1
            End If
        End If
    End If
errHandler:
    If Not err.Number = 0 Then
        MsgBox err.Number & Chr$(13) & Chr$(13) & " " & err.Description, vbExclamation, "Delete"
    End If
End Function

'This will used where to be update
Public Function LvEdit(Index As Integer)
    If Index = 0 And iRec(Index) > 0 Then
        frmPublisher_AE.bStat = False
        Set frmPublisher_AE.fCur = Me
        frmPublisher_AE.Show 1
    End If
End Function

'This will be use to CLosed the Records
Public Function LvClose(Index As Integer)
    If Index = 0 Then
        Unload Me
    ElseIf Index = 1 Then
        fraList(Index).Visible = False
        picConn.Visible = False
        view_other = False
        Listview_Resize
    End If
End Function

Public Function LvRefresh(Index As Integer)
    If Index = 0 Then
        Lv_MainInfo
        If lvList(1).Visible = True Then
            LvRefresh 1
        'ElseIf lvList(2).Visible = True Then
        '    LvRefresh 2
        End If
    ElseIf Index = 1 Then
        sSQL(1) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc " & _
                "From tbl_books " & _
                "WHERE (((tbl_books.pubid) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_books.desc;"
        Lv_OtherInfo
    End If
End Function

Public Function SearchItem()
    If iLvIndex = 0 Then
    sSearchFields = sFields(0)
        With frmSearch
            .srcNoOfCol = CountSplitItem(sColumns(0), ",") + 2
            Set .srcForm = Me
            .srcColumnHeaders = sColumns(0)
            .Show 1
        End With
    End If
End Function

Public Function Execute_SearchItem(sFilter As String)
    Dim vFields As Variant
    sSQL(0) = "SELECT tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst " & _
            "From tbl_publishers " & _
            "WHERE " & sFilter & " " & _
            "GROUP BY tbl_publishers.pubid, tbl_publishers.cmpny, tbl_publishers.cntry, tbl_publishers.add, tbl_publishers.tel, tbl_publishers.fx, tbl_publishers.wbst;"
    LvRefresh 0
End Function

