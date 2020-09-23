VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
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
      Height          =   975
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame fraList 
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   9255
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   5520
         ScaleHeight     =   345
         ScaleWidth      =   3465
         TabIndex        =   19
         Top             =   2520
         Width           =   3465
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   0
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   0
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   0
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   0
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   0
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   0
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "View New Records"
            Top             =   0
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   0
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "View All Records"
            Top             =   0
            Visible         =   0   'False
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
            Left            =   0
            TabIndex        =   20
            Top             =   45
            Width           =   2055
         End
      End
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   9015
         Begin VB.CommandButton btnCN 
            Height          =   315
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnES 
            Height          =   315
            Index           =   0
            Left            =   320
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnS 
            Height          =   315
            Index           =   0
            Left            =   640
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnD 
            Height          =   315
            Index           =   0
            Left            =   980
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnR 
            Height          =   315
            Index           =   0
            Left            =   1300
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnP 
            Height          =   315
            Index           =   0
            Left            =   1640
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnC 
            Height          =   315
            Index           =   0
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
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
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1230
      End
   End
   Begin VB.Label lblDef 
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
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   360
      Width           =   6990
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Information"
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
      Left            =   720
      TabIndex        =   22
      Top             =   120
      Width           =   1785
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape spMag 
      BackColor       =   &H000AA27C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedItem As String
Dim iStartPage(1) As Long
Dim iNoPage(1) As Integer
Dim iRec(1) As Long
Dim sSQL(1) As String
Public strSQL As String
Public sColumn As String, sColWidth As String, sFields As String
Public sNoRec As String
Public iIcon As Integer, iLoop As Integer


Private Sub btnFirst_Click(Index As Integer)
    iStartPage(Index) = 1
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon, sColumn, iLoop, sColWidth, sFields, sNoRec)
End Sub

Private Sub btnLast_Click(Index As Integer)
    Dim iLastNoPage As Long
    iLastNoPage = 1
    Do While iLastNoPage <= iRec(Index)
        iLastNoPage = iLastNoPage + iNoPage(Index)
    Loop
    iStartPage(Index) = iLastNoPage - iNoPage(Index)
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon, sColumn, iLoop, sColWidth, sFields, sNoRec)
End Sub



Private Sub btnNext_Click(Index As Integer)
    iStartPage(Index) = iStartPage(Index) + iNoPage(Index)
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon, sColumn, iLoop, sColWidth, sFields, sNoRec)
End Sub

Private Sub btnPrev_Click(Index As Integer)
    iStartPage(Index) = iStartPage(Index) - iNoPage(Index)
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon, sColumn, iLoop, sColWidth, sFields, sNoRec)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    SelectedItem = lvList(0).SelectedItem.SubItems(1)
    MsgBox SelectedItem
    Unload Me
    'ComputeWidth
End Sub


Private Sub Form_Load()
    cmdSelect.Picture = frmMain.iHead.ListImages(3).ExtractIcon
    cmdCancel.Picture = frmMain.iHead.ListImages(4).ExtractIcon
    sSQL(0) = strSQL
    iStartPage(0) = 1
    iNoPage(0) = 75
    Lv_MainInfo
    lvList(0).Refresh
    'ActivateXPTheme wXP
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Set_Icon_btn Me, 0
    Listview_Resize
    spMag(0).Move 120, 120, Me.fraList(0).Width
    lblDef(0).Move 840, 360, Me.fraList(0).Width - 840
    cmdSelect.Move Me.ScaleWidth - (cmdSelect.Width + 120)
    cmdCancel.Move Me.ScaleWidth - (cmdSelect.Width + 120)
End Sub

Private Sub lblSelected_Click(Index As Integer)
frmFind.Show 1
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

Private Sub lvList_DblClick(Index As Integer)
    If vbLeftButton = 1 Then
        SelectedItem = lvList(0).SelectedItem.SubItems(1)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If iRec(Index) > 0 Then
        lblSelected(Index).Caption = "Selected Record: " & (iStartPage(Index) - 1) + lvList(Index).SelectedItem.Index
    Else
        lblSelected(Index).Caption = "Selected Record: None"
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''List of New Function Created'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ComputeWidth()
    Dim i As Integer
    Dim s As String
    For i = 1 To 3
        Debug.Print lvList(0).ColumnHeaders(i).Width
    Next
    Debug.Print " "
    Debug.Print Me.Width & ", " & Me.Height
End Sub

Public Sub Lv_MainInfo()
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon, sColumn, iLoop, sColWidth, sFields, sNoRec)
End Sub

Private Sub Listview_Resize()
    On Error Resume Next
    Dim i As Integer
    i = 0
    fraList(i).Move 120, 720, Me.ScaleWidth - (cmdSelect.Width + 120 + 240), Me.ScaleHeight - (720 + 120)
    lvList(i).Move 120, 240, fraList(i).Width - 240, fraList(i).Height - (240 + 240 + 240 + 120)
    picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
    lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
    lvList(i).SelectedItem.EnsureVisible
    picData(i).Move 120, 840, fraList(i).Width - 240
    btnC(i).Left = picData(i).Width - (btnC(i).Width + 80)
End Sub
