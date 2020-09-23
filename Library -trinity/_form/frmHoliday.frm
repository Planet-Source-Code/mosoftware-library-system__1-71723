VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHoliday 
   Caption         =   "Holidays"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   11085
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   10
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   0
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   24
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   11
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   22
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   23
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
               TabIndex        =   20
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   21
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
               TabIndex        =   18
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   19
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
               TabIndex        =   16
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   17
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
               TabIndex        =   14
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   15
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
               TabIndex        =   12
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   1
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   0
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   0
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   0
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   0
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "First"
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
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   0
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "View All Updated Records"
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
            TabIndex        =   9
            Top             =   45
            Width           =   2055
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   26
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
         Index           =   2
         Left            =   840
         Picture         =   "frmHoliday.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   600
         Picture         =   "frmHoliday.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmHoliday.frx":0884
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
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view List of Holidays for checking the Return Date to be available to return."
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
         Left            =   1440
         TabIndex        =   28
         Top             =   480
         Width           =   6045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holidays"
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
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   795
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
End
Attribute VB_Name = "frmHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim view_other As Boolean
Dim INT_SIZE As Integer
Dim int_size_active As Integer
Dim CURR_COL As Integer
Dim iStartPage(0) As Long
Dim iNoPage(0) As Integer
Dim iRec(0) As Long
Dim sSQL(0) As String

Dim sColumns(0) As String, sColWidth(0) As String, sFields(0) As String
Dim iIcon(0) As Integer, iLoop(0) As Integer, sNoRec(0) As String
Public iLvIndex As Integer
Public sSearchFields As String

Private Sub btnAll_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status " & _
                    "FROM tbl_holiday;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Holiday(s) Records.")
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

Private Sub btnEdited_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status " & _
                    "From tbl_holiday " & _
                    "WHERE (((tbl_holiday.DateModified) Between #" & Date & "# And #" & Date & "#)) " & _
                    "GROUP BY tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Updated Holiday(s) Records.")
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
        Set dtrHoliday.DataSource = rsView1
        dtrHoliday.Show 1
    End If
End Function

Private Sub btnPrev_Click(Index As Integer)
    If Index = 0 Then
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

Private Sub Form_Load()
    view_other = False
    INT_SIZE = Me.ScaleHeight / 2
    Set_Icon_btn Me, 0
    'fraList(1).Visible = False
    sSQL(0) = "SELECT tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status " & _
            "FROM tbl_holiday;"
    Lv_MainInfo
    lvList(0).Refresh
    frmMain.TabMainIni 2, "Holidays", 15
    'ListOfDisabledBtn
End Sub

Private Sub Form_LostFocus()
    If frmFind.Visible = True Then
        frmFind.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call Listview_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.TabMainIni 1, "Holidays", 15
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
            frmHoliday_AE.bStat = False
            Set frmHoliday_AE.fCur = Me
            frmHoliday_AE.Show 1
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
    i = 0
    fraList(i).Move 120, 0, Me.ScaleWidth - 240, Me.ScaleHeight - 180
    lvList(i).Move 120, 1320, fraList(i).Width - 240, fraList(i).Height - (1320 + 240 + 240 + 120)
    spMag(i).Move 0, 240, fraList(i).Width
    picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
    lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
    lvList(i).SelectedItem.EnsureVisible
    picData(i).Move 120, 840, fraList(i).Width - 240
    picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
End Sub

'This Function is use to Refresh the Info on Listview(0)
Public Sub Lv_MainInfo()
    sColWidth(0) = "1700,2000,1700,1500"
    sColumns(0) = "Holiday ID,Holiday Name,Holiday Date,Status(On/Off)"
    iIcon(0) = 15
    iLoop(0) = CountSplitItem(sColumns(0), ",")
    sFields(0) = "h_id,h_name,h_date,h_status"
    sNoRec(0) = "No Current Holidays Word(s) Records."
    iStartPage(0) = 1
    iNoPage(0) = 75
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0))
End Sub

'This will used to Get New Create Item
Public Function btnNew_Load(Index As Integer)
     If Index = 0 Then
        sSQL(Index) = "SELECT tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status " & _
                    "From tbl_holiday " & _
                    "WHERE (((tbl_holiday.DateAdded) Between #" & Date & "# And #" & Date & "#)) " & _
                    "GROUP BY tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), "No Current New Holiday(s) Added Records.")
    End If
End Function

'This will be used to Where to Create New Item
Public Function LvNew(Index As Integer)
    If Index = 0 Then
        frmHoliday_AE.bStat = True
        Set frmHoliday_AE.fCur = Me
        frmHoliday_AE.Show 1
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
            sMsgId = "Holiday ID"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_holiday", "h_id", lvList(Index).SelectedItem.SubItems(1)
                lvList(Index).ListItems.Remove lvList(Index).SelectedItem.Index
                If lvList(Index).ListItems.Count = 0 Then
                    Call Lv_MainInfo
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
        frmHoliday_AE.bStat = False
        Set frmHoliday_AE.fCur = Me
        frmHoliday_AE.Show 1
    End If
End Function

'This will be use to CLosed the Records
Public Function LvClose(Index As Integer)
    If Index = 0 Then
        Unload Me
    End If
End Function

Public Function LvRefresh(Index As Integer)
    If Index = 0 Then
        Lv_MainInfo
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
    sSQL(0) = "SELECT tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status " & _
            "From tbl_holiday " & _
            "WHERE " & sFilter & " " & _
            "GROUP BY tbl_holiday.h_id, tbl_holiday.h_name, tbl_holiday.h_date, tbl_holiday.h_status;"
    LvRefresh 0
End Function
