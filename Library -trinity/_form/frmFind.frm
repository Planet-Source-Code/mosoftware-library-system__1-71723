VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraQuery 
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
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkWhole 
         Caption         =   "Find Whole Word Only"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   2115
         Width           =   2055
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Match Case"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   2115
         Width           =   1215
      End
      Begin VB.ComboBox cboFields 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmFind.frx":0000
         Left            =   1200
         List            =   "frmFind.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cboFind 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox cboDir 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmFind.frx":0004
         Left            =   1200
         List            =   "frmFind.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
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
         Height          =   340
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Direction:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1740
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "Find What:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Look In:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1395
         Width           =   585
      End
      Begin VB.Image Image3 
         Height          =   600
         Index           =   1
         Left            =   120
         Picture         =   "frmFind.frx":0024
         Stretch         =   -1  'True
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Text"
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
         Left            =   840
         TabIndex        =   8
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Fields then fill the following the click Search."
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
         Left            =   840
         TabIndex        =   7
         Top             =   405
         Width           =   3630
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim fForm  As Form
Dim fLv As ListView
Dim iSelectedIndex As Integer
Dim iSelectedSubItem As Integer
Dim iFind As Integer
Dim iFindY As Integer
Dim iForBoldText As Integer
Dim iEndLoop As Integer
Dim iSubItem As Integer

Private Sub cboDir_Click()
    On Error Resume Next
    If iFind > 0 Then
        iSelectedIndex = fLv.SelectedItem.Index
        iSelectedSubItem = 1
        iFind = iSelectedIndex
        iFindY = iSelectedSubItem
    End If
End Sub

Private Sub cboFind_Change()
    If isNull(cboFind.Text) = False Then cmdSearch.Enabled = True Else cmdSearch.Enabled = False
    
    iSelectedIndex = fLv.SelectedItem.Index
    iFind = iSelectedIndex
End Sub

Private Sub cboFind_Click()
    If isNull(cboFind.Text) = False Then cmdSearch.Enabled = True Else cmdSearch.Enabled = False
    
    iSelectedIndex = fLv.SelectedItem.Index
    iSelectedSubItem = 0
    iFind = iSelectedIndex
    iFindY = iSelectedSubItem
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    FindText_Field
End Sub

Public Function Refresh_Values(cLv As ListView)
    Dim i As Integer
    Set fLv = Nothing
    Set fLv = cLv
    cboFields.Clear
    cboFields.AddItem "All Fields"
    For i = 2 To cLv.ColumnHeaders.Count
        cboFields.AddItem cLv.ColumnHeaders(i).Text
    Next
    cboFields.ListIndex = 0
    cboDir.ListIndex = 0
    If Me.Visible = False Then
        Me.Show , frmMain
    End If
    iFind = 0
End Function

Public Function FindText_Field()
    Dim i As Integer
    Dim j As Integer
    If cboDir.ListIndex = 1 Then
        If iFind = 1 Then
            MsgBox "Finished searching records. The search item was not found.", vbInformation, "Search"
            Exit Function
        End If
        iFind = iFind - 1
        iEndLoop = 1
        For i = iFind To iEndLoop Step -1
            iSubItem = cboFields.ListIndex
            If cboFields.ListIndex > 0 And cboFields.ListIndex <= fLv.ColumnHeaders.Count Then
                If FindSubFields(i, iSubItem) = True Then
                    Exit For
                End If
            ElseIf cboFields.ListIndex = 0 Then
                For j = iFindY To fLv.ColumnHeaders.Count
                    If j = 0 Then
                        If FindFirstFields(i, j) = True Then
                            Exit Function
                        End If
                    ElseIf j > 0 And j <= (fLv.ColumnHeaders.Count - 1) Then
                        If FindSubFields(i, j) = True Then
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
    ElseIf cboDir.ListIndex = 2 Then
        If iFind = fLv.ListItems.Count Then
            MsgBox "Finished searching records. The search item was not found.", vbInformation, "Search"
            Exit Function
        End If
        iFind = iFind + 1
        iEndLoop = fLv.ListItems.Count
        For i = iFind To iEndLoop Step 1
            iSubItem = cboFields.ListIndex
            If cboFields.ListIndex > 0 And cboFields.ListIndex <= fLv.ColumnHeaders.Count Then
                If FindSubFields(i, iSubItem) = True Then
                    Exit For
                End If
            ElseIf cboFields.ListIndex = 0 Then
                For j = iFindY To fLv.ColumnHeaders.Count
                    If j = 0 Then
                        If FindFirstFields(i, j) = True Then
                            Exit Function
                        End If
                    ElseIf j > 0 And j <= (fLv.ColumnHeaders.Count - 1) Then
                        iFindY = cboFields.ListIndex - 1
                        If FindSubFields(i, j) = True Then
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
    ElseIf cboDir.ListIndex = 0 Then
        If iFind > fLv.ListItems.Count Then
            iFind = 1
            iEndLoop = iSelectedIndex
        Else
            If cboFields.ListIndex = 0 Then
                If iFindY = (fLv.ColumnHeaders.Count - 1) Then
                    iFind = iFind + 1
                    iFindY = 0
                Else
                    iFindY = iFindY + 1
                End If
            Else
                iFind = iFind + 1
            End If
            iEndLoop = fLv.ListItems.Count
        End If
        If iFind = iSelectedIndex And iFindY = iSelectedSubItem Then
            MsgBox "Finished searching records. The search item was not found.", vbInformation, "Search"
            Exit Function
        End If
        For i = iFind To iEndLoop Step 1
            iSubItem = cboFields.ListIndex
            If cboFields.ListIndex > 0 And cboFields.ListIndex <= fLv.ColumnHeaders.Count Then
                If FindSubFields(i, iSubItem) = True Then
                    Exit For
                End If
            ElseIf cboFields.ListIndex = 0 Then
                For j = iFindY To fLv.ColumnHeaders.Count
                    If j > iFindY And j <= (fLv.ColumnHeaders.Count - 1) Then
                        If FindSubFields(i, j) = True Then
                            Exit Function
                        End If
                    End If
                Next
            End If
        Next
    End If
End Function

Public Function FindFirstFields(i As Integer, iSubItem As Integer) As Boolean
    If chkCase.Value = 0 And chkWhole.Value = 0 Then 'not Case Sensitive and not Search Whole Word
        If InStr(UCase(fLv.ListItems(i).Text), UCase(cboFind.Text)) Then
            'Set bold text to false
            
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindFirstFields = True
            iFind = i
        Else
            FindFirstFields = False
        End If
    ElseIf chkCase.Value = 0 And chkWhole.Value = 1 Then
        If UCase(fLv.ListItems(i).Text) = UCase(cboFind.Text) Then
            'Set bold text to false
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindFirstFields = True
            iFind = i
        Else
            FindFirstFields = False
        End If
    ElseIf chkCase.Value = 1 And chkWhole = 1 Then
        If fLv.ListItems(i).Text = cboFind.Text Then
        
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindFirstFields = True
            iFind = i
        Else
            FindFirstFields = False
        End If
    ElseIf chkCase.Value = 1 And chkWhole = 0 Then
        If InStr(fLv.ListItems(i).Text, cboFind.Text) > 0 Then
            If Mid(fLv.ListItems(i).Text, InStr(fLv.ListItems(i).Text, cboFind.Text), Len(cboFind.Text)) = cboFind.Text Then
                
                SetFindBoldtoFalse
                SetFindBoldtoTrue i, iSubItem
                
                fLv.ListItems(i).Selected = True
                fLv.ListItems(i).EnsureVisible
                FindFirstFields = True
                iFind = i
            Else
                FindFirstFields = False
            End If
        End If
    End If
End Function

Public Function FindSubFields(i As Integer, iSubItem As Integer) As Boolean
    If chkCase.Value = 0 And chkWhole.Value = 0 Then
        If InStr(UCase(fLv.ListItems(i).SubItems(iSubItem)), UCase(cboFind.Text)) Then
        
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindSubFields = True
            iFind = i
        Else
            FindSubFields = False
        End If
    ElseIf chkCase.Value = 0 And chkWhole.Value = 1 Then
        If UCase(fLv.ListItems(i).SubItems(iSubItem)) = UCase(cboFind.Text) Then
        
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindSubFields = True
            iFind = i
        Else
            FindSubFields = False
        End If
    ElseIf chkCase.Value = 1 And chkWhole = 1 Then
        If fLv.ListItems(i).SubItems(iSubItem) = cboFind.Text Then
            
            SetFindBoldtoFalse
            SetFindBoldtoTrue i, iSubItem
            
            fLv.ListItems(i).Selected = True
            fLv.ListItems(i).EnsureVisible
            FindSubFields = True
            iFind = i
            iFindY = iSubItem
        Else
            FindSubFields = False
        End If
    ElseIf chkCase.Value = 1 And chkWhole = 0 Then
        If InStr(fLv.ListItems(i).SubItems(iSubItem), cboFind.Text) > 0 Then
            If Mid(fLv.ListItems(i).SubItems(iSubItem), _
                InStr(fLv.ListItems(i).SubItems(iSubItem), cboFind.Text), _
                    Len(cboFind.Text)) = cboFind.Text Then
                
                SetFindBoldtoFalse
                SetFindBoldtoTrue i, iSubItem

                fLv.ListItems(i).Selected = True
                fLv.ListItems(i).EnsureVisible
                FindSubFields = True
                iFind = i
                iFindY = iSubItem
            Else
                FindSubFields = False
            End If
        End If
    End If
End Function

Public Function SetFindBoldtoTrue(i As Integer, iSubItem As Integer)
    If iSubItem > 0 And iSubItem <= (fLv.ColumnHeaders.Count - 1) Then
        fLv.ListItems(i).ListSubItems(iSubItem).Bold = True
    End If
End Function

Public Function SetFindBoldtoFalse()
    On Error Resume Next
    If cboDir.ListIndex = 1 Then
        For iForBoldText = 0 To fLv.ColumnHeaders.Count - 1
            If iForBoldText = 0 And iFind >= iEndLoop Then
                fLv.ListItems(iFind + 1).Bold = False
            ElseIf iFind >= iEndLoop Then
                fLv.ListItems(iFind + 1).ListSubItems(iForBoldText).Bold = False
            End If
        Next
    ElseIf cboDir.ListIndex = 2 And iFind <= iEndLoop Then
        For iForBoldText = 0 To fLv.ColumnHeaders.Count - 1
            If iForBoldText = 0 Then
                fLv.ListItems(iFind - 1).Bold = False
            ElseIf iFind <= iEndLoop Then
                fLv.ListItems(iFind - 1).ListSubItems(iForBoldText).Bold = False
            End If
        Next
    ElseIf cboDir.ListIndex = 0 And iFind <= iEndLoop Then
        For iForBoldText = 0 To fLv.ColumnHeaders.Count - 1
            If iForBoldText = 0 Then
                fLv.ListItems(iFind - 1).Bold = False
            ElseIf iFind <= iEndLoop Then
                fLv.ListItems(iFind - 1).ListSubItems(iForBoldText).Bold = False
            End If
        Next
    End If
End Function

Private Sub Form_Activate()
    iSelectedIndex = fLv.SelectedItem.Index
    iFind = iSelectedIndex
End Sub

                              ''''''''''''''''''''''''''''''''''''''''''''''''''
                              '''''''List of New Function\Methods Created'''''''
                              ''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
End Sub

