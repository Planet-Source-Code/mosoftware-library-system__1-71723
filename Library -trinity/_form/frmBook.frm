VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBook 
   Caption         =   "Books"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9525
   Begin VB.CommandButton cmdAu 
      Caption         =   "View Book Author(s)"
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
      Left            =   2760
      TabIndex        =   46
      Top             =   6360
      Width           =   2535
   End
   Begin VB.PictureBox picConn 
      Height          =   135
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   9195
      TabIndex        =   57
      Top             =   3120
      Width           =   9255
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "View Registered Books"
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
      TabIndex        =   45
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   47
      Top             =   0
      Width           =   9255
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   59
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   0
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   75
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   6
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
            TabIndex        =   68
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   74
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   0
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
               TabIndex        =   73
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   1
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
               TabIndex        =   72
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   2
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
               TabIndex        =   71
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   3
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
               TabIndex        =   70
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   4
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
               TabIndex        =   69
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   5
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
         TabIndex        =   53
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   0
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   0
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   0
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   0
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   0
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   0
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   0
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   54
            Top             =   45
            Width           =   2055
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1080
         Index           =   0
         Left            =   120
         TabIndex        =   7
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
         ForeColor       =   0
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
         Left            =   240
         TabIndex        =   61
         Top             =   2640
         Width           =   1230
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
         Left            =   960
         TabIndex        =   49
         Top             =   480
         Width           =   6990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Acquisition"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   48
         Top             =   240
         Width           =   1590
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmBook.frx":0000
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
         Width           =   9255
      End
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   50
      Top             =   3240
      Width           =   9255
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   1
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   55
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   1
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   1
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   1
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   1
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   1
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   1
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Previous 250"
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
            TabIndex        =   56
            Top             =   45
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   22
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
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   1
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   60
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   1
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   83
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   1
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   76
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   1
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   77
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   15
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
               TabIndex        =   78
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   16
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
               TabIndex        =   79
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   17
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
               TabIndex        =   80
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   18
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
               TabIndex        =   81
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   19
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
               TabIndex        =   82
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
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
         TabIndex        =   58
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmBook.frx":169B2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Books Registered"
         BeginProperty Font 
            Name            =   "Arial"
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
         Left            =   840
         TabIndex        =   52
         Top             =   240
         Width           =   2310
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of registered books and their current status."
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
         TabIndex        =   51
         Top             =   480
         Width           =   4440
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   2
      Left            =   1440
      TabIndex        =   62
      Top             =   3240
      Visible         =   0   'False
      Width           =   9255
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   2
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   9015
         TabIndex        =   84
         Top             =   840
         Width           =   9015
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   86
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   2
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   92
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   30
                  ToolTipText     =   "Insert Author"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnES 
               Height          =   375
               Index           =   2
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   91
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   2
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   90
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnD 
               Height          =   375
               Index           =   2
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   89
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnR 
               Height          =   375
               Index           =   2
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   88
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   2
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   87
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   2
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   85
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   2
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   2
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   4545
         TabIndex        =   63
         Top             =   2520
         Width           =   4545
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   2
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Previous 250"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   2
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   2
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   2
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   2
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   2
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   44
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
            Index           =   2
            Left            =   -480
            TabIndex        =   64
            Top             =   45
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   37
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
         Caption         =   "Yoi can view list of author of Selected Book."
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
         Left            =   1080
         TabIndex        =   67
         Top             =   480
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Book Author(s)"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   66
         Top             =   240
         Width           =   2280
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   2
         Left            =   240
         Picture         =   "frmBook.frx":2D364
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
         Index           =   2
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   2
         Left            =   120
         Top             =   240
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmBook"
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
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc " & _
            "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Book Info Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
            cmdAu.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
            cmdAu.Enabled = False
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Registered Books."
    ElseIf Index = 2 Then
         sSQL(2) = "SELECT tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn " & _
            "FROM tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid " & _
            "WHERE (((tbl_bookauthor.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn;"
            iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Books Author(s)."
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
    ElseIf Index = 2 Then
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
    ElseIf Index = 2 Then
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
    ElseIf Index = 2 Then
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
        Set dtrBook.DataSource = rsView1
        dtrBook.Show 1
    ElseIf Index = 1 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrRegBook.Sections("Section2")
            .Controls("lblisbn").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblpub").Caption = lvList(0).SelectedItem.SubItems(2)
            .Controls("lbltitle").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblyr").Caption = lvList(0).SelectedItem.SubItems(3)
            .Controls("lbldesc").Caption = lvList(0).SelectedItem.SubItems(4)
        End With
        Set dtrRegBook.DataSource = rsView1
        dtrRegBook.Show 1
    ElseIf Index = 2 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrRegBook.Sections("Section2")
            .Controls("lblisbn").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblpub").Caption = lvList(0).SelectedItem.SubItems(2)
            .Controls("lbltitle").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblyr").Caption = lvList(0).SelectedItem.SubItems(3)
            .Controls("lbldesc").Caption = lvList(0).SelectedItem.SubItems(4)
        End With
        Set dtrBookAu.DataSource = rsView1
        dtrBookAu.Show 1
    End If
End Function

Private Sub btnPrev_Click(Index As Integer)
    If Index = 0 Then
        iStartPage(Index) = iStartPage(Index) - iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    ElseIf Index = 1 Then
        iStartPage(Index) = iStartPage(Index) - iNoPage(Index)
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), sNoRec(Index)
    ElseIf Index = 2 Then
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

Private Sub cmdAu_Click()
    If view_other = True And lvList(2).Visible = True Then
        fraList(2).Visible = False
        picConn.Visible = False
        view_other = False
    Else
        sSQL(2) = "SELECT tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn " & _
            "FROM tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid " & _
            "WHERE (((tbl_bookauthor.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn;"
        Lv_OtherInfo2
        fraList(2).Visible = True
        picConn.Visible = True
        view_other = True
        fraList(1).Visible = False
    End If
    INT_SIZE = Me.ScaleHeight / 2
    Listview_Resize
End Sub

Private Sub cmdReg_Click()
    If view_other = True And lvList(1).Visible = True Then
        fraList(1).Visible = False
        picConn.Visible = False
        view_other = False
    Else
        sSQL(1) = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks;"
        Lv_OtherInfo
        fraList(1).Visible = True
        picConn.Visible = True
        view_other = True
        fraList(2).Visible = False
    End If
    INT_SIZE = Me.ScaleHeight / 2
    Listview_Resize
End Sub

Private Sub Form_Load()
    view_other = False
    INT_SIZE = Me.ScaleHeight / 2
    Set_Icon_btn Me, 2
    fraList(1).Visible = False
    picConn.Visible = False
    sSQL(0) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc " & _
        "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid;"
    Lv_MainInfo
    lvList(0).Refresh
    frmMain.TabMainIni 2, "Books", 2
    ListOfDisabledBtn
End Sub

Public Function ListOfDisabledBtn()
    btnES(2).Visible = False
    btnNew(2).Visible = False
    btnEdited(2).Visible = False
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
    cmdAu.Top = Me.ScaleHeight - (cmdAu.Height + 120)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.TabMainIni 1, "Books", 2
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
            frmBook_AE.bStat = False
            Set frmBook_AE.fCur = Me
            frmBook_AE.Show 1
        End If
    ElseIf Index = 1 Then
        If iRec(Index) > 0 Then
            frmBookReg_AE.bStat = False
            Set frmBookReg_AE.fCur = Me
            frmBookReg_AE.Show 1
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
        ElseIf lvList(2).Visible = True Then
            LvRefresh 2
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
        i = 2
        fraList(i).Move 120, (picConn.Top + picConn.Height) - 15, Me.ScaleWidth - 240, (Me.ScaleHeight - ((picConn.Top - 15) + 240 + 240 + 120 + 80))
        lvList(i).Move 120, 1320, fraList(i).Width - 240, fraList(i).Height - (1320 + 240 + 240 + 120)
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 90)
'        lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 120, 840, fraList(i).Width - 240
        picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
    End If
End Sub

'This Function is use to Refresh the Info on Listview(0)
Public Sub Lv_MainInfo()
    sColWidth(0) = "1650,3990,2369,1235,1964"
    sColumns(0) = "ISBN,Title,Publisher,Year Publish,Description"
    iIcon(0) = 2
    iLoop(0) = CountSplitItem(sColumns(0), ",")
    sFields(0) = "isbn,title,cmpny,yrpub,desc"
    sNoRec(0) = "No Current Book Info Records."
    iStartPage(0) = 1
    iNoPage(0) = 75
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0))
    If iRec(0) > 0 Then
        cmdReg.Enabled = True
        cmdAu.Enabled = True
    Else
        view_other = False
        picConn.Visible = False
        Listview_Resize
        cmdReg.Enabled = False
        cmdAu.Enabled = False
    End If
End Sub

'This Function is used to Refresh the Info on Listview(1) Index
Public Sub Lv_OtherInfo()
    sColWidth(1) = "1700,1200,1000,1000,1000,2000"
    sColumns(1) = "Registered ID,Barcode,Borrow,Pending,Remarks"
    sFields(1) = "rb_id,barcode,borrow,pending,remarks"
    iIcon(1) = 3
    iLoop(1) = CountSplitItem(sColumns(1), ",")
    sNoRec(1) = "No Current Registered Books."
    iStartPage(1) = 1
    iNoPage(1) = 75
    iRec(1) = LvPageStat(Me, 1, sSQL(1), iStartPage(1), iNoPage(1), iIcon(1), sColumns(1), iLoop(1), sColWidth(1), sFields(1), sNoRec(1))
    lvList(1).Refresh
End Sub

'This function is used to Refresh the Info on Listview(2) Index
Public Sub Lv_OtherInfo2()
    sColWidth(2) = "1154,2954,1484"
    sColumns(2) = "Author ID,Author Name,Year Born"
    sFields(2) = "auid,author,yrborn"
    iIcon(2) = 4
    iLoop(2) = CountSplitItem(sColumns(2), ",")
    sNoRec(2) = "No Current Author(s) Info."
    iStartPage(2) = 1
    iNoPage(2) = 75
    iRec(2) = LvPageStat(Me, 2, sSQL(2), iStartPage(2), iNoPage(2), iIcon(2), sColumns(2), iLoop(2), sColWidth(2), sFields(2), sNoRec(2))
    lvList(2).Refresh
End Sub

'This will be used to Control the Update of Information
Private Sub btnEdited_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc " & _
            "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid " & _
            "Where tbl_books.DateModified between #" & Date & "# And #" & Date & "# " & _
            "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Updated Book Info Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
            cmdAu.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
            cmdAu.Enabled = False
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_reg_books.DateModified) Between #" & Date & "# And #" & Date & "#)) " & _
                "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Updated Registered Books."
    ElseIf Index = 2 Then
        sSQL(Index) = "SELECT tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn " & _
            "FROM tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid " & _
            "WHERE (((tbl_authors.DateModified) between #" & Date & "# And #" & Date & "#) AND ((tbl_bookauthor.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Updated Authors(0) Books."
    End If
End Sub

'This will used to Get New Create Item
Public Function btnNew_Load(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc " & _
            "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid " & _
            "Where tbl_books.DateAdded between #" & Date & "# And #" & Date & "# " & _
            "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), "No Current New Book Info Added Records.")
        If iRec(0) > 0 Then
            cmdReg.Enabled = True
            cmdAu.Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdReg.Enabled = False
            cmdAu.Enabled = False
        End If
        If lvList(1).Visible = True Then
            LvRefresh 1
        ElseIf lvList(2).Visible = True Then
            LvRefresh 2
        End If
    ElseIf Index = 1 Then
        sSQL(Index) = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks " & _
                "From tbl_reg_books " & _
                "WHERE (((tbl_reg_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_reg_books.DateAdded) Between #" & Date & "# And #" & Date & "#)) " & _
                "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current New Registered Books."
    ElseIf Index = 2 Then
        sSQL(Index) = "SELECT tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn " & _
            "FROM tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid " & _
            "WHERE tbl_authors.DateAdded between #" & Date & "# And #" & Date & "# AND tbl_bookauthor.isbn Like '" & lvList(0).SelectedItem.SubItems(1) & "' " & _
            "GROUP BY tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn;"
        iStartPage(Index) = 1
        LvPageStat Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current New Author(s) Books."
    End If
End Function

'This will be used to Where to Create New Item
Public Function LvNew(Index As Integer)
    If Index = 0 Then
        frmBook_AE.bStat = True
        Set frmBook_AE.fCur = Me
        frmBook_AE.Show 1
    ElseIf Index = 1 Then
        frmBookReg_AE.bStat = True
        Set frmBookReg_AE.fCur = Me
        frmBookReg_AE.Show 1
    ElseIf Index = 2 Then
        AddAuthor 2
    End If
End Function

Public Function AddAuthor(Index As Integer)
    Dim newFormSelect As New frmSelect
    Dim gSQL As String, gLblHead As String, gLblDef As String
    Dim gXY As String, gTitle As String, gColumns As String
    Dim gColWidth As String, gFields As String, gLoop As Integer
    Dim gIcon As Integer, gLvIcon As Integer, sNoRec As String
    Dim sSQL As String
    Dim mRow As ListItem
    Dim SelectedItem As String
    
    If Index = 2 Then
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
    
    If Index = 2 And Len(SelectedItem) > 0 Then
        sSQL = "SELECT tbl_bookauthor.isbn " & _
            "From tbl_bookauthor " & _
            "WHERE (((tbl_bookauthor.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_bookauthor.auid) Like '" & SelectedItem & "')) " & _
            "GROUP BY tbl_bookauthor.isbn;"
        If isRecordExist(sSQL) = False Or lvList(2).ListItems.Count = 0 Then
            INSERT_DATA "tbl_bookauthor", "auid,isbn", SelectedItem & "," & lvList(0).SelectedItem.SubItems(1), ",", False
            Set mRow = frmBook.lvList(2).ListItems.Add(, , , , 4)
            mRow.SubItems(1) = SelectedItem
            mRow.SubItems(2) = FindFieldValue("SELECT tbl_authors.author " & _
                        "From tbl_authors " & _
                        "WHERE (((tbl_authors.auid) Like '" & SelectedItem & "')) " & _
                        "GROUP BY tbl_authors.author;", "author")
            mRow.SubItems(3) = FindFieldValue("SELECT tbl_authors.yrborn " & _
                        "From tbl_authors " & _
                        "WHERE (((tbl_authors.auid) Like '" & SelectedItem & "')) " & _
                        "GROUP BY tbl_authors.yrborn;", "yrborn")
            LvSearchItem frmBook.lvList(2), SelectedItem
        Else
            MsgBox "You already add this Author.", vbExclamation, "Record Exist"
        End If
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
            sMsgId = "ISBN"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_books", "ISBN", lvList(Index).SelectedItem.SubItems(1)
                lvList(Index).ListItems.Remove lvList(Index).SelectedItem.Index
                If lvList(Index).ListItems.Count = 0 Then
                    Call Lv_MainInfo
                End If
                If lvList(1).Visible = True Then
                    LvRefresh 1
                ElseIf lvList(2).Visible = True Then
                    LvRefresh 2
                End If
                lblPageInfo(Index).Caption = iStartPage(Index) & " - " & (iStartPage(Index) + (iNoPage(Index) - 2)) & " of " & iRec(Index) - 1
            End If
        ElseIf Index = 1 Then
            sMsgDel = "You are about to delete this record?"
            sMsgId = "Registered Book"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_reg_books", "isbn,rb_id", lvList(0).SelectedItem.SubItems(1) & "," & lvList(1).SelectedItem.SubItems(1)
                lvList(Index).ListItems.Remove lvList(Index).SelectedItem.Index
                If lvList(Index).ListItems.Count = 0 Then
                    Call Lv_OtherInfo
                End If
                lblPageInfo(Index).Caption = iStartPage(Index) & " - " & (iStartPage(Index) + (iNoPage(Index) - 2)) & " of " & iRec(Index) - 1
            End If
        ElseIf Index = 2 Then
            sMsgDel = "You are about to delete this record?"
            sMsgId = "Author ID"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_bookauthor", "isbn,auid", lvList(0).SelectedItem.SubItems(1) & "," & lvList(2).SelectedItem.SubItems(1)
                lvList(Index).ListItems.Remove lvList(Index).SelectedItem.Index
                If lvList(Index).ListItems.Count = 0 Then
                    Call Lv_OtherInfo2
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
        frmBook_AE.bStat = False
        Set frmBook_AE.fCur = Me
        frmBook_AE.Show 1
    ElseIf Index = 1 Then
        frmBookReg_AE.bStat = False
        Set frmBookReg_AE.fCur = Me
        frmBookReg_AE.Show 1
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
    ElseIf Index = 2 Then
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
        ElseIf lvList(2).Visible = True Then
            LvRefresh 2
        End If
    ElseIf Index = 1 Then
        sSQL(1) = "SELECT tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks " & _
            "From tbl_reg_books " & _
            "WHERE (((tbl_reg_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_reg_books.rb_id, tbl_reg_books.barcode, tbl_reg_books.borrow, tbl_reg_books.pending, tbl_reg_books.remarks;"
        Lv_OtherInfo
    ElseIf Index = 2 Then
        sSQL(2) = "SELECT tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn " & _
            "FROM tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid " & _
            "WHERE (((tbl_bookauthor.isbn) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
            "GROUP BY tbl_authors.auid, tbl_authors.author, tbl_authors.yrborn;"
        Lv_OtherInfo2
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
    sSQL(0) = "SELECT tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc " & _
        "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid " & _
        "WHERE " & sFilter & "  " & _
        "GROUP BY tbl_books.isbn, tbl_books.title, tbl_books.yrpub, tbl_publishers.cmpny, tbl_books.desc;"
    LvRefresh 0
End Function