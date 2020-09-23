VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBorrower 
   Caption         =   "Borrowers"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   13830
   Begin VB.Frame fraList 
      Height          =   5535
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton cmdProfile 
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "View Profile"
         Top             =   1320
         Width           =   315
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
         Height          =   1935
         Left            =   480
         TabIndex        =   93
         Top             =   1200
         Visible         =   0   'False
         Width           =   10815
         Begin VB.Label lblCel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "CELL NO."
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
            Left            =   6480
            TabIndex        =   101
            Top             =   1080
            Width           =   2805
         End
         Begin VB.Label lblTel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TEL. NO."
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
            Left            =   6480
            TabIndex        =   100
            Top             =   840
            Width           =   2805
         End
         Begin VB.Label lblBtype 
            BackColor       =   &H00FFFFFF&
            Caption         =   "B. TYPE"
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
            Left            =   6480
            TabIndex        =   99
            Top             =   600
            Width           =   2805
         End
         Begin VB.Label lblGen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "GENDER"
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
            Left            =   6480
            TabIndex        =   97
            Top             =   360
            Width           =   2805
         End
         Begin VB.Label lblBday 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BDAY"
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
            Left            =   2520
            TabIndex        =   96
            Top             =   840
            Width           =   3045
         End
         Begin VB.Label lblName 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NAME"
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
            Left            =   2520
            TabIndex        =   95
            Top             =   600
            Width           =   3045
         End
         Begin VB.Label lblID 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ID NO."
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
            Left            =   2520
            TabIndex        =   94
            Top             =   360
            Width           =   3045
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Borrowers Profile"
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
            Index           =   3
            Left            =   120
            TabIndex        =   110
            Top             =   120
            Width           =   1230
         End
         Begin VB.Image ImgPic 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1410
            Left            =   120
            Picture         =   "frmBorrower.frx":0000
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1485
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
            Left            =   1680
            TabIndex        =   109
            Top             =   360
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
            Left            =   1680
            TabIndex        =   108
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label7 
            BackColor       =   &H00808080&
            Caption         =   "BDAY"
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
            Left            =   1680
            TabIndex        =   107
            Top             =   840
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
            Left            =   5640
            TabIndex        =   106
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label11 
            BackColor       =   &H00808080&
            Caption         =   "ADDRESS"
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
            Left            =   1680
            TabIndex        =   105
            Top             =   1080
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
            Left            =   5640
            TabIndex        =   104
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            Caption         =   "TEL. NO."
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
            Left            =   5640
            TabIndex        =   103
            Top             =   840
            Width           =   765
         End
         Begin VB.Label Label8 
            BackColor       =   &H00808080&
            Caption         =   "CELL NO."
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
            Left            =   5640
            TabIndex        =   102
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label lblAdd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ADDRESS"
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
            Height          =   675
            Left            =   2520
            TabIndex        =   98
            Top             =   1080
            Width           =   3045
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   0
         Left            =   7800
         ScaleHeight     =   345
         ScaleWidth      =   3585
         TabIndex        =   36
         Top             =   5040
         Width           =   3585
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   0
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   0
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   0
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   0
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   0
            Left            =   1515
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Previous"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   0
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   0
            Left            =   3195
            Style           =   1  'Graphical
            TabIndex        =   37
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
            Left            =   0
            TabIndex        =   44
            Top             =   45
            Width           =   1095
         End
      End
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   480
         ScaleHeight     =   345
         ScaleWidth      =   10815
         TabIndex        =   28
         Top             =   840
         Width           =   10815
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   30
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   0
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   24
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
            Begin VB.PictureBox picBtnR 
               Height          =   375
               Index           =   0
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   35
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
            Begin VB.PictureBox picBtnD 
               Height          =   375
               Index           =   0
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   34
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
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   0
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   33
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
            Begin VB.PictureBox picBtnES 
               Height          =   375
               Index           =   0
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   32
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
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   0
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   31
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
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   0
            Left            =   10440
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   29
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
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1680
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   3240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   2963
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
      Begin VB.Image Image3 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmBorrower.frx":7391
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borrowers Information"
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
         TabIndex        =   47
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of Borrowers and their other Information"
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
         TabIndex        =   46
         Top             =   480
         Width           =   4170
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
         Left            =   480
         TabIndex        =   45
         Top             =   5160
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
         Width           =   11415
      End
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   2
      Left            =   7320
      TabIndex        =   48
      Top             =   5760
      Visible         =   0   'False
      Width           =   6495
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   2
         Left            =   2880
         ScaleHeight     =   345
         ScaleWidth      =   3435
         TabIndex        =   65
         Top             =   2520
         Width           =   3435
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   2
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "View All Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   2
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   2
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   2
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   2
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   2
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   66
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
            Index           =   2
            Left            =   -480
            TabIndex        =   73
            Top             =   45
            Width           =   2535
         End
      End
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   2
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   5535
         TabIndex        =   49
         Top             =   840
         Width           =   5535
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   2
            Left            =   5160
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   63
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   2
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   50
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   2
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   61
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  ToolTipText     =   "Print"
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
               TabIndex        =   59
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  ToolTipText     =   "Refresh"
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
               TabIndex        =   57
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   2
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   55
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  ToolTipText     =   "Search"
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
               TabIndex        =   53
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   2
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   51
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   2
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   52
                  ToolTipText     =   "Insert Author"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   74
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
         TabIndex        =   77
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   2
         Left            =   240
         Picture         =   "frmBorrower.frx":1DD43
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books been Returned"
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
         TabIndex        =   76
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of Books that been borrowed by the Borrower and status of the borrowed Book."
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
         TabIndex        =   75
         Top             =   480
         Width           =   7020
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   2
         Left            =   120
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   3
      Left            =   6840
      TabIndex        =   113
      Top             =   5760
      Visible         =   0   'False
      Width           =   6495
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   3
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   5535
         TabIndex        =   123
         Top             =   840
         Width           =   5535
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   126
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   3
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   137
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   138
                  ToolTipText     =   "Insert Author"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnES 
               Height          =   375
               Index           =   3
               Left            =   360
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   135
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   136
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   3
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   133
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   134
                  ToolTipText     =   "Search"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnD 
               Height          =   375
               Index           =   3
               Left            =   1080
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   131
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   132
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnR 
               Height          =   375
               Index           =   3
               Left            =   1440
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   129
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   130
                  ToolTipText     =   "Refresh"
                  Top             =   0
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   3
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   127
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   3
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   128
                  ToolTipText     =   "Print"
                  Top             =   0
                  Width           =   315
               End
            End
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   3
            Left            =   5160
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   124
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   3
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   125
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   3
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   3465
         TabIndex        =   114
         Top             =   2520
         Width           =   3465
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   3
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Previous 250"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   3
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   3
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   3
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   3
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "View New Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   3
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   3
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   115
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
            Index           =   3
            Left            =   -480
            TabIndex        =   122
            Top             =   45
            Width           =   2535
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   3
         Left            =   120
         TabIndex        =   139
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
         Caption         =   "You can view List of Borrowers Log."
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
         Index           =   3
         Left            =   960
         TabIndex        =   142
         Top             =   480
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Selected Borrowers Log"
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
         Index           =   4
         Left            =   960
         TabIndex        =   141
         Top             =   240
         Width           =   2835
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   3
         Left            =   240
         Picture         =   "frmBorrower.frx":346F5
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
         Index           =   3
         Left            =   120
         TabIndex        =   140
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Shape spMag 
         BackColor       =   &H000AA27C&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Index           =   3
         Left            =   120
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "Borrower's Log"
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
      Index           =   2
      Left            =   4800
      TabIndex        =   112
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "View Book(s) Retuned"
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
      Index           =   1
      Left            =   2520
      TabIndex        =   111
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "View Book(s) Borrowed"
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
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   8880
      Width           =   2175
   End
   Begin VB.PictureBox picConn 
      Height          =   135
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   11475
      TabIndex        =   25
      Top             =   5640
      Width           =   11535
   End
   Begin VB.Frame fraList 
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   78
      Top             =   5760
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox picData 
         BorderStyle     =   0  'None
         Height          =   340
         Index           =   1
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   6375
         TabIndex        =   81
         Top             =   840
         Width           =   6375
         Begin VB.PictureBox pcBtns 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2295
            TabIndex        =   83
            Top             =   0
            Width           =   2295
            Begin VB.PictureBox picBtnP 
               Height          =   375
               Index           =   1
               Left            =   1800
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   89
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnP 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  ToolTipText     =   "Print"
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
               TabIndex        =   88
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnR 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  ToolTipText     =   "Refresh"
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
               TabIndex        =   87
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnD 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   "Delete Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnS 
               Height          =   375
               Index           =   1
               Left            =   720
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   86
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnS 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  ToolTipText     =   "Search"
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
               TabIndex        =   85
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnES 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Edit Selected"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
            Begin VB.PictureBox picBtnCN 
               Height          =   375
               Index           =   1
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   84
               Top             =   0
               Width           =   375
               Begin VB.CommandButton btnCN 
                  Height          =   315
                  Index           =   1
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  ToolTipText     =   "Create New"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   315
               End
            End
         End
         Begin VB.PictureBox picBtnC 
            Height          =   375
            Index           =   1
            Left            =   5880
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   82
            Top             =   0
            Width           =   375
            Begin VB.CommandButton btnC 
               Height          =   315
               Index           =   1
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "First"
               Top             =   0
               Width           =   315
            End
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   16
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
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   350
         Index           =   1
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   3495
         TabIndex        =   79
         Top             =   2520
         Width           =   3495
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Index           =   1
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Previous 250"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Index           =   1
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "First"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Index           =   1
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Last"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Index           =   1
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Next"
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton btnNew 
            Height          =   315
            Index           =   1
            Left            =   3520
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "View New Records"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnEdited 
            Height          =   315
            Index           =   1
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "View All Updated Records"
            Top             =   0
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton btnAll 
            Height          =   315
            Index           =   1
            Left            =   4160
            Style           =   1  'Graphical
            TabIndex        =   23
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
            Index           =   1
            Left            =   -480
            TabIndex        =   80
            Top             =   45
            Width           =   2535
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can view list of Books that been borrowed by the Borrower and status of the borrowed Book."
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
         TabIndex        =   92
         Top             =   480
         Width           =   7020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Books been Borrowed"
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
         TabIndex        =   91
         Top             =   240
         Width           =   2700
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   1
         Left            =   240
         Picture         =   "frmBorrower.frx":3A86F
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
         TabIndex        =   90
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
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmBorrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim view_other As Boolean
Dim INT_SIZE As Integer
Dim int_size_active As Integer
Dim CURR_COL As Integer
Dim iStartPage(3) As Long
Dim iNoPage(3) As Integer
Dim iRec(3) As Long
Dim sSQL(3) As String
Dim iList As Integer
Dim sColumns(3) As String, sColWidth(3) As String, sFields(3) As String
Dim iIcon(3) As Integer, iLoop(3) As Integer, sNoRec(3) As String

Dim bProfile As Boolean
Public iLvIndex As Integer, iDragSize As Integer
Public sSearchFields As String

Private Sub btnAll_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender " & _
            "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Borrower(s) Records.")
        If iRec(0) > 0 Then
            cmdOther(0).Enabled = True
            cmdOther(1).Enabled = True
            cmdOther(2).Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdOther(0).Enabled = False
            cmdOther(1).Enabled = False
            cmdOther(2).Enabled = False
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
        Set dtrBorrower.DataSource = rsView1
        dtrBorrower.Show 1
    ElseIf Index = 1 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrBorrowed.Sections("Section2")
            .Controls("lblid").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblname").Caption = lvList(0).SelectedItem.SubItems(1) & " " & Left(lvList(0).SelectedItem.SubItems(2), 1) & ". " & lvList(0).SelectedItem.SubItems(3)
            .Controls("lblBtype").Caption = lvList(0).SelectedItem.SubItems(5)
            .Controls("lblAdd").Caption = lvList(0).SelectedItem.SubItems(6)
            .Controls("lblBday").Caption = lvList(0).SelectedItem.SubItems(9)
            .Controls("lblGender").Caption = lvList(0).SelectedItem.SubItems(4)
            .Controls("lblTel").Caption = lvList(0).SelectedItem.SubItems(7)
            .Controls("lblCel").Caption = lvList(0).SelectedItem.SubItems(8)
        End With
        Set dtrBorrowed.DataSource = rsView1
        dtrBorrowed.Show 1
    ElseIf Index = 2 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrReturned.Sections("Section2")
            .Controls("lblid").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblname").Caption = lvList(0).SelectedItem.SubItems(1) & " " & Left(lvList(0).SelectedItem.SubItems(2), 1) & ". " & lvList(0).SelectedItem.SubItems(3)
            .Controls("lblBtype").Caption = lvList(0).SelectedItem.SubItems(5)
            .Controls("lblAdd").Caption = lvList(0).SelectedItem.SubItems(6)
            .Controls("lblBday").Caption = lvList(0).SelectedItem.SubItems(9)
            .Controls("lblGender").Caption = lvList(0).SelectedItem.SubItems(4)
            .Controls("lblTel").Caption = lvList(0).SelectedItem.SubItems(7)
            .Controls("lblCel").Caption = lvList(0).SelectedItem.SubItems(8)
        End With
        Set dtrReturned.DataSource = rsView1
        dtrReturned.Show 1
    ElseIf Index = 3 Then
        rsView1.Open sSQL(Index), adoCon, 3, 3
        With dtrLoged.Sections("Section2")
            .Controls("lblid").Caption = lvList(0).SelectedItem.SubItems(1)
            .Controls("lblname").Caption = lvList(0).SelectedItem.SubItems(1) & " " & Left(lvList(0).SelectedItem.SubItems(2), 1) & ". " & lvList(0).SelectedItem.SubItems(3)
            .Controls("lblBtype").Caption = lvList(0).SelectedItem.SubItems(5)
            .Controls("lblAdd").Caption = lvList(0).SelectedItem.SubItems(6)
            .Controls("lblBday").Caption = lvList(0).SelectedItem.SubItems(9)
            .Controls("lblGender").Caption = lvList(0).SelectedItem.SubItems(4)
            .Controls("lblTel").Caption = lvList(0).SelectedItem.SubItems(7)
            .Controls("lblCel").Caption = lvList(0).SelectedItem.SubItems(8)
        End With
        Set dtrLoged.DataSource = rsView1
        dtrLoged.Show 1
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

Private Sub cmdOther_Click(Index As Integer)
    If Index = 0 Then
        If view_other = True And lvList(1).Visible = True Then
            fraList(1).Visible = False
            picConn.Visible = False
            view_other = False
        Else
            sSQL(1) = "SELECT tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
                "FROM tbl_books INNER JOIN (tbl_borrow_record INNER JOIN tbl_reg_books ON tbl_borrow_record.rb_id = tbl_reg_books.rb_id) ON tbl_books.isbn = tbl_reg_books.isbn " & _
                "WHERE (((tbl_borrow_record.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
                "GROUP BY tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
            Lv_OtherInfo
            fraList(1).Visible = True
            picConn.Visible = True
            view_other = True
            fraList(2).Visible = False
            fraList(3).Visible = False
        End If
    ElseIf Index = 1 Then
        If view_other = True And lvList(2).Visible = True Then
            fraList(2).Visible = False
            picConn.Visible = False
            view_other = False
        Else
            sSQL(2) = "SELECT tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
                "FROM tbl_books INNER JOIN (tbl_borrow_record INNER JOIN tbl_reg_books ON tbl_borrow_record.rb_id = tbl_reg_books.rb_id) ON tbl_books.isbn = tbl_reg_books.isbn " & _
                "WHERE (((tbl_borrow_record.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_borrow_record.s_return) Like '1')) " & _
                "GROUP BY tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
            Lv_OtherInfo2
            fraList(2).Visible = True
            picConn.Visible = True
            view_other = True
            fraList(1).Visible = False
            fraList(3).Visible = False
        End If
    ElseIf Index = 2 Then
        If view_other = True And lvList(3).Visible = True Then
            fraList(3).Visible = False
            picConn.Visible = False
            view_other = False
        Else
            sSQL(3) = "SELECT tbl_borrower_log.log_id, tbl_borrower_log.logtime, tbl_borrower_log.logdate " & _
                    "From tbl_borrower_log " & _
                    "WHERE (((tbl_borrower_log.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                    "GROUP BY tbl_borrower_log.log_id, tbl_borrower_log.logtime, tbl_borrower_log.logdate;"
            Lv_OtherInfo3
            fraList(3).Visible = True
            picConn.Visible = True
            view_other = True
            fraList(1).Visible = False
            fraList(2).Visible = False
        End If
    End If

    INT_SIZE = Me.ScaleHeight / 2
    Listview_Resize
End Sub

Private Sub cmdProfile_Click()
    If bProfile = False And iRec(0) > 0 Then
        bProfile = True
        fraProfile.Visible = True
        cmdProfile.Picture = frmMain.iPageEnabled.ListImages(9).ExtractIcon
        GetBorrowerProfile True
        iDragSize = 4200
        Listview_Resize
    Else
        bProfile = False
        GetBorrowerProfile False
        fraProfile.Visible = False
        cmdProfile.Picture = frmMain.iPageEnabled.ListImages(8).ExtractIcon
        iDragSize = 2600
        Listview_Resize
    End If
    lvList(0).SetFocus
End Sub

Private Sub Form_Load()
    view_other = False
    INT_SIZE = Me.ScaleHeight / 2
    Set_Icon_btn Me, 3
    fraList(1).Visible = False
    picConn.Visible = False
    sSQL(0) = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender " & _
            "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id;"
    Lv_MainInfo
    lvList(0).Refresh
    frmMain.TabMainIni 2, "Borrowers", 12
    bProfile = False
    iDragSize = 2600
    cmdProfile.Picture = frmMain.iPageEnabled.ListImages(8).ExtractIcon
    iList = 1
End Sub

Private Sub Form_LostFocus()
    If frmFind.Visible = True Then
        frmFind.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call Listview_Resize
    cmdOther(0).Top = Me.ScaleHeight - (cmdOther(0).Height + 120)
    cmdOther(1).Top = Me.ScaleHeight - (cmdOther(1).Height + 120)
    cmdOther(2).Top = Me.ScaleHeight - (cmdOther(1).Height + 120)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.TabMainIni 1, "Borrowers", 12
    bProfile = False
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
            frmBorrower_AE.bStat = False
            Set frmBorrower_AE.fCur = Me
            frmBorrower_AE.Show 1
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
        If bProfile = True Then
            If iRec(0) > 0 Then
                GetBorrowerProfile True
            Else
                GetBorrowerProfile False
            End If
        End If
        If lvList(1).Visible = True Then
            LvRefresh 1
        ElseIf lvList(2).Visible = True Then
            LvRefresh 2
        ElseIf lvList(3).Visible = True Then
            LvRefresh 3
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
            If picConn.Top > (fraList(0).Top + iDragSize) Then
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
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, Me.ScaleHeight - (cmdOther(0).Height + 180)
        If bProfile = True Then
            fraProfile.Move 480, 1200, fraList(i).Width - (480 + 120)
            lvList(i).Move 480, 3240, fraList(i).Width - (480 + 120), fraList(i).Height - (3240 + 240 + 240 + 120)
        Else
            lvList(i).Move 480, 1320, fraList(i).Width - (480 + 120), fraList(i).Height - (1320 + 240 + 240 + 120)
        End If
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 480, fraList(i).Height - (350 + 90)
        lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 480, 840, fraList(i).Width - (120 + 480)
        picBtnC(i).Left = picData(i).Width - (picBtnC(i).Width)
    Else
        picConn.Width = Me.ScaleWidth - 240
        picConn.Top = INT_SIZE
        'this will use for listview(0)
        i = 0
        fraList(i).Move 120, 0, Me.ScaleWidth - 240, picConn.Top - (105)
        If bProfile = True Then
            fraProfile.Move 480, 1200, fraList(i).Width - (480 + 120)
            lvList(i).Move 480, 3240, fraList(i).Width - (480 + 120), fraList(i).Height - (3240 + 240 + 240 + 120)
        Else
            lvList(i).Move 480, 1320, fraList(i).Width - (480 + 120), fraList(i).Height - (1320 + 240 + 240 + 120)
        End If
        spMag(i).Move 0, 240, fraList(i).Width
        picPage(i).Move fraList(i).Width - (picPage(i).Width + 120), fraList(i).Height - (350 + 120)
        lblSelected(i).Move 120, fraList(i).Height - (350 + 120)
        lvList(i).SelectedItem.EnsureVisible
        picData(i).Move 480, 840, fraList(i).Width - (120 + 480)
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
        i = 3
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
    sColWidth(0) = "1700,1300,1300,1300,1300,1300,2500,1300,1300,1300"
    sColumns(0) = "Borrower ID,First Name,Middle Name,Last Name,Gender,B.Type,Address,Tel. No.,Cell No.,Birthdate"
    iIcon(0) = 12
    iLoop(0) = CountSplitItem(sColumns(0), ",")
    sFields(0) = "B_id,fn,mn,ln,gender,b_type,add,tel,cel,bday"
    sNoRec(0) = "No Current Borrowers Records."
    iStartPage(0) = 1
    iNoPage(0) = 75
    iRec(0) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), sNoRec(0))
    If iRec(0) > 0 Then
        cmdOther(0).Enabled = True
        cmdOther(1).Enabled = True
    Else
        view_other = False
        picConn.Visible = False
        Listview_Resize
        cmdOther(0).Enabled = False
        cmdOther(1).Enabled = False
    End If
End Sub

'This Function is used to Refresh the Info on Listview(1) Index
Public Sub Lv_OtherInfo()
    sColWidth(1) = "2000,1400,4000,1200,1200,1250,1200,1250"
    sColumns(1) = "ID,ISBN,Title,Borrow Date,Return Date,Date Returned,Returned?,Penalty Day(s)"
    sFields(1) = "br_id,isbn,title,b_date,r_date,datereturned,s_return,d_penalty"
    iIcon(1) = 2
    iLoop(1) = CountSplitItem(sColumns(1), ",") 'this is count of columns to loop
    sNoRec(1) = "No Current Borrowed Books."
    iStartPage(1) = 1
    iNoPage(1) = 75
    iRec(1) = LvPageStat(Me, 1, sSQL(1), iStartPage(1), iNoPage(1), iIcon(1), sColumns(1), iLoop(1), sColWidth(1), sFields(1), sNoRec(1))
    lvList(1).Refresh
End Sub

'This function is used to Refresh the Info on Listview(2) Index
Public Sub Lv_OtherInfo2()
    sColWidth(2) = "2000,1400,4000,1200,1200,1250,1200,1250"
    sColumns(2) = "ID,ISBN,Title,Borrow Date,Return Date,Date Returned,Returned?,Penalty Day(s)"
    sFields(2) = "br_id,isbn,title,b_date,r_date,datereturned,s_return,d_penalty"
    iIcon(2) = 14
    iLoop(2) = CountSplitItem(sColumns(2), ",") 'this is count of columns to loop
    sNoRec(2) = "No Current Returned Books."
    iStartPage(2) = 1
    iNoPage(2) = 75
    iRec(2) = LvPageStat(Me, 2, sSQL(2), iStartPage(2), iNoPage(2), iIcon(2), sColumns(2), iLoop(2), sColWidth(2), sFields(2), sNoRec(2))
    lvList(2).Refresh
End Sub

Public Sub Lv_OtherInfo3()
    sColWidth(3) = "1700,2000,2000"
    sColumns(3) = "Log ID,Log Time,Log Date"
    sFields(3) = "log_id,logtime,logdate"
    iIcon(3) = 16
    iLoop(3) = CountSplitItem(sColumns(3), ",") 'this is count of columns to loop
    sNoRec(3) = "No Current Borrower's Log Records."
    iStartPage(3) = 1
    iNoPage(3) = 75
    iRec(3) = LvPageStat(Me, 3, sSQL(3), iStartPage(3), iNoPage(3), iIcon(3), sColumns(3), iLoop(3), sColWidth(3), sFields(3), sNoRec(3))
    lvList(3).Refresh
End Sub

'This will be used to Control the Update of Information
Private Sub btnEdited_Click(Index As Integer)
    If Index = 0 Then
        sSQL(Index) = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender " & _
                    "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
                    "WHERE (((tbl_borrowers.DateModified) Between #0" & Date & "# And #" & Date & "#)) " & _
                    "GROUP BY tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, Index, sSQL(Index), iStartPage(Index), iNoPage(Index), iIcon(Index), sColumns(Index), iLoop(Index), sColWidth(Index), sFields(Index), "No Current Updated Borrowers Info Records.")
        
        If bProfile = True Then
            GetBorrowerProfile True
        End If
        
        If iRec(0) > 0 Then
            cmdOther(0).Enabled = True
            cmdOther(1).Enabled = True
        Else
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdOther(0).Enabled = False
            cmdOther(1).Enabled = False
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
        sSQL(Index) = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender " & _
                    "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
                    "WHERE (((tbl_borrowers.DateAdded) Between #0" & Date & "# And #" & Date & "#)) " & _
                    "GROUP BY tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender;"
        iStartPage(Index) = 1
        iRec(Index) = LvPageStat(Me, 0, sSQL(0), iStartPage(0), iNoPage(0), iIcon(0), sColumns(0), iLoop(0), sColWidth(0), sFields(0), "No Current New Borrowers Info Added Records.")

        If iRec(0) > 0 Then
            If bProfile = True Then
                GetBorrowerProfile True
            End If
            cmdOther(0).Enabled = True
            cmdOther(1).Enabled = True
        Else
            If bProfile = True Then
                GetBorrowerProfile False
            End If
            view_other = False
            picConn.Visible = False
            Listview_Resize
            cmdOther(0).Enabled = False
            cmdOther(1).Enabled = False
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
        frmBorrower_AE.bStat = True
        Set frmBorrower_AE.fCur = Me
        frmBorrower_AE.Show 1
    End If
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
            sMsgId = "Borrower ID"
            sMsgFooter = Chr$(13) & Chr$(13) & "If you click Yes, you won't be able to undo the deletion."
            sMsgDel = sMsgDel & Chr$(13) & Chr$(13) & sMsgId & ": " & lvList(Index).SelectedItem.SubItems(1) & sMsgFooter
            If MsgBox(sMsgDel, vbQuestion + vbYesNo, "Delete") = vbYes Then
                DELETE_DATA "tbl_borrowers", "B_id", lvList(Index).SelectedItem.SubItems(1)
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
        frmBorrower_AE.bStat = False
        Set frmBorrower_AE.fCur = Me
        frmBorrower_AE.Show 1
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
        
        If bProfile = True Then
            GetBorrowerProfile True
        End If
        
        If lvList(1).Visible = True Then
            LvRefresh 1
        ElseIf lvList(2).Visible = True Then
            LvRefresh 2
        End If
    ElseIf Index = 1 Then
        sSQL(1) = "SELECT tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
            "FROM tbl_books INNER JOIN (tbl_borrow_record INNER JOIN tbl_reg_books ON tbl_borrow_record.rb_id = tbl_reg_books.rb_id) ON tbl_books.isbn = tbl_reg_books.isbn " & _
            "WHERE (((tbl_borrow_record.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_borrow_record.s_return) Like '0')) " & _
            "GROUP BY tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
        Lv_OtherInfo
    ElseIf Index = 2 Then
        sSQL(2) = "SELECT tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty " & _
            "FROM tbl_books INNER JOIN (tbl_borrow_record INNER JOIN tbl_reg_books ON tbl_borrow_record.rb_id = tbl_reg_books.rb_id) ON tbl_books.isbn = tbl_reg_books.isbn " & _
            "WHERE (((tbl_borrow_record.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "') AND ((tbl_borrow_record.s_return) Like '1')) " & _
            "GROUP BY tbl_borrow_record.br_id, tbl_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date, tbl_borrow_record.datereturned, tbl_borrow_record.s_return, tbl_borrow_record.d_penalty;"
        Lv_OtherInfo2
    ElseIf Index = 3 Then
        sSQL(3) = "SELECT tbl_borrower_log.log_id, tbl_borrower_log.logtime, tbl_borrower_log.logdate " & _
                    "From tbl_borrower_log " & _
                    "WHERE (((tbl_borrower_log.B_id) Like '" & lvList(0).SelectedItem.SubItems(1) & "')) " & _
                    "GROUP BY tbl_borrower_log.log_id, tbl_borrower_log.logtime, tbl_borrower_log.logdate;"
        Lv_OtherInfo3
    End If
End Function

Public Function GetBorrowerProfile(bVisible As Boolean)
    On Error GoTo errHandler
    If bVisible = True Then
        lblID.Caption = lvList(0).SelectedItem.SubItems(1)
        lblName.Caption = lvList(0).SelectedItem.SubItems(2) & " " & Left(lvList(0).SelectedItem.SubItems(3), 1) & ". " & lvList(0).SelectedItem.SubItems(4)
        lblBday.Caption = Format(lvList(0).SelectedItem.SubItems(10), "dddddd")
        lblAdd.Caption = lvList(0).SelectedItem.SubItems(7)
        lblGen.Caption = lvList(0).SelectedItem.SubItems(5)
        lblBtype.Caption = lvList(0).SelectedItem.SubItems(6)
        lblTel.Caption = lvList(0).SelectedItem.SubItems(8)
        lblCel.Caption = lvList(0).SelectedItem.SubItems(9)
        ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & lvList(0).SelectedItem.SubItems(1) & ".jpg")
    Else
        lblID.Caption = ""
        lblName.Caption = ""
        lblBday.Caption = ""
        lblAdd.Caption = ""
        lblGen.Caption = ""
        lblBtype.Caption = ""
        lblTel.Caption = ""
        lblCel.Caption = ""
        ImgPic.Picture = LoadPicture("")
        bProfile = False
        fraProfile.Visible = False
        cmdProfile.Picture = frmMain.iPageEnabled.ListImages(8).ExtractIcon
        iDragSize = 2600
        Listview_Resize
    End If
errHandler:
    If err.Number = 53 Then ImgPic.Picture = LoadPicture(App.Path & "\_borrower pics\" & "nopic" & ".jpg")
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
    sSQL(0) = "SELECT tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender " & _
            "FROM tbl_borrower_type INNER JOIN tbl_borrowers ON tbl_borrower_type.bt_id = tbl_borrowers.bt_id " & _
            "WHERE " & sFilter & " " & _
            "GROUP BY tbl_borrowers.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_borrower_type.b_type, tbl_borrowers.add, tbl_borrowers.tel, tbl_borrowers.cel, tbl_borrowers.bday, tbl_borrowers.gender;"
    LvRefresh 0
End Function

