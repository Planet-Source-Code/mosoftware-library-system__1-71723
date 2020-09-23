VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOPAC 
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   Begin MSComctlLib.ImageList i16x16 
      Left            =   3840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":169C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":2D384
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   3240
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":3350E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":4A5A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":50732
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":568BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPAC.frx":5CA46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   661
      ImageList       =   "i16x16"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Books"
            Key             =   "Book"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Periodicals"
            Key             =   "Periodicals"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraBook 
      BorderStyle     =   0  'None
      Height          =   8175
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   14775
      Begin VB.Frame Frame1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC3E0F&
         Height          =   1095
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   14775
         Begin VB.TextBox txtFind2 
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
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   15
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox txtFind1 
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
            Height          =   285
            Index           =   1
            Left            =   2760
            TabIndex        =   14
            Top             =   360
            Width           =   4815
         End
         Begin VB.ComboBox cboFind2 
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
            Height          =   315
            Index           =   1
            ItemData        =   "frmOPAC.frx":62BD0
            Left            =   1320
            List            =   "frmOPAC.frx":62BDD
            TabIndex        =   41
            Text            =   "Any"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cboFind1 
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
            Index           =   1
            ItemData        =   "frmOPAC.frx":62BF7
            Left            =   1320
            List            =   "frmOPAC.frx":62C04
            TabIndex        =   40
            Text            =   "Any"
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdRefresh 
            Height          =   495
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdFind 
            Height          =   495
            Index           =   1
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton cmdClear 
            Height          =   495
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Search Option"
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
            Left            =   1320
            TabIndex        =   43
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Words, Title of the Book or Author's name"
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
            Left            =   2760
            TabIndex        =   42
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label lblBtn 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.Image Image2 
            Height          =   720
            Index           =   1
            Left            =   13920
            Picture         =   "frmOPAC.frx":62C1E
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame fraDesc 
         Caption         =   "Ohers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC3E0F&
         Height          =   1815
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   6240
         Width           =   14775
         Begin VB.Label lblDesc 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Description"
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
            Height          =   315
            Left            =   2160
            TabIndex        =   47
            Top             =   1320
            Width           =   3045
         End
         Begin VB.Label lblDate 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date Publish"
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
            Height          =   315
            Left            =   2160
            TabIndex        =   46
            Top             =   960
            Width           =   3045
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "Date Publish"
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
            Height          =   195
            Index           =   2
            Left            =   960
            TabIndex        =   45
            Top             =   1005
            Width           =   1245
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808080&
            Caption         =   "Description"
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
            Height          =   195
            Index           =   2
            Left            =   960
            TabIndex        =   44
            Top             =   1365
            Width           =   1245
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808080&
            Caption         =   "Edition"
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
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   33
            Top             =   645
            Width           =   1245
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "Volume"
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
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   32
            Top             =   280
            Width           =   1245
         End
         Begin VB.Label lblVol 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Volume"
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
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label lblEdition 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Edition"
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
            Height          =   315
            Left            =   2160
            TabIndex        =   31
            Top             =   600
            Width           =   3045
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   1
            Left            =   120
            Picture         =   "frmOPAC.frx":69470
            Top             =   600
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   5040
         Index           =   1
         Left            =   0
         TabIndex        =   37
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8890
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
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
   Begin VB.Frame fraBook 
      BorderStyle     =   0  'None
      Height          =   8295
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   14775
      Begin VB.Frame fraDesc 
         Caption         =   "Ohers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC3E0F&
         Height          =   1815
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   6240
         Width           =   14775
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   5040
            TabIndex        =   38
            Top             =   120
            Width           =   9615
            Begin VB.CommandButton cmdClr 
               Height          =   315
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "View Profile"
               Top             =   420
               Width           =   315
            End
            Begin VB.CommandButton cmdPrint 
               Height          =   315
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "View Profile"
               Top             =   120
               Width           =   315
            End
            Begin MSComctlLib.ListView lvPrint 
               Height          =   1440
               Left            =   360
               TabIndex        =   10
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   2540
               View            =   3
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               PictureAlignment=   3
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
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "You can add and remove selected book also Print Book Information for Borrowing the Book."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   8160
               TabIndex        =   39
               Top             =   240
               Width           =   1335
            End
            Begin VB.Image Image3 
               Height          =   360
               Index           =   3
               Left            =   7680
               Picture         =   "frmOPAC.frx":6F5EA
               Stretch         =   -1  'True
               Top             =   240
               Width           =   360
            End
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   0
            Left            =   120
            Picture         =   "frmOPAC.frx":75E3C
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblAuthor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AUTHOR(S)"
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
            Index           =   0
            Left            =   2280
            TabIndex        =   7
            Top             =   960
            Width           =   2685
         End
         Begin VB.Label lblYB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   6
            Top             =   600
            Width           =   2685
         End
         Begin VB.Label lblPub 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   5
            Top             =   240
            Width           =   2685
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            Caption         =   "PUBLISHER"
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
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Top             =   285
            Width           =   1365
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808080&
            Caption         =   "YEAR PUBLISH"
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
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   27
            Top             =   645
            Width           =   1365
         End
         Begin VB.Label Label7 
            BackColor       =   &H00808080&
            Caption         =   "AUTHOR(S)"
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
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   26
            Top             =   1005
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC3E0F&
         Height          =   1095
         Index           =   0
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   14775
         Begin VB.CommandButton cmdRefresh 
            Height          =   495
            Index           =   0
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   495
         End
         Begin VB.ComboBox cboFind1 
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
            Index           =   0
            ItemData        =   "frmOPAC.frx":7BFB6
            Left            =   1320
            List            =   "frmOPAC.frx":7BFC3
            TabIndex        =   13
            Text            =   "Any"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cboFind2 
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
            Height          =   315
            Index           =   0
            ItemData        =   "frmOPAC.frx":7BFDB
            Left            =   1320
            List            =   "frmOPAC.frx":7BFE8
            TabIndex        =   1
            Text            =   "Any"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtFind1 
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
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   0
            Top             =   360
            Width           =   4815
         End
         Begin VB.TextBox txtFind2 
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
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   2
            Top             =   720
            Width           =   4815
         End
         Begin VB.CommandButton cmdClear 
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdFind 
            Default         =   -1  'True
            Height          =   495
            Index           =   0
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   400
            Width           =   495
         End
         Begin VB.Image Image2 
            Height          =   720
            Index           =   0
            Left            =   13920
            Picture         =   "frmOPAC.frx":7C000
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Words, Title of the Book or Author's name"
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
            Left            =   2760
            TabIndex        =   24
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Search Option"
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
            Left            =   1320
            TabIndex        =   23
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblBtn 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
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
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   5040
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8890
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
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
   Begin VB.Image Image4 
      Height          =   1140
      Left            =   120
      Picture         =   "frmOPAC.frx":82852
      Top             =   360
      Width           =   10590
   End
End
Attribute VB_Name = "frmOPAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String

Public Sub Lv_MainInfo()
    Dim mRow As ListItem
    On Error Resume Next
    lvList(0).ColumnHeaders.Clear
    lvList(0).ListItems.Clear
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        'MsgBox "Record Find: " & adoRes.RecordCount
        If adoRes.RecordCount > 0 Then
            lvList(0).ColumnHeaders.Add , , , 650
            lvList(0).ColumnHeaders.Add , , "Call No.", 1300
            lvList(0).ColumnHeaders.Add , , "ISBN", 1400
            lvList(0).ColumnHeaders.Add , , "Title", 8000
            lvList(0).ColumnHeaders.Add , , "Description", 4200
            lvList(0).ForeColor = vbBlack
            Do While Not adoRes.EOF
                Set mRow = lvList(0).ListItems.Add(, , , , 2)
                mRow.SubItems(1) = adoRes.Fields("acronym") & "-" & adoRes.Fields("callno")
                mRow.SubItems(2) = adoRes.Fields("isbn")
                mRow.SubItems(3) = adoRes.Fields("title")
                mRow.SubItems(4) = adoRes.Fields("desc")
                adoRes.MoveNext
            Loop
        Else
            lvList(0).ColumnHeaders.Add , , "", 8000
            lvList(0).ListItems.Add , , "No Current Record Found.", , 1
            lvList(0).SelectedItem.ForeColor = vbRed
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    CheckRecordsQty
End Sub

Public Sub Lv_MainInfo2()
    Dim mRow As ListItem
    On Error Resume Next
    lvList(1).ColumnHeaders.Clear
    lvList(1).ListItems.Clear
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        'MsgBox adoRes.RecordCount
        If adoRes.RecordCount > 0 Then
            lvList(1).ColumnHeaders.Add , , , 300
            lvList(1).ColumnHeaders.Add , , "ISSN", 1400
            lvList(1).ColumnHeaders.Add , , "Title", 6000
            lvList(1).ColumnHeaders.Add , , "Article", 2500
            lvList(1).ColumnHeaders.Add , , "Author(s)", 2000
            lvList(1).ColumnHeaders.Add , , "Periodic Type", 1500
            lvList(1).ForeColor = vbBlack
            Do While Not adoRes.EOF
                Set mRow = lvList(1).ListItems.Add(, , , , 6)
                mRow.SubItems(1) = adoRes.Fields("issn")
                mRow.SubItems(2) = adoRes.Fields("title")
                mRow.SubItems(3) = adoRes.Fields("article")
                mRow.SubItems(4) = adoRes.Fields("authors")
                mRow.SubItems(5) = adoRes.Fields("s_type")
                adoRes.MoveNext
            Loop
        Else
            lvList(1).ColumnHeaders.Add , , "", 8000
            lvList(1).ListItems.Add , , "No Current Record Found.", , 1
            lvList(1).SelectedItem.ForeColor = vbRed
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Sub

Public Function SqlConstructor() As String

End Function

Private Sub cmdClr_Click()
    lvPrint.ListItems.Clear
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Dim sWhereFinal As String, sItemwoOr As String
    sWhereFinal = ""
    If cboFind1(Index).ListIndex = 0 And Len(Trim(txtFind1(Index).Text)) > 0 Then
        sWhereFinal = sWhereFinal & WhereCreator(1, Trim(txtFind1(Index).Text), Index) & " OR " & sWhereFinal & WhereCreator(2, Trim(txtFind1(Index).Text), Index)
    ElseIf cboFind1(Index).ListIndex = 1 Then
        sWhereFinal = sWhereFinal & WhereCreator(1, Trim(txtFind1(Index).Text), Index)
    ElseIf cboFind1(Index).ListIndex = 2 Then
        sWhereFinal = sWhereFinal & WhereCreator(2, Trim(txtFind1(Index).Text), Index)
    End If
    
    If Len(Trim(sWhereFinal)) > 0 And Len(txtFind2(0).Text) > 0 Then
        sWhereFinal = sWhereFinal & " OR "
    End If
    
    If cboFind2(Index).ListIndex = 0 And Len(Trim(txtFind2(Index).Text)) > 0 Then
        sWhereFinal = sWhereFinal & WhereCreator(1, Trim(txtFind2(Index).Text), Index) & " OR " & sWhereFinal & WhereCreator(2, Trim(txtFind2(Index).Text), Index)
    ElseIf cboFind2(Index).ListIndex = 1 Then
        sWhereFinal = sWhereFinal & WhereCreator(1, Trim(txtFind2(Index).Text), Index)
    ElseIf cboFind2(Index).ListIndex = 2 Then
        sWhereFinal = sWhereFinal & WhereCreator(2, Trim(txtFind2(Index).Text), Index)
    End If
    sItemwoOr = Replace(Trim(sWhereFinal), "OR", "")
    If Index = 0 Then
        If Len(Trim(sItemwoOr)) > 0 Then
            sSQL = "SELECT tbl_shelfs.acronym, tbl_shelfbooks.callno, tbl_books.isbn, tbl_books.title, tbl_books.desc " & _
                "FROM tbl_shelfs INNER JOIN (((tbl_books INNER JOIN (tbl_authors INNER JOIN tbl_bookauthor ON tbl_authors.auid = tbl_bookauthor.auid) ON tbl_books.isbn = tbl_bookauthor.isbn) INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN tbl_shelfbooks ON tbl_reg_books.rb_id = tbl_shelfbooks.rb_id) ON tbl_shelfs.sh_id = tbl_shelfbooks.sh_id " & _
                "Where " & sWhereFinal & " " & _
                "GROUP BY tbl_shelfs.acronym, tbl_shelfbooks.callno, tbl_books.isbn, tbl_books.title, tbl_books.desc;"
            Lv_MainInfo
        Else
            lvList(0).ColumnHeaders.Clear
            lvList(0).ListItems.Clear
            lvList(0).ColumnHeaders.Add , , "", 8000
            lvList(0).ListItems.Add , , "No Current Record Found.", , 1
            lvList(0).SelectedItem.ForeColor = vbRed
        End If
    ElseIf Index = 1 Then
        If Len(Trim(sItemwoOr)) > 0 Then
            sSQL = "SELECT tbl_magazines.issn, tbl_magazines.title, tbl_reg_magazines.article, tbl_reg_magazines.authors, tbl_magazines.s_type " & _
                "FROM tbl_magazines INNER JOIN tbl_reg_magazines ON tbl_magazines.issn = tbl_reg_magazines.issn " & _
                "Where " & sWhereFinal & " " & _
                "GROUP BY tbl_magazines.issn, tbl_magazines.title, tbl_reg_magazines.article, tbl_reg_magazines.authors, tbl_magazines.s_type;"
            Lv_MainInfo2
        Else
            lvList(1).ColumnHeaders.Clear
            lvList(1).ListItems.Clear
            lvList(1).ColumnHeaders.Add , , "", 8000
            lvList(1).ListItems.Add , , "No Current Record Found.", , 1
            lvList(1).SelectedItem.ForeColor = vbRed
        End If
    End If
End Sub

Public Function SearchItem()
    Dim sWhere As String
    If Index = 0 Then
        If Len(Trim(txtFind1(0).Text)) > 0 Or Len(Trim(txtFind2(0).Text)) > 0 Then
        
        Else
            MsgBox "You forgot to Enter Book Title that you wanted to find.", vbExclamation, "NulLPointerException"
            txtFind1(0).SetFocus
        End If
    ElseIf Index = 1 Then
        If Len(Trim(txtFind1(1).Text)) > 0 Or Len(Trim(txtFind2(1).Text)) > 0 Then
        sSQL = "SELECT tbl_magazines.issn, tbl_magazines.title, tbl_magazines.comp, tbl_magazines.dsc " & _
            "FROM tbl_magazines INNER JOIN tbl_reg_magazines ON tbl_magazines.issn = tbl_reg_magazines.issn " & _
            "WHERE (((tbl_magazines.title) Like '')) " & _
            "GROUP BY tbl_magazines.issn, tbl_magazines.title, tbl_magazines.comp, tbl_magazines.dsc;"
            Lv_MainInfo2
        Else
            MsgBox "You forgot to Enter Book Title that you wanted to find.", vbExclamation, "NulLPointerException"
            txtFind1(0).SetFocus
        End If
    End If
End Function

Private Sub cmdPrint_Click()
    Dim a As VbMsgBoxResult
    If lvPrint.ListItems.Count > 0 Then
        a = MsgBox("Do you want to Print you're Selected Book?", vbQuestion + vbYesNo, "Print")
        If a = vbYes Then
            GeneratePrint
        End If
    End If
End Sub

Private Sub Form_Load()
    Set lvList(0).SmallIcons = frmMain.iLv
    Set lvList(1).SmallIcons = frmMain.iLv
    SetButtonsPic
    SetListview
    SetButtonPic
    cboFind1(0).ListIndex = 0
    cboFind2(0).ListIndex = 0
    cboFind1(1).ListIndex = 0
    cboFind2(1).ListIndex = 0
End Sub

Public Function SetButtonsPic()
    cmdClear(0).Picture = i32x32.ListImages(1).ExtractIcon
    cmdRefresh(0).Picture = i32x32.ListImages(2).ExtractIcon
    cmdFind(0).Picture = i32x32.ListImages(3).ExtractIcon
    
    cmdClear(1).Picture = i32x32.ListImages(1).ExtractIcon
    cmdRefresh(1).Picture = i32x32.ListImages(2).ExtractIcon
    cmdFind(1).Picture = i32x32.ListImages(3).ExtractIcon
End Function

Public Function GetBookOtherInfo()
    sSQL = "SELECT tbl_publishers.cmpny, tbl_books.yrpub " & _
        "FROM tbl_publishers INNER JOIN tbl_books ON tbl_publishers.pubid = tbl_books.pubid " & _
        "WHERE (((tbl_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(2) & "')) " & _
        "GROUP BY tbl_publishers.cmpny, tbl_books.yrpub;"
    lblPub(0).Caption = FindFieldValue(sSQL, "cmpny")
    lblYB(0).Caption = FindFieldValue(sSQL, "yrpub")
    lblAuthor(0).Caption = GetBookAuthors
End Function

Public Function GetOtherInfoPeriodicals()
    sSQL = "SELECT tbl_magazines.dsc, tbl_magazines.s_type, tbl_magazines.s_vol, tbl_magazines.s_edition, tbl_magazines.d_publish " & _
        "From tbl_magazines " & _
        "WHERE (((tbl_magazines.issn) Like '" & lvList(1).SelectedItem.SubItems(1) & "')) " & _
        "GROUP BY tbl_magazines.dsc, tbl_magazines.s_type, tbl_magazines.s_vol, tbl_magazines.s_edition, tbl_magazines.d_publish;"
    lblVol.Caption = FindFieldValue(sSQL, "s_vol")
    lblEdition.Caption = FindFieldValue(sSQL, "s_edition")
    lblDate.Caption = FindFieldValue(sSQL, "d_publish")
    lblDesc.Caption = FindFieldValue(sSQL, "dsc")
End Function

Public Function GetBookAuthors() As String
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    sSQL = "SELECT tbl_authors.author " & _
        "FROM tbl_authors INNER JOIN (tbl_books INNER JOIN tbl_bookauthor ON tbl_books.isbn = tbl_bookauthor.isbn) ON tbl_authors.auid = tbl_bookauthor.auid " & _
        "WHERE (((tbl_books.isbn) Like '" & lvList(0).SelectedItem.SubItems(2) & "')) " & _
        "GROUP BY tbl_authors.author;"
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        Do While Not adoRes.EOF
            If Not adoRes.EOF Then
                GetBookAuthors = GetBookAuthors & adoRes.Fields("author") & ","
            Else
                GetBookAuthors = GetBookAuthors & adoRes.Fields("author")
            End If
            adoRes.MoveNext
        Loop
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Function

Private Sub lvList_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Index = 0 Then
        GetBookOtherInfo
    ElseIf Index = 1 Then
        GetOtherInfoPeriodicals
    End If
End Sub

Private Sub lvList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Index = 0 Then
        frmMain.mnuII.Visible = True
        frmMain.mnuRI.Visible = False
        PopupMenu frmMain.mnuSrch
    End If
End Sub

Private Sub lvPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        frmMain.mnuRI.Visible = True
        frmMain.mnuII.Visible = False
        PopupMenu frmMain.mnuSrch
    End If
End Sub

Private Sub tbSearch_Click()
    Select Case tbSearch.SelectedItem.Caption
        Case "Search Books":
            fraBook(0).Visible = True
            fraBook(1).Visible = False
            cmdFind(0).Default = True
        Case "Search Periodicals":
            fraBook(1).Visible = True
            fraBook(0).Visible = False
            cmdFind(1).Default = True
    End Select
End Sub

Public Function SetButtonPic()
    cmdClr.Picture = frmMain.i16x16.ListImages(4).ExtractIcon
    cmdPrint.Picture = frmMain.i16x16.ListImages(6).ExtractIcon
End Function

Public Function SetListview()
    Set lvPrint.SmallIcons = frmMain.iLv
    lvPrint.ColumnHeaders.Clear
    lvPrint.ListItems.Clear
    lvPrint.ColumnHeaders.Add , , , 300
    lvPrint.ColumnHeaders.Add , , "Call No.", 1500
    lvPrint.ColumnHeaders.Add , , "ISBN", 1500
    lvPrint.ColumnHeaders.Add , , "Title", 3800
End Function

Public Function RemoveBook()
    If lvPrint.ListItems.Count > 0 Then
        lvPrint.ListItems.Remove lvPrint.SelectedItem.Index
    End If
End Function
Public Function InsertBook()
    Dim mRow As ListItem
    Dim vCountItem As Variant
    If isItemExist = False Then
        vCountItem = Split(lvList(0).SelectedItem.Text, "(")
        If Val(vCountItem(1)) > 0 Then
            Set mRow = lvPrint.ListItems.Add(, , , , 2)
            mRow.SubItems(1) = lvList(0).SelectedItem.SubItems(1)
            mRow.SubItems(2) = lvList(0).SelectedItem.SubItems(2)
            mRow.SubItems(3) = lvList(0).SelectedItem.SubItems(3)
        Else
            MsgBox "Unabled to add to Print Items. Quantity in Shelf is 0.", vbExclamation, "UnabledItem"
        End If
    Else
        MsgBox "Record Already Exist!", vbExclamation, "ItemExist"
    End If
End Function

Public Function isItemExist() As Boolean
    Dim i As Integer
    isItemExist = False
    For i = 1 To lvPrint.ListItems.Count
        If lvPrint.ListItems(i).SubItems(1) = lvList(0).SelectedItem.SubItems(1) Then
            isItemExist = True
            Exit For
        End If
    Next
End Function

Public Function GeneratePrint()
    Dim i As Integer
    Dim sSQL As String
    Dim sReceipt As String
    On Error GoTo errHandler
    sReceipt = ""
    sReceipt = sReceipt & "TRINITY UNIVERSITY" & vbCrLf
    sReceipt = sReceipt & "QUEZON CITY" & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Print Date: " & Date & vbCrLf & vbCrLf
    sReceipt = sReceipt & "Call No." & vbTab & vbTab & "ISBN" & vbCrLf
    sReceipt = sReceipt & "------------------------------------------" & vbCrLf
    For i = 1 To lvPrint.ListItems.Count
        sReceipt = sReceipt & lvPrint.ListItems(i).SubItems(1) & vbTab & vbTab & lvPrint.ListItems(i).SubItems(2) & vbCrLf
    Next
    sReceipt = sReceipt & "------------------------------------------" & vbCrLf
    sReceipt = sReceipt & "Library Official Receipt" & vbCrLf
    sReceipt = sReceipt & "of Transaction" & vbCrLf
    'MsgBox sReceipt
    Printer.Print sReceipt
errHandler:
    If err.Number = 487 Then
        MsgBox "Error Printing. No Printer Detected.", vbExclamation, "PrinterException"
    End If
End Function

Public Function CheckRecordsQty()
    Dim i As Integer
    Dim vCallNo As Variant, sSQL As String
    Dim lCountItem As Long
    For i = 1 To lvList(0).ListItems.Count
        vCallNo = Split(lvList(0).ListItems(i).SubItems(1), "-")
        sSQL = "SELECT tbl_reg_books.rb_id " & _
            "FROM tbl_shelfs INNER JOIN (tbl_reg_books INNER JOIN tbl_shelfbooks ON tbl_reg_books.rb_id = tbl_shelfbooks.rb_id) ON tbl_shelfs.sh_id = tbl_shelfbooks.sh_id " & _
            "WHERE (((tbl_shelfbooks.callno) Like '" & vCallNo(1) & "') AND ((tbl_shelfs.acronym) Like '" & vCallNo(0) & "') AND ((tbl_reg_books.borrow) Like '0' And (tbl_reg_books.borrow) Like '0')) " & _
            "GROUP BY tbl_reg_books.rb_id;"
        lCountItem = isCountItem(sSQL)
        lvList(0).ListItems(i).Text = "(" & lCountItem & ")"
        If lCountItem > 0 Then
            lvList(0).ListItems(i).ForeColor = vbBlack
            lvList(0).ListItems(i).SmallIcon = 2
        Else
            lvList(0).ListItems(i).SmallIcon = 1
            lvList(0).ListItems(i).ForeColor = vbRed
            lvList(0).ListItems(i).ListSubItems(1).ForeColor = vbRed
            lvList(0).ListItems(i).ListSubItems(2).ForeColor = vbRed
            lvList(0).ListItems(i).ListSubItems(3).ForeColor = vbRed
            lvList(0).ListItems(i).ListSubItems(4).ForeColor = vbRed
        End If
    Next
End Function

Public Function WhereCreator(iSearch As Integer, sItem As String, Index As Integer) As String
    Dim vItems As Variant, sField As String
    Dim i As Integer, iCountItem As Long, j As Integer
    Dim lJplus1 As Long
    WhereCreator = ""
    If Index = 0 Then
        If iSearch = 1 Then
            sField = "title"
        ElseIf iSearch = 2 Then
            sField = "author"
        End If
    ElseIf Index = 1 Then
        If iSearch = 1 Then
            sField = "article"
        ElseIf iSearch = 2 Then
            sField = "authors"
        End If
    End If
    vItems = Split(sItem, " ")
    iCountItem = CountSplitItem(sItem, " ")
    For i = 0 To iCountItem
        If isInvalid(vItems(i)) = False Then
            WhereCreator = WhereCreator & sField & " Like '%" & vItems(i) & "%' "
            If iCountItem > 0 And Not i = iCountItem Then
                For j = (i + 1) To iCountItem
                    If isInvalid(vItems(j)) = False And Not Val(j) > iCountItem Then
                        If Len(Trim(vItems(j))) > 0 Then
                            'MsgBox vItems(j)
                            Debug.Print "-" & vItems(j) & "-"
                            WhereCreator = WhereCreator & " OR "
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    Next
    If Len(Trim(WhereCreator)) > 0 Then
        WhereCreator = "(" & WhereCreator & ")"
    End If
End Function

Public Function isInvalid(sInvalid As Variant) As Boolean
    Dim sSQL As String
    sSQL = "SELECT tbl_invalid_words.InvalidWord " & _
        "From tbl_invalid_words " & _
        "WHERE (((tbl_invalid_words.InvalidWord) Like '" & sInvalid & "')) " & _
        "GROUP BY tbl_invalid_words.InvalidWord;"
    'MsgBox sInvalid
    If isRecordExist(sSQL) = False Then
        isInvalid = False
    Else
        isInvalid = True
    End If
End Function
