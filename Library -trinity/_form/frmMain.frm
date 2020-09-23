VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Trinity University Of Asia Library System v.1.0.0"
   ClientHeight    =   10530
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14295
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrPopup 
      Interval        =   3000
      Left            =   4440
      Top             =   3120
   End
   Begin VB.PictureBox picHead 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14295
      TabIndex        =   7
      Top             =   0
      Width           =   14295
      Begin MSComctlLib.Toolbar tbPenalty 
         Height          =   330
         Left            =   2415
         TabIndex        =   12
         Top             =   0
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar tbBR 
         Height          =   330
         Left            =   855
         TabIndex        =   11
         Top             =   0
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList iHead2 
         Left            =   5640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1992
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3324
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4CB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6648
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7FDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":996C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B2FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11488
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList iHead 
         Left            =   5040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12E1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":136F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":13FCE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":29140
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F2CA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbHead 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picHeadLn 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   14235
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   14295
      Begin MSComctlLib.TabStrip tabMain 
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   873
         TabFixedWidth   =   3528
         MultiSelect     =   -1  'True
         Separators      =   -1  'True
         TabMinWidth     =   3528
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
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
   End
   Begin VB.PictureBox picfrmBtn 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   9540
      Left            =   2805
      MousePointer    =   9  'Size W E
      ScaleHeight     =   9540
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   735
      Width           =   75
      Begin VB.PictureBox PicLeftBtn 
         Height          =   1095
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   1035
         ScaleWidth      =   75
         TabIndex        =   5
         Top             =   4560
         Width           =   135
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   9540
      Left            =   0
      ScaleHeight     =   9540
      ScaleWidth      =   2805
      TabIndex        =   1
      Top             =   735
      Width           =   2800
      Begin MSComctlLib.ImageList iShortcut 
         Left            =   240
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   25
         ImageHeight     =   25
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":45C8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5C64E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":73010
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":88182
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9F21C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A5A7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":ABC08
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B1D92
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C8754
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CEFB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E5978
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":EBB02
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1024C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10864E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10E7D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12519A
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":13BB5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15251E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList iOpen 
         Left            =   120
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   21
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":168EE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16F06A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1758CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":17BA56
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":190BC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A758A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1AD714
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B389E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1CA260
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E0C22
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F75E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1FD76E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2038F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21A2BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":230C7C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24763E
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25E000
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2749C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28B384
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A1D46
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B8708
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraShortcut 
         Height          =   465
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   2610
         Begin VB.Image Image2 
            Height          =   240
            Left            =   120
            Picture         =   "frmMain.frx":2BE892
            Stretch         =   -1  'True
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Shortcuts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00BB5900&
            Height          =   240
            Left            =   480
            TabIndex        =   3
            Top             =   150
            Width           =   930
         End
      End
      Begin MSComctlLib.ListView lvShortcut 
         Height          =   5415
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         Height          =   885
         Left            =   -120
         Picture         =   "frmMain.frx":2BF15C
         Top             =   0
         Width           =   2970
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iPageEnabled 
      Left            =   3240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C7AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C7E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C8230
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C85CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C8964
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DD6D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F4098
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FA222
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3003AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iPageDisabled 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":306536
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3068D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":306C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":307004
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30739E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31C110
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":332AD2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iLv 
      Left            =   3240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":338C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34F61E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":365FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37C9A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":393364
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A84D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BEE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D585A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EC8F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F3156
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F92E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FF46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":415E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41C68E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":433050
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4391DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43F364
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4454EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45BEB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   3840
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":472872
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":473284
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":473C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4746A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4750BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":475ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47C32E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47E468
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":495502
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49B68C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A1816
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A79A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BCB12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D3BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E8D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EEEA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F0D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F6EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FD716
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5038A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iListView 
      Left            =   4440
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":509A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5203EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":536DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54D770
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5628E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5792A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58FC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":595DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59BF7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A27DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B9876
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C00D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i25x25 
      Left            =   4440
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C6262
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DCC24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F35E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFES 
         Caption         =   "Edit Selected"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Begin VB.Menu mnuFtxt 
            Caption         =   "Find Text"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuSQ 
            Caption         =   "Search Query"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu lnFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAClose 
         Caption         =   "Close All Form"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log-off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRec 
      Caption         =   "&Records"
      Begin VB.Menu mnuTrans 
         Caption         =   "&Transaction"
         Begin VB.Menu mnuBRT 
            Caption         =   "Borrow and Return"
         End
         Begin VB.Menu mnuPP 
            Caption         =   "Penalty Payment"
         End
      End
      Begin VB.Menu mnuLR 
         Caption         =   "&Library Records"
         Begin VB.Menu mnuSect 
            Caption         =   "Sections"
         End
         Begin VB.Menu mnuShlf 
            Caption         =   "Shelfs"
         End
         Begin VB.Menu mnuBI 
            Caption         =   "Books"
         End
         Begin VB.Menu mnuRB 
            Caption         =   "Reg. Books"
         End
         Begin VB.Menu muPeriod 
            Caption         =   "Periodicals"
         End
         Begin VB.Menu lneLR1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPub 
            Caption         =   "Publishers"
         End
         Begin VB.Menu mnuCount 
            Caption         =   "Countries"
         End
         Begin VB.Menu mnuBorrower 
            Caption         =   "Borrowers"
         End
         Begin VB.Menu mnuBT 
            Caption         =   "Borrower Types"
         End
         Begin VB.Menu mnuAut 
            Caption         =   "Authors"
         End
         Begin VB.Menu lneLR2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHoli 
            Caption         =   "Holidays"
         End
         Begin VB.Menu mnuIW 
            Caption         =   "Invalid Words"
         End
         Begin VB.Menu LR3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOB 
            Caption         =   "Overdue Books"
         End
         Begin VB.Menu mnuPending 
            Caption         =   "Pending Books"
         End
         Begin VB.Menu mnuBB 
            Caption         =   "Borrowed Books"
         End
         Begin VB.Menu mnuBL 
            Caption         =   "Borrowers Logs"
         End
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "R&eports"
      Begin VB.Menu mnuBookR 
         Caption         =   "Books"
         Begin VB.Menu mnuBBR 
            Caption         =   "Books Returned Report"
         End
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utility"
      Begin VB.Menu mnuSU 
         Caption         =   "System User"
      End
      Begin VB.Menu mnuSec 
         Caption         =   "Security"
         Begin VB.Menu mnuChange 
            Caption         =   "Set Password"
         End
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Option"
         Begin VB.Menu mnuDT 
            Caption         =   "Set Date and Time"
         End
         Begin VB.Menu mnuPopup 
            Caption         =   "Pop-Up Ovedue Books"
         End
      End
      Begin VB.Menu lneUtil1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTH 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuTV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuAI 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About TUA Library System v.1.0.0"
      End
   End
   Begin VB.Menu mnuAct 
      Caption         =   "Action"
      Begin VB.Menu mnuCN 
         Caption         =   "Create New"
      End
      Begin VB.Menu mnuES 
         Caption         =   "Edited Selected"
      End
      Begin VB.Menu mnuS 
         Caption         =   "Search"
         Begin VB.Menu mnuFindTxt 
            Caption         =   "Find Text"
         End
         Begin VB.Menu mnuSearchQ 
            Caption         =   "Search Query"
         End
      End
      Begin VB.Menu mnuDS 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuR 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuP 
         Caption         =   "Print"
      End
      Begin VB.Menu lneAct1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuC 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuFS 
      Caption         =   "Search"
      Begin VB.Menu mnuF 
         Caption         =   "Find Text"
      End
      Begin VB.Menu mnuQ 
         Caption         =   "Search Query"
      End
   End
   Begin VB.Menu mnuTrans2 
      Caption         =   "Transaction"
      Begin VB.Menu mnuBorrow 
         Caption         =   "Borrow"
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "Return"
      End
   End
   Begin VB.Menu mnuSrch 
      Caption         =   "Search"
      Begin VB.Menu mnuII 
         Caption         =   "Insert Item"
      End
      Begin VB.Menu mnuRI 
         Caption         =   "Remove Item"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INT_SIZE As Integer
Dim IntSizeActive As Integer
Public iLvIndex As Integer

Private Sub lvShortcut_Click()
    lvShortcut.ToolTipText = lvShortcut.SelectedItem.Text
End Sub

Private Sub MDIForm_Load()
    iUSER = 1
    Top = 0
    Left = 0
    INT_SIZE = Me.picLeft.Height / 2
    Call PIC_RESIZE_LEFT
    MDIForm_Resize
    PIC_LEFT_LOST_FOCUS picLeft
    Load_tbHeader tbHead, i16x16, 16
    Activate_tbHeader tbHead, ""
    Show_tbHeader True, frmMain
    ToolMenuStatus False
    frmLogin.Show
    frmLogin.WindowState = vbMaximized
    Call FOR_LEFT_PICTURE
    mnuAct.Visible = False
    mnuFS.Visible = False
    tabMain.Tabs.Clear
End Sub

Private Sub lvShortcut_DblClick()
    Select Case lvShortcut.SelectedItem.Key
        Case "L1": TabLoadForm frmBook
        Case "L2": TabLoadForm frmMagazine
        Case "L3": TabLoadForm frmPublisher
        Case "L7": TabLoadForm frmCountries
        Case "L4": TabLoadForm frmSection
        Case "L6": TabLoadForm frmUser
        Case "L5": TabLoadForm frmShelf
        Case "L8": TabLoadForm frmBorrower
        Case "L9": TabLoadForm frmBType
        Case "L10": TabLoadForm frmInvalidWord
        Case "L11": TabLoadForm frmHoliday
        Case "L12": TabLoadForm frmAuthor
        Case "L13": TabLoadForm frmBorrowerLog
        Case "L14": TabLoadForm frmBookBorrowed
        Case "L15": TabLoadForm frmBookPending
        Case "L16": TabLoadForm frmBookDueDate
        Case "L17": TabLoadForm frmRegBook
        End Select
End Sub

Public Function FOR_LEFT_PICTURE()
    Dim lv_listitems As ListItem
    Dim cRow As ColumnHeader
    Set lvShortcut.SmallIcons = iShortcut
    Set lvShortcut.Icons = iShortcut
    lvShortcut.ColumnHeaders.Add , , ""
    With lvShortcut
        .ListItems.Clear
        If iUSER = 1 Then 'ADMIN
            .ListItems.Add , "L1", "Books", 1, 1
            .ListItems.Add , "L2", "Periodicals", 2, 2
            .ListItems.Add , "L3", "Publishers", 3, 3
            .ListItems.Add , "L4", "Sections", 4, 4
            .ListItems.Add , "L5", "Shelfs", 5, 5
            .ListItems.Add , "L6", "Users", 6, 6
            .ListItems.Add , "L8", "Borrowers", 8, 8
            .ListItems.Add , "L7", "Countries", 7, 7
            .ListItems.Add , "L9", "B.Type", 9, 9
            .ListItems.Add , "L10", "Invalid Words", 10, 10
            .ListItems.Add , "L11", "Holidays", 11, 11
            .ListItems.Add , "L12", "Authors", 12, 12
            .ListItems.Add , "L13", "Borrowers Log", 13, 13
            .ListItems.Add , "L14", "Borrowed Books", 14, 14
            .ListItems.Add , "L15", "Pending Books", 15, 15
            .ListItems.Add , "L16", "Overdue Books", 16, 16
            .ListItems.Add , "L17", "Reg. Books", 17, 17
            '.ListItems.Add , "L18", "Reg. Periodicals", 18, 18
        ElseIf iUSER = 2 Then 'BORROWING
        
        ElseIf iUSER = 3 Then 'OPAC

        End If
    End With
End Function


Private Sub MDIForm_Resize()
    On Error Resume Next
    If picHead.ScaleWidth > 0 Then
        tabMain.Width = picHeadLn.ScaleWidth
    End If
    picHead.Height = tbHead.Height + 40
End Sub

Sub PIC_RESIZE_LEFT()
    On Error Resume Next
    fraShortcut.Move 80, 900, picLeft.Width - 80
    lvShortcut.Move 80, 1440, picLeft.Width - 80, picLeft.Height - (1440 - 80)
    PicLeftBtn.Move 0, (picfrmBtn.Height / 2) - PicLeftBtn.Height
    lvShortcut.ColumnHeaders(1).Width = lvShortcut.Width - 80
    tabMain.Move Me.picLeft.Width, 0, Me.picHeadLn.ScaleWidth - picLeft.Width
End Sub

Private Sub mnuAClose_Click()
    On Error Resume Next
    Do While Not err.Number = 91
        activeForm.Visible = False
    Loop
    tabMain.Tabs.Clear
    picHeadLn.Visible = False
End Sub

Private Sub mnuAI_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuAut_Click()
    SearchLVItem 12
    TabLoadForm frmAuthor
End Sub

Private Sub mnuBB_Click()
    SearchLVItem 14
    TabLoadForm frmBookBorrowed
End Sub

Private Sub mnuBBR_Click()
    frmSales_Rpt.Show 1
End Sub

Private Sub mnuBI_Click()
    SearchLVItem 1
    TabLoadForm frmBook
End Sub

Private Sub mnuBL_Click()
    SearchLVItem 13
    TabLoadForm frmBorrowerLog
End Sub

Private Sub mnuBorrower_Click()
    SearchLVItem 7
    TabLoadForm frmBorrower
End Sub

Private Sub mnuBRT_Click()
    LoadForm frmBorrow
End Sub

Private Sub mnuBT_Click()
    SearchLVItem 9
    TabLoadForm frmBType
End Sub

Private Sub mnuC_Click()
    activeForm.LvClose activeForm.iLvIndex
End Sub

Private Sub mnuCalc_Click()
    On Error Resume Next
    Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuChange_Click()
    frmSetPass.Show 1
End Sub

Private Sub mnuClose_Click()
    On Error Resume Next
    Unload activeForm
End Sub

Private Sub mnuCN_Click()
    On Error Resume Next
    activeForm.LvNew activeForm.iLvIndex
End Sub

Private Sub mnuCount_Click()
    SearchLVItem 8
    TabLoadForm frmCountries
End Sub

Private Sub mnuDel_Click()
    On Error Resume Next
    activeForm.LvDelete activeForm.iLvIndex
End Sub

Private Sub mnuDS_Click()
    On Error Resume Next
    activeForm.LvDelete activeForm.iLvIndex
End Sub

Private Sub mnuDT_Click()
    frmDateChecker.Show 1
End Sub

Private Sub mnuES_Click()
    On Error Resume Next
    activeForm.LvEdit activeForm.iLvIndex
End Sub

Private Sub mnuF_Click()
    activeForm.FindText
End Sub

Private Sub mnuFES_Click()
    On Error Resume Next
    activeForm.LvEdit activeForm.iLvIndex
End Sub

Private Sub mnuFindTxt_Click()
    activeForm.FindText
End Sub

Private Sub mnuFtxt_Click()
    On Error Resume Next
    activeForm.FindText
End Sub

Private Sub mnuHoli_Click()
    SearchLVItem 11
    TabLoadForm frmHoliday
End Sub

Private Sub mnuII_Click()
    frmOPAC.InsertBook
End Sub

Private Sub mnuIW_Click()
    SearchLVItem 10
    TabLoadForm frmInvalidWord
End Sub

Private Sub mnuLogOff_Click()
    ToolMenuStatus False
    frmLogin.Show
    frmLogin.WindowState = vbMaximized
End Sub

Private Sub mnuNew_Click()
    On Error Resume Next
    activeForm.LvNew activeForm.iLvIndex
End Sub

Private Sub mnuNotepad_Click()
    On Error Resume Next
    Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub mnuOB_Click()
    SearchLVItem 16
    TabLoadForm frmBookDueDate
End Sub

Private Sub mnuP_Click()
    On Error Resume Next
    activeForm.PRINT_RECORD iLvIndex
End Sub

Private Sub mnuPending_Click()
    SearchLVItem 15
Private Sub mnuPopup_Click()
    frmOption.Show 1
End Sub

Private Sub mnuPP_Click()
    LoadForm frmPenaltyPayment
End Sub

Private Sub mnuPub_Click()
    SearchLVItem 3
    TabLoadForm frmPublisher
End Sub

Private Sub mnuQ_Click()
    On Error Resume Next
    activeForm.SearchItem
End Sub

Private Sub mnuR_Click()
    On Error Resume Next
    activeForm.LvRefresh iLvIndex
End Sub

Private Sub mnuRB_Click()
    SearchLVItem 17
    TabLoadForm frmRegBook
End Sub

Private Sub mnuRefresh_Click()
    On Error Resume Next
    activeForm.LvRefresh iLvIndex
End Sub

Private Sub mnuRI_Click()
    frmOPAC.RemoveBook
End Sub

Private Sub mnuSearchQ_Click()
    On Error Resume Next
    activeForm.SearchItem
End Sub

Private Sub mnuSect_Click()
    SearchLVItem 4
    TabLoadForm frmSection
End Sub

Private Sub mnuShlf_Click()
    SearchLVItem 5
    TabLoadForm frmShelf
End Sub

Private Sub mnuSQ_Click()
    On Error Resume Next
    activeForm.SearchItem
End Sub

Private Sub mnuSU_Click()
    SearchLVItem 6
    TabLoadForm frmUser
End Sub

Private Sub mnuTH_Click()
    On Error Resume Next
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTV_Click()
    On Error Resume Next
    Me.Arrange vbTileVertical
End Sub

Private Sub muPeriod_Click()
    SearchLVItem 2
    TabLoadForm frmMagazine
End Sub

Private Sub picfrmBtn_Click()
    If picLeft.Width > 15 Then
        picLeft.Width = 0
    Else
        picLeft.Width = 2800
    End If
    PIC_RESIZE_LEFT
End Sub

Private Sub PicLeftBtn_Click()
    If picLeft.Width > 15 Then
        picLeft.Width = 0
    Else
        picLeft.Width = 2800
    End If
    PIC_RESIZE_LEFT
End Sub

Private Sub tabMain_Click()
    Dim iTabLoop As Integer
    Dim sSelected As String
    sSelected = tabMain.SelectedItem.Key
    Select Case sSelected
        Case "Books": LoadForm frmBook
        Case "Periodicals": LoadForm frmMagazine
        Case "Publishers": LoadForm frmPublisher
        Case "Countries": LoadForm frmCountries
        Case "Sections": LoadForm frmSection
        Case "Users": LoadForm frmUser
        Case "Shelfs": LoadForm frmShelf
        Case "Borrowers": LoadForm frmBorrower
        Case "B.Type": LoadForm frmBType
        Case "Invalid Words": LoadForm frmInvalidWord
        Case "Holidays": LoadForm frmHoliday
        Case "Borrowers Log": LoadForm frmBorrowerLog
        Case "Authors": LoadForm frmAuthor
        Case "Borrowed Books": LoadForm frmBookBorrowed
        Case "Pending Books": LoadForm frmBookPending
        Case "Overdue Books": LoadForm frmBookDueDate
        Case "Reg. Books": LoadForm frmRegBook
    End Select
End Sub

Private Sub tbBR_ButtonClick(ByVal Button As MSComctlLib.Button)
    LoadForm frmBorrow
End Sub

Private Sub tbHead_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            If picLeft.Width > 15 Then
                picLeft.Width = 0
            Else
                picLeft.Width = 2800
            End If
            PIC_RESIZE_LEFT
    End Select
End Sub

Public Function ToolMenuStatus(blStat As Boolean)
    If blStat = True Then
        picfrmBtn.Visible = True
        picLeft.Visible = True
        picHead.Visible = True
        Me.mnuFile.Visible = True
        Me.mnuRec.Visible = True
        Me.mnuRep.Visible = True
        Me.mnuUtil.Visible = True
        Me.mnuHlp.Visible = True
        Me.mnuWindow.Visible = True
        Me.StatusBar1.Visible = True
    Else
        picfrmBtn.Visible = False
        picLeft.Visible = False
        picHeadLn.Visible = False
        picHead.Visible = False
        Me.mnuFile.Visible = False
        Me.mnuRec.Visible = False
        Me.mnuRep.Visible = False
        Me.mnuUtil.Visible = False
        Me.mnuAct.Visible = False
        Me.mnuHlp.Visible = False
        Me.StatusBar1.Visible = False
        Me.mnuWindow.Visible = False
        Me.mnuTrans2.Visible = False
        Me.mnuSrch.Visible = False
    End If
End Function

Public Function TabMainIni(iExe As Integer, sTab As String, iIcon As Integer) As Boolean
    Dim iTabLoop As Integer
    Set tabMain.ImageList = iLv
    If iExe = 1 Then 'remove tab
        For iTabLoop = 1 To tabMain.Tabs.Count
            If tabMain.Tabs(iTabLoop).Caption = sTab Then
                Select Case tabMain.Tabs(iTabLoop).Caption
                    Case "Books": tabMain.Tabs.Remove iTabLoop
                    Case "Periodicals": tabMain.Tabs.Remove iTabLoop
                    Case "Publishers": tabMain.Tabs.Remove iTabLoop
                    Case "Sections": tabMain.Tabs.Remove iTabLoop
                    Case "Users": tabMain.Tabs.Remove iTabLoop
                    Case "Countries": tabMain.Tabs.Remove iTabLoop
                    Case "Shelfs": tabMain.Tabs.Remove iTabLoop
                    Case "Borrowers": tabMain.Tabs.Remove iTabLoop
                    Case "B.Type": tabMain.Tabs.Remove iTabLoop
                    Case "Invalid Words": tabMain.Tabs.Remove iTabLoop
                    Case "Holidays": tabMain.Tabs.Remove iTabLoop
                    Case "Authors": tabMain.Tabs.Remove iTabLoop
                    Case "Borrowers Log": tabMain.Tabs.Remove iTabLoop
                    Case "Borrowed Books": tabMain.Tabs.Remove iTabLoop
                    Case "Pending Books": tabMain.Tabs.Remove iTabLoop
                    Case "Overdue Books":  tabMain.Tabs.Remove iTabLoop
                    Case "Reg. Books":  tabMain.Tabs.Remove iTabLoop
                End Select
                If tabMain.Tabs.Count = 0 Then
                    picHeadLn.Visible = False
                End If
                Exit For
            End If
        Next
    ElseIf iExe = 2 Then 'insert
        TabMainIni = False
        For iTabLoop = 1 To tabMain.Tabs.Count
            If tabMain.Tabs(iTabLoop).Caption = sTab Then
                tabMain.Tabs(iTabLoop).Selected = True
                tabMain.SelectedItem.HighLighted = True
                TabMainIni = True
                Exit For
            End If
        Next
        If TabMainIni = False Then
            picHeadLn.Visible = True
            tabMain.Tabs.Add , sTab, sTab, iIcon
            tabMain.Tabs(tabMain.Tabs.Count).Selected = True
        Else
            For iTabLoop = 1 To tabMain.Tabs.Count
                Select Case tabMain.Tabs(iTabLoop).Key
                    Case "Books": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Periodicals": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Publishers": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Countries": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Sections": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Users": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Shelfs": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Borrowers": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "B.Type": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Invalid Words": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Holidays": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Authors": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Borrowers Log": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Borrowed Books": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Pending Books": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Overdue Books": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                    Case "Reg. Books": tabMain.Tabs(iTabLoop).Selected = True: Exit For
                End Select
            Next
        End If
    End If
End Function

Public Sub TabLoadForm(ByRef srcForm As Form)
    Dim iTabLoop As Integer
    Dim sSelected As String
    srcForm.Show
    srcForm.SetFocus
    sSelected = lvShortcut.SelectedItem.Text
    Select Case sSelected
        Case "Books": LoadForm frmBook
        Case "Periodicals": LoadForm frmMagazine
        Case "Publishers": LoadForm frmPublisher
        Case "Countries": LoadForm frmCountries
        Case "Sections": LoadForm frmSection
        Case "Users": LoadForm frmUser
        Case "Shelfs": LoadForm frmShelf
        Case "Borrowers": LoadForm frmBorrower
        Case "B.Type": LoadForm frmBType
        Case "Invalid Words": LoadForm frmInvalidWord
        Case "Holidays": LoadForm frmHoliday
        Case "Authors": LoadForm frmAuthor
        Case "Borrowers Log": LoadForm frmBorrowerLog
        Case "Borrowed Books": LoadForm frmBookBorrowed
        Case "Pending Books": LoadForm frmBookPending
        Case "Overdue Books": LoadForm frmBookDueDate
        Case "Reg. Books": LoadForm frmRegBook
    End Select
    For iTabLoop = 1 To tabMain.Tabs.Count
        If tabMain.Tabs(iTabLoop).Caption = lvShortcut.SelectedItem.Text Then
            tabMain.Tabs(iTabLoop).Selected = True
            Exit For
        End If
    Next
End Sub

Private Sub tbPenalty_ButtonClick(ByVal Button As MSComctlLib.Button)
    LoadForm frmPenaltyPayment
End Sub

Private Sub tmrPopup_Timer()
    Dim sSQL As String, iCountItem As Integer
    On Error Resume Next
    sSQL = "SELECT tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date " & _
        "FROM (tbl_books INNER JOIN tbl_reg_books ON tbl_books.isbn = tbl_reg_books.isbn) INNER JOIN (tbl_borrowers INNER JOIN tbl_borrow_record ON tbl_borrowers.B_id = tbl_borrow_record.B_id) ON tbl_reg_books.rb_id = tbl_borrow_record.rb_id " & _
        "WHERE (tbl_borrow_record.s_return Like '0') AND (tbl_borrow_record.r_date<date()) " & _
        "GROUP BY tbl_borrow_record.br_id, tbl_borrow_record.B_id, tbl_borrowers.fn, tbl_borrowers.mn, tbl_borrowers.ln, tbl_reg_books.isbn, tbl_books.title, tbl_borrow_record.b_date, tbl_borrow_record.r_date;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, adOpenStatic, adLockOptimistic
        iCountItem = adoRes.RecordCount
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
    sSQL = "SELECT tbl_popup.popuptime, tbl_popup.popupstat " & _
        "FROM tbl_popup;"
    Set adoCon = New ADODB.Connection
    Set adoRes = New ADODB.Recordset
    adoCon.Open sCon
    adoRes.Open sSQL, adoCon, 3, 3
        tmrPopup.Interval = adoRes.Fields("popuptime")
        If Val(adoRes.Fields("popupstat")) = 1 Then
            If iCountItem > 0 Then
                frmPopUp.Show 1
            End If
        End If
    adoRes.Close
    adoCon.Close
    Set adoCon = Nothing
    Set adoRes = Nothing
End Sub

Public Function SearchLVItem(iLVitem As Long)
    lvShortcut.ListItems(iLVitem).Selected = True
End Function
