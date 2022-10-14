VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmShow 
   BackColor       =   &H00000000&
   Caption         =   "GPJ Show Plans & Show Requlation Abstracts"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13620
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDir 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   12000
      TabIndex        =   4
      Top             =   600
      Width           =   12000
      Begin TabDlg.SSTab sst1 
         Height          =   7575
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   13361
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   2
         ShowFocusRect   =   0   'False
         BackColor       =   0
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmShow.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblCount"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblShow"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblMess"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "rtbMess"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "tvw1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdEngCodes"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cboSHYR(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cboCUNO(1)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "optSort(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "optSort(1)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "fraEngCodes"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "fraButtons"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         Begin VB.Frame fraButtons 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   180
            TabIndex        =   34
            Top             =   7020
            Width           =   6915
            Begin VB.CommandButton cmdViewShowplan 
               Caption         =   "Showplan"
               Enabled         =   0   'False
               Height          =   435
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Click to View Overall Showplan"
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.CommandButton cmdSRA 
               Caption         =   "Show Regulation Abstract"
               Height          =   435
               Left            =   4740
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.CommandButton cmdViewComposite 
               Caption         =   "Composite"
               Enabled         =   0   'False
               Height          =   435
               Left            =   3660
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Click to View Composite Showplan"
               Top             =   0
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Image Image1 
               Height          =   315
               Index           =   3
               Left            =   1020
               Picture         =   "frmShow.frx":08E6
               Stretch         =   -1  'True
               ToolTipText     =   "This icon signifies that Engineering Code Requirements are available"
               Top             =   90
               Width           =   315
            End
            Begin VB.Image Image1 
               Height          =   315
               Index           =   2
               Left            =   660
               Picture         =   "frmShow.frx":0E70
               Stretch         =   -1  'True
               ToolTipText     =   "View list of Attending Clients in this Hall by clicking nodes with this Icon"
               Top             =   90
               Width           =   315
            End
            Begin VB.Image Image1 
               Height          =   255
               Index           =   1
               Left            =   0
               Picture         =   "frmShow.frx":173A
               Stretch         =   -1  'True
               ToolTipText     =   "This icon signifies that an Overall Showplan or a Composite Plan is available"
               Top             =   120
               Width           =   255
            End
            Begin VB.Image Image1 
               Height          =   315
               Index           =   0
               Left            =   300
               Picture         =   "frmShow.frx":1CC4
               Stretch         =   -1  'True
               ToolTipText     =   "Additional Info is available by clicking nodes with this Icon"
               Top             =   90
               Width           =   315
            End
         End
         Begin VB.Frame fraEngCodes 
            Height          =   6255
            Left            =   7260
            TabIndex        =   9
            Top             =   600
            Visible         =   0   'False
            Width           =   4515
            Begin VB.ComboBox cboCodeName 
               Height          =   315
               ItemData        =   "frmShow.frx":224E
               Left            =   120
               List            =   "frmShow.frx":2250
               TabIndex        =   15
               Top             =   570
               Width           =   4275
            End
            Begin VB.CommandButton cmdCodeDelete 
               Caption         =   "Delete Selection"
               Enabled         =   0   'False
               Height          =   375
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   5340
               Width           =   1635
            End
            Begin VB.CommandButton cmdCodeUpdate 
               Caption         =   "Update Entry"
               Height          =   375
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   5340
               Width           =   1635
            End
            Begin VB.CommandButton cmdCodeClear 
               Caption         =   "Clear Selection"
               Height          =   375
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   5760
               Width           =   1635
            End
            Begin VB.CommandButton cmdCodeSave 
               Caption         =   "Save New Entry"
               Height          =   375
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   5760
               Width           =   1635
            End
            Begin VB.CommandButton cmdApprove 
               Caption         =   "Review && Approve..."
               Height          =   375
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   1935
            End
            Begin RichTextLib.RichTextBox rtbCodeDesc 
               Height          =   3315
               Left            =   120
               TabIndex        =   16
               Top             =   1230
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   5847
               _Version        =   393217
               BorderStyle     =   0
               Enabled         =   -1  'True
               ScrollBars      =   2
               MaxLength       =   1000
               TextRTF         =   $"frmShow.frx":2252
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code Name:"
               Height          =   195
               Left            =   180
               TabIndex        =   21
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code Description:"
               Height          =   195
               Left            =   180
               TabIndex        =   20
               Top             =   1020
               Width           =   1275
            End
            Begin VB.Label lblEditUser 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   4560
               Width           =   45
            End
            Begin VB.Label lblCodeID 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4320
               TabIndex        =   18
               Top             =   960
               Width           =   90
            End
            Begin VB.Label lblConfUser 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   4860
               Width           =   4005
               WordWrap        =   -1  'True
            End
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmShow.frx":22CD
            Height          =   495
            Index           =   1
            Left            =   6480
            Picture         =   "frmShow.frx":25D7
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Sort Show List Chronologically (by Show Open Date)"
            Top             =   180
            Width           =   555
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmShow.frx":2EA1
            Height          =   495
            Index           =   0
            Left            =   5940
            Picture         =   "frmShow.frx":31AB
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Sort Show List Alphabetically"
            Top             =   180
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.ComboBox cboCUNO 
            Height          =   315
            Index           =   1
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   300
            Width           =   4695
         End
         Begin VB.ComboBox cboSHYR 
            Height          =   315
            Index           =   1
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdEngCodes 
            Caption         =   "Edit Engineering Code Regulations"
            Height          =   435
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6960
            Visible         =   0   'False
            Width           =   3675
         End
         Begin MSComctlLib.TreeView tvw1 
            Height          =   6135
            Left            =   180
            TabIndex        =   27
            Top             =   720
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   10821
            _Version        =   393217
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            MousePointer    =   99
            MouseIcon       =   "frmShow.frx":3A75
         End
         Begin RichTextLib.RichTextBox rtbMess 
            Height          =   5655
            Left            =   7320
            TabIndex        =   28
            Top             =   960
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   9975
            _Version        =   393217
            BackColor       =   -2147483633
            BorderStyle     =   0
            Enabled         =   0   'False
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmShow.frx":3D8F
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
         Begin VB.Label lblMess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7260
            TabIndex        =   33
            Top             =   165
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7260
            TabIndex        =   32
            Top             =   585
            UseMnemonic     =   0   'False
            Width           =   45
         End
         Begin VB.Label lblCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   210
            TabIndex        =   31
            Top             =   6900
            Width           =   45
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Show Year:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   30
            Top             =   90
            Width           =   870
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Client:"
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   29
            Top             =   90
            Width           =   510
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12240
      Top             =   1500
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
            Picture         =   "frmShow.frx":3E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":4124
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":46BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":4C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":5532
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":568C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShow.frx":5C26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   1080
      ScaleHeight     =   8025
      ScaleWidth      =   12465
      TabIndex        =   1
      Top             =   1140
      Width           =   12495
   End
   Begin VOLOVIEWXLibCtl.AvViewX volShowplan 
      Height          =   615
      Left            =   10980
      TabIndex        =   0
      Top             =   660
      Width           =   795
      _cx             =   4195706
      _cy             =   5080
      Appearance      =   0
      BorderStyle     =   0
      BackgroundColor =   "DefaultColors"
      Enabled         =   -1  'True
      UserMode        =   "Pan"
      HighlightLinks  =   0   'False
      src             =   ""
      LayersOn        =   ""
      LayersOff       =   ""
      SrcTemp         =   ""
      SupportPath     =   $"frmShow.frx":61C0
      FontPath        =   $"frmShow.frx":6336
      NamedView       =   ""
      GeometryColor   =   "DefaultColors"
      PrintBackgroundColor=   "16777215"
      PrintGeometryColor=   "0"
      ShadingMode     =   "Gouraud"
      ProjectionMode  =   "Parallel"
      EnableUIMode    =   "DisableRightClickMenu"
      Layout          =   ""
      DisplayMode     =   -1
   End
   Begin VB.Label lblViewer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      MouseIcon       =   "frmShow.frx":64AC
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   765
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12968
      MouseIcon       =   "frmShow.frx":67B6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12960
      MouseIcon       =   "frmShow.frx":6AC0
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   180
      Width           =   510
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Label lblDWF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1140
      TabIndex        =   2
      Top             =   765
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Image imgClose 
      Height          =   945
      Left            =   12540
      Top             =   0
      Width           =   1080
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmShow.frx":6DCA
      MousePointer    =   99  'Custom
      Picture         =   "frmShow.frx":70D4
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Width           =   720
   End
   Begin VB.Image imgViewer 
      Height          =   570
      Left            =   0
      Picture         =   "frmShow.frx":7C1E
      Top             =   600
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Menu mnuShowData 
      Caption         =   "mnuShowData"
      Visible         =   0   'False
      Begin VB.Menu mnuShowName 
         Caption         =   "Show Name:"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowLoc 
         Caption         =   "Show Location:"
      End
      Begin VB.Menu mnuShowOpen 
         Caption         =   "Show Open:"
      End
      Begin VB.Menu mnuShowClose 
         Caption         =   "Show Close:"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttendingClients 
         Caption         =   "Attending Clients..."
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPlan 
         Caption         =   "View Overall Show Plan"
      End
      Begin VB.Menu mnuVignette 
         Caption         =   "View Vignette Floorplan"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuSRA 
      Caption         =   "mnuSRA"
      Visible         =   0   'False
      Begin VB.Menu mnuSRAView 
         Caption         =   "View Show Regulation Abstract"
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuVolo 
      Caption         =   "mnuVolo"
      Visible         =   0   'False
      Begin VB.Menu mnuVPan 
         Caption         =   "Pan"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuVZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuVZoomW 
         Caption         =   "Zoom Window"
      End
      Begin VB.Menu mnuVFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuVMainDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Default Colors"
            Index           =   0
         End
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Black on White"
            Index           =   1
         End
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Clear Scale"
            Index           =   2
         End
      End
      Begin VB.Menu mnuVPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuDash05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmailPDF 
         Caption         =   "Email PDF of Current Plan..."
      End
      Begin VB.Menu mnuDownloadPDF 
         Caption         =   "Download PDF of Current Plan..."
      End
      Begin VB.Menu mnuDownloadDWF 
         Caption         =   "Download Copy of DWF File..."
      End
      Begin VB.Menu mnuDash06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuVCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rDW As Long
'''''Dim Conn As ADODB.Connection
Dim iSHYR As Integer
Dim lSHCD As Long
Dim strOrderBy As String, sCList As String
Dim bTrack As Boolean
Dim cSHNode As String, pSHNode As String
Dim bViewSet As Boolean
Dim dLeft As Double, dRight As Double, dTop As Double, dBottom As Double
Dim E_Mode As Boolean, bAbort As Boolean, bDirsOpen As Boolean
Dim strEmailHdr As String, strEmailMsg As String, strEmailTo As String
Dim sShowDates As String
Dim lCUNO As Long
Dim sFBCN As String, sSHNM As String

    
Private Sub cboCodeName_Change()
    CodeSet cboCodeName, cboCodeName.Text
    If Len(cboCodeName) > 40 Then
        cboCodeName.Text = Left(cboCodeName.Text, 40)
        cboCodeName.SelStart = Len(cboCodeName.Text)
    End If
End Sub

Private Sub cboCodeName_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT * FROM " & SRAEngCodeReq & " " & _
                "WHERE UPPER(ENGCODENAME) = '" & UCase(cboCodeName.Text) & "' " & _
                "AND AN8_SHCD = " & Mid(cSHNode, 3)
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rtbCodeDesc.Text = ""
        lblCodeID = "0"
        cmdCodeUpdate.Enabled = False
        cmdCodeSave.Enabled = True
        cmdCodeDelete.Enabled = False
    Else
        rtbCodeDesc.Text = rst.Fields("ENGCODEDESC")
        lblCodeID = rst.Fields("ENGCODEID")
        cmdCodeUpdate.Enabled = True
        cmdCodeSave.Enabled = False
        cmdCodeDelete.Enabled = True
    End If
    rst.Close: Set rst = Nothing
    
End Sub

'''Private Sub cboCodeName_KeyPress(KeyAscii As Integer)
'''    If Len(cboCodeName) > 40 Then
'''        cboCodeName.Text = Left(cboCodeName.Text, 40)
'''        cboCodeName.SelStart = Len(cboCodeName.Text)
'''    End If
'''End Sub

Private Sub cboCUNO_Click(Index As Integer)
'    Dim lCUNO As Long
    If cboCUNO(Index).Text <> "" And Trim(cboCUNO(Index).Text) <> "<Select by Client>" Then
        Screen.MousePointer = 11
        Select Case Index
'''''            Case 0
'''''                shp1.Visible = False
'''''                If cboCUNO(0).ItemData(cboCUNO(0).ListIndex) = 0 Then
'''''                    sCList = ""
'''''                Else
'''''                    sCList = CStr(cboCUNO(0).ItemData(cboCUNO(0).ListIndex))
'''''                End If
'''''                Call PopShows(sCList)
'''''                cboMatrix.Text = "<Select by Matrix>"
            Case 1
                lCUNO = cboCUNO(1).ItemData(cboCUNO(1).ListIndex)
                If lCUNO = 0 Then
                    sCList = ""
                    sFBCN = "GPJ"
                Else
                    sCList = CStr(lCUNO)
                    sFBCN = UCase(cboCUNO(1).Text)
                End If
                Call PopTree(sCList)
        End Select
        Screen.MousePointer = 0
    End If
End Sub

'''''Private Sub cboMatrix_Click()
'''''    If Left(cboMatrix.Text, 1) <> "<" Then
'''''        Screen.MousePointer = 11
'''''        cboCUNO(0).Text = "<Select by Client>"
'''''        sCList = "SELECT CMY56CUNO FROM " & F5620 & " " & _
'''''                        "WHERE CMY56MATX = " & cboMatrix.ItemData(cboMatrix.ListIndex)
'''''        Call PopShows(sCList)
'''''        Screen.MousePointer = 0
'''''    End If
'''''End Sub

Private Sub cboSHYR_Click(Index As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sClient As String
    
    iSHYR = CInt(cboSHYR(Index).Text)
    Select Case Index
'''''        Case 0
'''''            cboCUNO(0).Clear
'''''            cboCUNO(0).AddItem "<All Shows>"
'''''            cboCUNO(0).ItemData(cboCUNO(0).NewIndex) = 0
'''''            strSelect = "SELECT DISTINCT CS.CSY56CUNO, AB.ABALPH " & _
'''''                        "FROM " & F5611 & " CS, " & F0101 & " AB " & _
'''''                        "Where CS.CSY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND CS.CSY56CUNO = AB.ABAN8 " & _
'''''                        "AND AB.ABAT1 = 'C' " & _
'''''                        "ORDER BY UPPER(AB.ABALPH)"
'''''            Set rst = Conn.Execute(strSelect)
'''''            Do While Not rst.EOF
'''''                cboCUNO(0).AddItem Trim(rst.FIELDS("ABALPH"))
'''''                cboCUNO(0).ItemData(cboCUNO(0).NewIndex) = rst.FIELDS("CSY56CUNO")
'''''                rst.MoveNext
'''''            Loop
'''''            rst.Close
'''''            Set rst = Nothing
        Case 1
            If cboCUNO(1).Text <> "" Then sClient = cboCUNO(1).Text
            cboCUNO(1).Clear
            tvw1.Visible = False: tvw1.Nodes.Clear: tvw1.Visible = True
            cboCUNO(1).AddItem "<All Shows>"
            cboCUNO(1).ItemData(cboCUNO(1).NewIndex) = 0
            If bClientAll_Enabled Then
                strSelect = "SELECT DISTINCT CS.CSY56CUNO, AB.ABALPH " & _
                            "FROM " & F5611 & " CS, " & F0101 & " AB " & _
                            "Where CS.CSY56SHYR = " & cboSHYR(1).Text & " " & _
                            "AND CS.CSY56SHCD > 0 " & _
                            "AND CS.CSY56CUNO > 0 " & _
                            "AND CS.CSY56CUNO = AB.ABAN8 " & _
                            "AND AB.ABAT1 = 'C' " & _
                            "ORDER BY UPPER(AB.ABALPH)"
            Else
                strSelect = "SELECT DISTINCT CS.CSY56CUNO, AB.ABALPH " & _
                            "FROM " & F5611 & " CS, " & F0101 & " AB " & _
                            "Where CS.CSY56SHYR = " & cboSHYR(1).Text & " " & _
                            "AND CS.CSY56SHCD > 0 " & _
                            "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                            "AND CS.CSY56CUNO = AB.ABAN8 " & _
                            "AND AB.ABAT1 = 'C' " & _
                            "ORDER BY UPPER(AB.ABALPH)"
            End If
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                cboCUNO(1).AddItem Trim(rst.Fields("ABALPH"))
                cboCUNO(1).ItemData(cboCUNO(1).NewIndex) = rst.Fields("CSY56CUNO")
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            On Error Resume Next
            If sClient <> "" Then cboCUNO(1).Text = sClient
    End Select
End Sub

Private Sub cmdApprove_Click()
    With frmEngConfirm
        .PassSHCD = Mid(cSHNode, 3)
        .PassSHNM = tvw1.Nodes(cSHNode).Text
        .Show 1
    End With
End Sub

Private Sub imgViewer_Click()
    If volShowplan.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgViewer.Left, imgViewer.Top + imgViewer.Height
    End If
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

'''''Private Sub cmdBlock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''    shp1.Top = cmdBlock(Index).Top - 15
'''''    shp1.Visible = True
'''''    lSHCD = CLng(cmdBlock(Index).Tag)
'''''    Call PopShowInfo(Index, lSHCD)
'''''
'''''End Sub

Private Sub cmdCodeClear_Click()
    cboCodeName.Text = ""
    rtbCodeDesc.Text = ""
    lblCodeID.Caption = "0"
    lblEditUser.Caption = ""
    lblConfUser.Caption = ""
    cmdCodeUpdate.Enabled = False
    cmdCodeSave.Enabled = True
    cmdCodeDelete.Enabled = False
End Sub

Private Sub cmdCodeDelete_Click()
    Dim strDelete As String
    Dim Resp As VbMsgBoxResult
    
    If CLng(lblCodeID) > 0 Then
        Resp = MsgBox("Are you certain you want to Delete the current Code?", _
                    vbExclamation + vbYesNoCancel, cboCodeName.Text & "...")
        If Resp = vbYes Then
            strDelete = "DELETE FROM " & SRAEngCodeReq & " " & _
                        "WHERE ENGCODEID = " & lblCodeID
            Conn.Execute (strDelete)
            tvw1.Nodes.Remove ("EC" & lblCodeID.Caption)
            cmdCodeClear_Click
            '///// SHOULD I DELETE CONFIRMATION IF LAST CODE? \\\\\
        End If
    End If
End Sub

Private Sub cmdCodeSave_Click()
    Dim strInsert As String, sChk As String
    Dim rstN As ADODB.Recordset
    Dim lCID As Long
    Dim nodX As Node
    Dim E_Alert As Boolean
    
    bAbort = False
    If Trim(cboCodeName.Text) <> "" _
                And Trim(rtbCodeDesc.Text) <> "" _
                And CLng(lblCodeID) = 0 Then
        On Error GoTo ErrorTrap
        tvw1.ImageList = ImageList1
        E_Alert = CheckForConfirmation("adding")
        If Not bAbort Then
            Set rstN = Conn.Execute("SELECT " & SRASeq & ".NEXTVAL FROM DUAL")
            lCID = rstN.Fields("NEXTVAL")
            rstN.Close: Set rstN = Nothing
            
            strInsert = "INSERT INTO " & SRAEngCodeReq & " " & _
                        "(ENGCODEID, AN8_SHCD, ENGCODENAME, " & _
                        "ENGCODEDESC, ADDUSER, ADDDTTM, " & _
                        "UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lCID & ", " & Mid(cSHNode, 3) & ", '" & Left(DeGlitch(cboCodeName.Text), 40) & "', " & _
                        "'" & Left(DeGlitch(rtbCodeDesc.Text), 500) & "', '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            
            On Error Resume Next
            sChk = tvw1.Nodes("ES" & Mid(cSHNode, 3)).Text
            If Err Then Set nodX = tvw1.Nodes.Add(cSHNode, tvwChild, "ES" & Mid(cSHNode, 3), "Engineering Code Requirements", 6)
            Set nodX = tvw1.Nodes.Add("ES" & Mid(cSHNode, 3), tvwChild, "EC" & lCID, cboCodeName.Text, 3)
            If E_Alert Then Call ConfirmAlert(strEmailTo, strEmailHdr, strEmailMsg)
        End If
    Else
        MsgBox "Not able to save entered data.", vbExclamation, "Incomplete Data..."
    End If
Exit Sub
ErrorTrap:
    MsgBox "Error: " & Err.Description, vbExclamation, "Error Encountered..."
End Sub

Private Sub cmdCodeUpdate_Click()
    Dim strSelect As String, sMess As String, sConfUser As String, strUpdate As String, sConfShort As String
    Dim Resp As VbMsgBoxResult
    Dim E_Alert As Boolean
    
    E_Alert = False: bAbort = False
    If Trim(cboCodeName.Text) <> "" _
                And Trim(rtbCodeDesc.Text) <> "" _
                And CLng(lblCodeID) > 0 Then
        Screen.MousePointer = 11
        On Error GoTo ErrorTrap
        Conn.BeginTrans
        E_Alert = CheckForConfirmation("editing")

        If bAbort = False Then
            '///// NOW DO THE ACTUAL UPDATE \\\\\
            strUpdate = "UPDATE " & SRAEngCodeReq & " " & _
                        "SET ENGCODENAME = '" & Left(DeGlitch(Trim(cboCodeName.Text)), 40) & "', " & _
                        "ENGCODEDESC = '" & Left(DeGlitch(rtbCodeDesc.Text), 500) & "', " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                        "WHERE ENGCODEID = " & lblCodeID.Caption
            Conn.Execute (strUpdate)
            tvw1.Nodes("EC" & lblCodeID.Caption).Text = cboCodeName.Text
'''            Conn.CommitTrans
            If E_Alert Then Call ConfirmAlert(strEmailTo, strEmailHdr, strEmailMsg)
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
    End If
'''CancelIt:
    Screen.MousePointer = 0
Exit Sub
ErrorTrap:
    Conn.RollbackTrans
    Screen.MousePointer = 0
    MsgBox "Error: " & Err.Description, vbExclamation, "Data Not Saved..."
End Sub

Private Sub imgDirs_Click()
    If picDir.Visible = False Then
        picDir.Visible = True
        bDirsOpen = True
        imgDirs.ToolTipText = "Click to Close File Index"
    Else
        picDir.Visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
'''        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
    End If
End Sub

Private Sub cmdEngCodes_Click()
    If fraEngCodes.Visible = True Then
        fraEngCodes.Visible = False
        cmdEngCodes.Caption = "Edit Engineering Code Regulations"
        E_Mode = False
        cmdCodeClear_Click
    Else
        fraEngCodes.Visible = True
        cmdEngCodes.Caption = "Close Engineering Code Interface"
        cmdCodeUpdate.Enabled = False
        E_Mode = True
    End If
End Sub

Private Sub cmdSRA_Click()
    Dim sFile As String
    
'''    MsgBox lCUNO & ", " & iSHYR & ", " & lSHCD & vbNewLine & sFBCN & vbNewLine & sSHNM
'''    MsgBox "The Show Regulation Abstract Report is currently in development, " & _
'''                "and will be available soon!", vbInformation, "Coming Soon..."
    
    sIFile = "SI-" & CStr((CLng(Format(Now, "h")) * 60 * 60) + (CLng(Format(Now, "n")) * 60) + _
                (CLng(Format(Now, "s")))) & ".htm"
    sFile = PopShowInfo(lCUNO, iSHYR, lSHCD, sSHNM, sFBCN)
    frmHTMLViewer.PassFile = sFile
    frmHTMLViewer.PassFrom = Me.Name
    frmHTMLViewer.PassHDR = sFBCN & " - " & iSHYR & " " & sSHNM
    frmHTMLViewer.Show 1, Me
End Sub

Private Sub cmdViewComposite_Click()
    Screen.MousePointer = 11
    bViewSet = False
    picDir.Visible = False
    bDirsOpen = False
    imgDirs.ToolTipText = "Click to Open File Index..."
'''    Set imgDirs.Picture = imlDirs.ListImages(1).Picture
    volShowplan.src = cmdViewComposite.Tag
    picFrame.Visible = False
    volShowplan.Visible = True
    imgViewer.Visible = True
    lblViewer.Visible = True
    picDir.BackColor = vbBlack
    lblWelcome = iSHYR & " " & UCase(tvw1.Nodes(cSHNode).Text)
    lblDWF = "Overall Composite Showplan"
    Call CheckForPDFs(cmdViewComposite.Tag)
    Screen.MousePointer = 0
End Sub

Private Sub cmdViewShowplan_Click()
'''    MsgBox cmdViewShowplan.Tag
    Screen.MousePointer = 11
    bViewSet = False
    picDir.Visible = False
    bDirsOpen = False
    imgDirs.ToolTipText = "Click to Open File Index..."
'''    Set imgDirs.Picture = imlDirs.ListImages(1).Picture
    volShowplan.src = cmdViewShowplan.Tag
    picFrame.Visible = False
    volShowplan.Visible = True
    imgViewer.Visible = True
    lblViewer.Visible = True
    picDir.BackColor = vbBlack
    lblWelcome = iSHYR & " " & UCase(tvw1.Nodes(cSHNode).Text)
    lblDWF = "Overall Showplan"
    Call CheckForPDFs(cmdViewShowplan.Tag)
    Screen.MousePointer = 0
End Sub

'''''Private Sub flx1_Click()
'''''    Dim rPlace As Single
'''''
'''''    shp1.Top = flx1.RowSel * flx1.RowHeight(0)
'''''    shp1.Visible = True
'''''    rPlace = cmdBlock(flx1.RowSel).Left - ((picInner.Width - cmdBlock(flx1.RowSel).Width) / 2)
'''''    Select Case rPlace
'''''        Case Is < 0
'''''            hsc1.Value = 0
'''''        Case Else
'''''            If rPlace / (rDW * 7) > hsc1.Max Then
'''''                hsc1.Value = hsc1.Max
'''''            Else
'''''                hsc1.Value = rPlace / (rDW * 7)
'''''            End If
'''''    End Select
'''''    lSHCD = CLng(flx1.TextMatrix(flx1.RowSel, 1))
'''''    Call PopShowInfo(flx1.RowSel, lSHCD)
'''''End Sub

Private Sub Form_Load()
    Dim i As Integer, iStart As Integer, iEnd As Integer, iLen As Integer
    Dim l As Long, lCnt As Long
    Dim ConnStr As String, strSelect As String
    Dim rst As ADODB.Recordset
    
'''''    ConnStr = "DSN=JDE;UID=ANNOTATOR_APP_USER;PWD=q2eNqsgHxcKqre3"
'''''    Set Conn = CreateObject("ADODB.Connection")
'''''    Conn.Open (ConnStr)
    
    strOrderBy = "ORDER BY SHY56BEGDT, SHY56NAMA"
    pSHNode = ""
    
    
    iSHYR = CInt(Format(Now, "YYYY"))
    For i = -15 To 1
'''''        cboSHYR(0).AddItem iSHYR + i
        cboSHYR(1).AddItem iSHYR + i
    Next i
''''    cboSHYR(0).Text = iSHYR
    cboSHYR(1).Text = iSHYR
    
'''''    cboCUNO(0).AddItem "<Select by Client>"
'''''    cboCUNO(0).AddItem "<All Shows>"
'''''    cboCUNO(0).ItemData(cboCUNO(0).NewIndex) = 0
''    cboCUNO(1).AddItem "<All Shows>"
''    cboCUNO(1).ItemData(cboCUNO(1).NewIndex) = 0
'''    cboCUNO(0).Text = "<All Shows>"
'''    cboCUNO(1).Text = "<All Shows>"
    
'''''    cmdKey(0).BackColor = vbRed
'''''    cmdKey(1).BackColor = vbYellow
'''''    cmdKey(2).BackColor = vbBlue
'''''    cmdKey(3).BackColor = vbGreen
    
'''    lblByGeorge(0).ForeColor = lGeo_Back '' RGB(30, 30, 21)
'''    lblByGeorge(1).ForeColor = lGeo_Fore '' RGB(100, 100, 68)
    
'''''    '///// POP MATRIX COMBO \\\\\
'''''    cboMatrix.AddItem "<Select by Matrix>"
'''''    strSelect = "SELECT DISTINCT MTX.CMY56MATX MTX, UDC.DRDL01 DSC " & _
'''''                "FROM " & F5620 & " MTX, " & F0005 & " UDC " & _
'''''                "WHERE TRIM(UDC.DRSY) = '06' " & _
'''''                "AND UDC.DRRT = '15' " & _
'''''                "AND TRIM(MTX.CMY56MATX) = TRIM(UDC.DRKY) " & _
'''''                "ORDER BY DSC"
'''''    Set rst = Conn.Execute(strSelect)
'''''    Do While Not rst.EOF
'''''        cboMatrix.AddItem UCase(Trim(rst.FIELDS("DSC")))
'''''        cboMatrix.ItemData(cboMatrix.NewIndex) = CLng(rst.FIELDS("MTX"))
'''''        rst.MoveNext
'''''    Loop
'''''    rst.Close: Set rst = Nothing
'''''
'''''    cboCUNO(0).Text = "<All Shows>"
'    cboCUNO(1).Text = "<All Shows>"
'''    cboMatrix.Text = "<Select by Matrix>"
    
    If cboCUNO(1).ListCount = 2 Then cboCUNO(1).Text = cboCUNO(1).List(1)
    lblWelcome.Caption = "...Ready for your selection..."
    If bPerm(50) Then cmdApprove.Visible = True Else cmdApprove.Visible = False '/// RIGHT TO APPROVE \\\
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(False)
    If bENABLE_PRINTERS Then mnuVPrint.Visible = True Else mnuVPrint.Visible = False
    '\\\\\ -------------------------------------------------------- /////
    
    '///// POP EXISTING ENG CODES \\\\\
    strSelect = "SELECT DISTINCT ENGCODENAME FROM " & SRAEngCodeReq & " " & _
                "ORDER BY ENGCODENAME"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboCodeName.AddItem Trim(rst.Fields("ENGCODENAME"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    bDirsOpen = True
End Sub

'''''Public Sub CreateCmdBlock(iD As Integer, iStart As Integer, iLen As Integer)
'''''    Dim btn As CommandButton
'''''    If iD <= 32767 Then
'''''        Load cmdBlock(iD)
'''''    End If
'''''End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        shpHDR.Width = Me.ScaleWidth
        
        If Me.ScaleHeight > picDir.Top + 2000 Then
            picDir.Height = Me.ScaleHeight - picDir.Top
            picDir.Width = Me.ScaleWidth
            sst1.Height = picDir.ScaleHeight
            sst1.Width = picDir.ScaleWidth
            tvw1.Height = sst1.Height - tvw1.Top - 720 ''600
            lblCount.Top = tvw1.Top + tvw1.Height + 30
            fraButtons.Top = sst1.Height - fraButtons.Height - 120 '' tvw1.Top + tvw1.Height + 60
            picFrame.Width = Me.ScaleWidth - picFrame.Left - 120
            picFrame.Height = Me.ScaleHeight - picFrame.Top - 120
            rtbMess.Width = picDir.ScaleWidth - rtbMess.Left - 180
            rtbMess.Height = picDir.ScaleHeight - rtbMess.Top - 600
        End If
        
        imgClose.Left = Me.ScaleWidth - imgClose.Width
        lblHelp.Left = imgClose.Left + (imgClose.Width / 2) - (lblHelp.Width / 2)
        lblClose.Left = imgClose.Left + (imgClose.Width / 2) - (lblClose.Width / 2)
            
'''        cmdClose.Left = Me.ScaleWidth - 120 - cmdClose.Width
'''        cmdClose.Top = 120
            
        volShowplan.Top = picFrame.Top
        volShowplan.Left = picFrame.Left
        volShowplan.Width = picFrame.Width
        volShowplan.Height = picFrame.Height
        AppWindowState = Me.WindowState
    End If
End Sub

'''''Private Sub hsc1_Change()
'''''    Pic1.Left = hsc1.Value * (-1 * (rDW * 7))
'''''    picHeader.Left = hsc1.Value * (-1 * (rDW * 7))
'''''    Debug.Print "Horiz Scroll Value = " & hsc1.Value
'''''End Sub

'''''Private Sub hsc1_Scroll()
'''''    Pic1.Left = hsc1.Value * (-1 * (rDW * 7))
'''''    picHeader.Left = hsc1.Value * (-1 * (rDW * 7))
'''''End Sub

'''''Private Sub mnuAttendingClients_Click()
'''''    Dim strSelect As String, sMess As String
'''''    Dim rst As ADODB.Recordset
'''''
'''''    sMess = ""
'''''    strSelect = "SELECT AB.ABALPH " & _
'''''                "FROM " & F5611 & " CS, " & F0101 & " AB " & _
'''''                "WHERE CS.CSY56SHYR = " & iSHYR & " " & _
'''''                "AND CS.CSY56SHCD = " & lSHCD & " " & _
'''''                "AND CS.CSY56CUNO = AB.ABAN8 " & _
'''''                "AND AB.ABAT1 = 'C' " & _
'''''                "ORDER BY UPPER(AB.ABALPH)"
'''''    Set rst = Conn.Execute(strSelect)
'''''    Do While Not rst.EOF
'''''        sMess = sMess & Space(5) & Trim(rst.FIELDS("ABALPH")) & vbNewLine
'''''        rst.MoveNext
'''''    Loop
'''''    rst.Close: Set rst = Nothing
'''''    If sMess = "" Then
'''''        sMess = "At this time, there are no Clients attending this Show."
'''''    Else
'''''        sMess = "Attending Clients:" & vbNewLine & sMess
'''''    End If
'''''    MsgBox sMess, vbInformation, mnuShowName.Caption & "...     "
'''''End Sub

'''''Private Sub mnuShowPlan_Click()
'''''    Dim strSelect As String
'''''    Dim rst As ADODB.Recordset
'''''
'''''    Me.MousePointer = 11
'''''    strSelect = "SELECT DD.DWFPATH " & _
'''''                "FROM " & DWGShow & " DS, " & DWGDwf & " DD " & _
'''''                "WHERE DS.AN8_SHCD = " & lSHCD & " " & _
'''''                "AND DS.SHYR = " & iSHYR & " " & _
'''''                "AND DS.DWGID = DD.DWGID " & _
'''''                "AND DD.DWFTYPE = 30"
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
''''''''        MsgBox Trim(rst.Fields("DWFPATH"))
'''''        picDir.Visible = False
'''''        imgDirs.tooltiptext = "Open File Index..."
'''''        volShowplan.src = Trim(rst.FIELDS("DWFPATH"))
'''''        picFrame.Visible = False
'''''        volShowplan.Visible = True
'''''        picDir.BackColor = vbBlack
'''''        lblWelcome = cboSHYR(0).Text & " " & UCase(mnuShowName.Caption)
'''''        lblDWF = "Overall Showplan"
'''''    Else
'''''        MsgBox "DWF File not found."
'''''    End If
'''''    rst.Close
'''''    Set rst = Nothing
'''''    Me.MousePointer = 0
'''''End Sub

Private Sub Image1_Click(Index As Integer)
    MsgBox Image1(Index).ToolTipText, vbInformation, "Icon Key..."
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblClose.ForeColor = vbWhite
End Sub

Private Sub lblHelp_Click()
    lblHelp.ForeColor = vbWhite '' lColor
    frmHelp.Show 1
End Sub


Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblHelp.ForeColor = vbWhite '' lColor
End Sub


Private Sub lblViewer_Click()
    If volShowplan.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgViewer.Left, imgViewer.Top + imgViewer.Height
    End If
End Sub

Private Sub mnuDownloadDWF_Click()
    With frmBrowse
        .PassFrom = Me.Name
        .PassBCC = lCUNO
        .PassFBCN = sFBCN
        .PassSHYR = iSHYR
        .PassSHCD = CLng(Mid(cSHNode, 3))
        .PassSHNM = UCase(tvw1.Nodes(cSHNode).Text)
        .PassDWGID = 0
        .PassFILETYPE = "DWF"
        .Show 1
    End With
End Sub

Private Sub mnuDownloadPDF_Click()
    With frmBrowse
        .PassFrom = Me.Name
        .PassBCC = lCUNO
        .PassFBCN = sFBCN
        .PassSHYR = iSHYR
        .PassSHCD = CLng(Mid(cSHNode, 3))
        .PassSHNM = UCase(tvw1.Nodes(cSHNode).Text)
        .PassDWGID = 0
        .PassFILETYPE = "PDF"
''        .PassFrom = Me.Name
        .Show 1
    End With
End Sub

Private Sub mnuEmailPDF_Click()
    With frmEmailFile
        .PassBCC = lCUNO
'''        .PassFBCN = ""
        .PassSHYR = iSHYR
        .PassSHCD = CLng(Mid(cSHNode, 3))
        .PassSHNM = UCase(tvw1.Nodes(cSHNode).Text)
        .PassDWGID = 0
        .PassFILETYPE = "PDF"
        .PassSHDT = sShowDates
        .PassFrom = Me.Name
        .Show 1
    End With

End Sub

Private Sub mnuVDisplay_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        If i = Index Then mnuVDisplay(i).Checked = True Else mnuVDisplay(i).Checked = False
    Next i
    Select Case Index
    Case 0
        volShowplan.GeometryColor = "DefaultColors"
        volShowplan.BackgroundColor = "DefaultColors"
    Case 1
        volShowplan.GeometryColor = vbBlack
        volShowplan.BackgroundColor = vbWhite
    Case 2
        volShowplan.GeometryColor = "ClearScale"
        volShowplan.BackgroundColor = "ClearScale"
    End Select
End Sub

Private Sub mnuVFullView_Click()
    volShowplan.SetCurrentView dLeft, dRight, dBottom, dTop
End Sub

Private Sub mnuVLayers_Click()
    volShowplan.ShowLayersDialog
End Sub

Private Sub mnuVPan_Click()
    ClearChecks
    mnuVPan.Checked = True
    volShowplan.UserMode = "Pan"
End Sub

Private Sub mnuVPrint_Click()
    volShowplan.ShowPrintDialog
End Sub

Private Sub mnuVZoom_Click()
    ClearChecks
    mnuVZoom.Checked = True
    volShowplan.UserMode = "Zoom"
End Sub

Private Sub mnuVZoomW_Click()
    ClearChecks
    mnuVZoomW.Checked = True
    volShowplan.UserMode = "ZoomToRect"
End Sub

Private Sub optSort_Click(Index As Integer)
    Screen.MousePointer = 11
    optSort(0).Refresh: optSort(1).Refresh
    Call PopTree(sCList)
    Screen.MousePointer = 0
End Sub

'''''Private Sub optOrder_Click(Index As Integer)
'''''    Screen.MousePointer = 11
'''''    shp1.Visible = False
'''''    Select Case Index
'''''        Case 0: strOrderBy = "ORDER BY SHY56BEGDT, SHY56NAMA"
'''''        Case 1: strOrderBy = "ORDER BY UPPER(SHY56NAMA)"
'''''    End Select
'''''    Call PopShows(sCList) '''(cboCUNO(0).ItemData(cboCUNO(0).ListIndex))
'''''    Screen.MousePointer = 0
'''''End Sub

'''''Private Sub picTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''    bTrack = True
'''''End Sub

'''''Private Sub picTrack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''    Dim Xt As Single, Yt As Single
'''''    If bTrack Then
'''''        Debug.Print "X = " & X & ",   Y = " & Y
'''''        If X > picTrack.ScaleWidth Then
'''''            Xt = Pic1.Width
'''''        ElseIf X < 0 Then
'''''            Xt = 0
'''''        Else
'''''            Xt = X
'''''        End If
'''''        If Y > picTrack.ScaleHeight Then
'''''            Yt = Pic1.Height
'''''        ElseIf Y < 0 Then
'''''            Yt = 0
'''''        Else
'''''            Yt = Y
'''''        End If
'''''        Debug.Print "X = " & Xt & ",   Y = " & Yt
'''''        Pic1.Left = Xt * -1
'''''        picHeader.Left = Xt * -1
'''''        picInner.Top = Yt * -1
'''''        flx1.Top = Yt * -1
'''''    End If
'''''End Sub

'''''Private Sub picTrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''    bTrack = False
'''''    Select Case X
'''''        Case Is < 0
'''''            hsc1.Value = 0
'''''        Case Is > picTrack.ScaleWidth
'''''            hsc1.Value = hsc1.Max
'''''        Case Else
'''''            If X / (rDW * 7) > hsc1.Max Then
'''''                hsc1.Value = hsc1.Max
'''''            Else
'''''                hsc1.Value = X / (rDW * 7)
'''''            End If
'''''    End Select
'''''    Select Case Y
'''''        Case Is < 0
'''''            vsc1.Value = 0
'''''        Case Is > picTrack.ScaleHeight
'''''            vsc1.Value = vsc1.Max
'''''        Case Else
'''''            If Y / flx1.RowHeight(0) > vsc1.Max Then
'''''                vsc1.Value = vsc1.Max
'''''            Else
'''''                vsc1.Value = Y / flx1.RowHeight(0)
'''''            End If
'''''    End Select
'''''End Sub

'''''Private Sub sst1_Click(PreviousTab As Integer)
'''''    shp1.Visible = False
'''''End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strSelect As String, sMess As String, sName As String, strFrom As String
    Dim rst As ADODB.Recordset
    Dim lFCCD As Long, lHallID As Long, lResID As Long, lEaseID As Long, _
                lSMID As Long, lGCID As Long ''', lSHCD As Long
    Dim i As Integer, iStr As Integer, iEnd As Integer, iDash As Integer
    Dim lErr As Long
    
'    lErr = LockWindowUpdate(Me.hwnd)
    
    Select Case Left(Node.Key, 2)
        Case "EA": cSHNode = Node.Parent.Parent.Parent.Key
        Case "LO", "ES": cSHNode = Node.Parent.Key
        Case "EC": cSHNode = Node.Parent.Parent.Key
        Case Else
            iDash = InStr(1, Node.Key, "-")
            If iDash = 0 Then cSHNode = Node.Key _
                        Else cSHNode = "SH" & Mid(Node.Key, 3, InStr(3, Node.Key, "-") - 3)
    End Select
    
    If pSHNode <> "" And pSHNode <> cSHNode Then
'''        tvw1.Nodes(pSHNode).Bold = False
        lblMess = "": rtbMess.Text = ""
'''        tvw1.Nodes(pSHNode).Expanded = False
        If Left(Node.Key, 2) <> "EC" And E_Mode Then
            fraEngCodes.Visible = False
            cmdEngCodes.Caption = "Edit Engineering Code Regulations"
            E_Mode = False
            cmdCodeClear_Click
        End If
    End If
'''    tvw1.Nodes(cSHNode).Bold = True
    lblShow = tvw1.Nodes(cSHNode).Text '''DblAmp(tvw1.Nodes(cSHNode).Text)
    pSHNode = cSHNode
    
'    If bPerm(49) Then cmdEngCodes.Visible = True Else cmdEngCodes.Visible = False '/// ACCESS TO CODES \\\
    
    lblMess.Visible = True
    If lCUNO > 0 Then cmdSRA.Visible = True Else cmdSRA.Visible = False
    
    cmdSRA.ToolTipText = "Click to view entire Abstract for '" & tvw1.Nodes(cSHNode).Text & "'"
    Call CheckForShowplan(iSHYR, CLng(Mid(cSHNode, 3)))
    
    Select Case UCase(Left(Node.Key, 2))
        Case "SH"
            lSHCD = CLng(Mid(Node.Key, 3))
            sSHNM = Node.Text
            Call PopShowInfo2(Mid(cSHNode, 3))
            Node.Expanded = True
        Case "FC"
            sName = UCase(Mid(Trim(Node.Text), 12)) ''', InStr(12, Node.Text, "   ") - 12))
            sMess = sName & vbNewLine
            lFCCD = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
            lSHCD = CLng(Mid(Node.Parent.Key, 3))
            sSHNM = Node.Parent.Text
            strSelect = "SELECT ALADD1, ALADD2, ALADD3, ALADD4, " & _
                        "ALCTY1, ALADDS, ALADDZ " & _
                        "FROM " & F0116 & " " & _
                        "WHERE ALAN8 = " & lFCCD & " " & _
                        "AND ALEFTB IN " & _
                        "(SELECT MAX(ALEFTB) " & _
                        "FROM " & F0116 & " " & _
                        "WHERE ALAN8 = " & lFCCD & ")"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                For i = 0 To 3
                    If Not IsNull(rst.Fields(i)) And Trim(rst.Fields(i)) <> "" Then
                        sMess = sMess & UCase(Trim(rst.Fields(i))) & vbNewLine
                    End If
                Next i
                sMess = sMess & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                            UCase(Trim(rst.Fields("ALADDS"))) & _
                            "  " & Trim(rst.Fields("ALADDZ")) & vbNewLine
            End If
            rst.Close: Set rst = Nothing
            
            sMess = sMess & vbNewLine
            strSelect = "SELECT WPPHTP, WPAR1, WPPH1 " & _
                        "FROM " & F0115 & " " & _
                        "WHERE WPAN8 = " & lFCCD & " " & _
                        "ORDER BY WPPHTP"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sMess = sMess & "Phone(s):" & vbNewLine
                Do While Not rst.EOF
                    If IsNull(rst.Fields("WPPHTP")) Or Trim(rst.Fields("WPPHTP")) = "" Then
                        sMess = sMess & Space(3) & "BUS:" & vbTab & _
                                    Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                    vbNewLine
                    Else
                        sMess = sMess & Space(3) & Trim(rst.Fields("WPPHTP")) & ":" & vbTab & _
                                    Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                    vbNewLine
                    End If
                    rst.MoveNext
                Loop
            End If
            rst.Close: Set rst = Nothing
            lblMess = "Facility Address:": lblMess.Visible = True
            rtbMess.Text = sMess
        Case "HA"
            lSHCD = CLng(Mid(Node.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Text
            If Node.Image = 3 Then
                Debug.Print Node.Key
                
                lHallID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT HALLNOTE FROM " & SRAHallMas & " " & _
                            "WHERE HALLID = " & lHallID
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    lblMess = "Hall Note:": lblMess.Visible = True
                    rtbMess.Text = Trim(rst.Fields("HALLNOTE"))
                End If
                rst.Close: Set rst = Nothing
            Else
                lblMess.Caption = ""
                rtbMess = ""
            End If
        Case "CH"
            lSHCD = CLng(Mid(Node.Parent.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Parent.Text
            If Node.Image = 3 Then
                Debug.Print Node.Key
                lHallID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT CLGNOTE FROM " & SRAHallMas & " " & _
                            "WHERE HALLID = " & lHallID
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    lblMess = "Ceiling Note:": lblMess.Visible = True
                    rtbMess.Text = Trim(rst.Fields("CLGNOTE"))
                End If
                rst.Close: Set rst = Nothing
            Else
                lblMess.Caption = ""
                rtbMess = ""
            End If
        Case "HR"
            lSHCD = CLng(Mid(Node.Parent.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Parent.Text
            If Node.Image = 3 Then
                Debug.Print Node.Key
                lResID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT RESNOTE FROM " & SRAHallRes & " " & _
                            "WHERE RESID = " & lResID
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    lblMess = "Show Restriction Note:": lblMess.Visible = True
                    rtbMess.Text = Trim(rst.Fields("RESNOTE"))
                End If
                rst.Close: Set rst = Nothing
            Else
                lblMess.Caption = ""
                rtbMess = ""
            End If
        Case "ED"
            lSHCD = CLng(Mid(Node.Parent.Parent.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Parent.Parent.Text
            If Node.Image = 3 Then
                Debug.Print Node.Key
                lEaseID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT EASEDESC FROM " & SRAEase & " " & _
                            "WHERE EASEID = " & lEaseID
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    iStr = InStr(1, Node.Text, "[") + 1
                    iEnd = InStr(1, Node.Text, "]")
                    sMess = Mid(Node.Text, iStr, iEnd - iStr)
                    lblMess = "'" & sMess & "' Description:": lblMess.Visible = True
                    rtbMess.Text = Trim(rst.Fields("EASEDESC"))
                End If
                rst.Close: Set rst = Nothing
            Else
                lblMess.Caption = ""
                rtbMess = ""
            End If
        Case "SM"
            lSHCD = CLng(Mid(Node.Parent.Key, 3))
            sSHNM = Node.Parent.Text
            If Node.Image = 3 Then
                sMess = ""
                lSMID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
                            "AL.ALCTY1, AL.ALADDS, AL.ALADDZ " & _
                            "FROM " & F0101 & " AB, " & F0116 & " AL " & _
                            "WHERE AB.ABAN8 = " & lSMID & " " & _
                            "AND AB.ABAN8 = AL.ALAN8"
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    sMess = UCase(Trim(rst.Fields("ABALPH"))) & vbNewLine
                    For i = 1 To 4
                        If Not IsNull(rst.Fields(i)) And Trim(rst.Fields(i)) <> "" Then
                            sMess = sMess & UCase(Trim(rst.Fields(i))) & vbNewLine
                        End If
                    Next i
                    sMess = sMess & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                                UCase(Trim(rst.Fields("ALADDS"))) & _
                                "  " & Trim(rst.Fields("ALADDZ")) & vbNewLine
                    rst.Close: Set rst = Nothing
                    
                    sMess = sMess & vbNewLine
                    strSelect = "SELECT WPPHTP, WPAR1, WPPH1 " & _
                                "FROM " & F0115 & " " & _
                                "WHERE WPAN8 = " & lSMID & " " & _
                                "ORDER BY WPPHTP"
                    Set rst = Conn.Execute(strSelect)
                    If Not rst.EOF Then
                        sMess = sMess & "Phone(s):" & vbNewLine
                        Do While Not rst.EOF
                            If IsNull(rst.Fields("WPPHTP")) Or Trim(rst.Fields("WPPHTP")) = "" Then
                                sMess = sMess & Space(3) & "BUS:" & vbTab & _
                                            Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                            vbNewLine
                            Else
                                sMess = sMess & Space(3) & Trim(rst.Fields("WPPHTP")) & ":" & vbTab & _
                                            Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                            vbNewLine
                            End If
                            rst.MoveNext
                        Loop
                    End If
                    rst.Close: Set rst = Nothing
                    lblMess = "Show Manager:": lblMess.Visible = True
                    rtbMess.Text = sMess
                Else
                    rst.Close: Set rst = Nothing
                End If
                
            End If
        
        Case "GC"
            lSHCD = CLng(Mid(Node.Parent.Key, 3))
            sSHNM = Node.Parent.Text
            If Node.Image = 3 Then
                sMess = ""
                lGCID = CLng(Mid(Node.Key, InStr(1, Node.Key, "-") + 1))
                strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
                            "AL.ALCTY1, AL.ALADDS, AL.ALADDZ " & _
                            "FROM " & F0101 & " AB, " & F0116 & " AL " & _
                            "WHERE AB.ABAN8 = " & lGCID & " " & _
                            "AND AB.ABAN8 = AL.ALAN8"
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    sMess = UCase(Trim(rst.Fields("ABALPH"))) & vbNewLine
                    For i = 1 To 4
                        If Not IsNull(rst.Fields(i)) And Trim(rst.Fields(i)) <> "" Then
                            sMess = sMess & UCase(Trim(rst.Fields(i))) & vbNewLine
                        End If
                    Next i
                    sMess = sMess & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                                UCase(Trim(rst.Fields("ALADDS"))) & _
                                "  " & Trim(rst.Fields("ALADDZ")) & vbNewLine
                    rst.Close: Set rst = Nothing
                    
                    sMess = sMess & vbNewLine
                    strSelect = "SELECT WPPHTP, WPAR1, WPPH1 " & _
                                "FROM " & F0115 & " " & _
                                "WHERE WPAN8 = " & lSMID & " " & _
                                "ORDER BY WPPHTP"
                    Set rst = Conn.Execute(strSelect)
                    If Not rst.EOF Then
                        sMess = sMess & "Phone(s):" & vbNewLine
                        Do While Not rst.EOF
                            If IsNull(rst.Fields("WPPHTP")) Or Trim(rst.Fields("WPPHTP")) = "" Then
                                sMess = sMess & Space(3) & "BUS:" & vbTab & _
                                            Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                            vbNewLine
                            Else
                                sMess = sMess & Space(3) & Trim(rst.Fields("WPPHTP")) & ":" & vbTab & _
                                            Trim(rst.Fields("WPAR1")) & " " & Trim(rst.Fields("WPPH1")) & _
                                            vbNewLine
                            End If
                            rst.MoveNext
                        Loop
                    End If
                    rst.Close: Set rst = Nothing
                    lblMess = "Show Contractor:": lblMess.Visible = True
                    rtbMess.Text = sMess
                Else
                    rst.Close: Set rst = Nothing
                End If
                
            End If
        Case "AC"
            lSHCD = CLng(Mid(Node.Parent.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Parent.Text
            iDash = InStr(1, Node.Key, "-")
'            lSHCD = CLng(Mid(Node.key, 3, iDash - 3))
            lHallID = CLng(Mid(Node.Key, iDash + 1))
            sMess = ""
            If bClientAll_Enabled Then
                strSelect = "SELECT AB.ABALPH " & _
                            "FROM " & SRACliHall & " CH, " & F5611 & " CS, " & F0101 & " AB " & _
                            "WHERE CS.CSY56SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                            "AND CS.CSY56SHCD = " & lSHCD & " " & _
                            "AND CS.CSY56SHYR = CH.SHYR " & _
                            "AND CS.CSY56SHCD = CH.AN8_SHCD " & _
                            "AND CS.CSY56CUNO = CH.AN8_CUNO " & _
                            "AND CH.HALLID = " & lHallID & " " & _
                            "AND CH.AN8_CUNO = AB.ABAN8 " & _
                            "AND AB.ABAT1 = 'C' " & _
                            "ORDER BY UPPER(AB.ABALPH)"
            Else
                strSelect = "SELECT AB.ABALPH " & _
                            "FROM " & SRACliHall & " CH, " & F5611 & " CS, " & F0101 & " AB " & _
                            "WHERE CS.CSY56SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                            "AND CS.CSY56SHCD = " & lSHCD & " " & _
                            "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                            "AND CS.CSY56SHYR = CH.SHYR " & _
                            "AND CS.CSY56SHCD = CH.AN8_SHCD " & _
                            "AND CS.CSY56CUNO = CH.AN8_CUNO " & _
                            "AND CH.HALLID = " & lHallID & " " & _
                            "AND CH.AN8_CUNO = AB.ABAN8 " & _
                            "AND AB.ABAT1 = 'C' " & _
                            "ORDER BY UPPER(AB.ABALPH)"
            End If
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sMess = sMess & Trim(rst.Fields("ABALPH")) & vbNewLine
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            lblMess = "Hall Clients:": lblMess.Visible = True
            rtbMess.Text = sMess
            
        Case "EC"
            '''GET DESC, ADD INFO & CONFIRM INFO
            '''IF E_MODE THEN WRITE TO EDITOR ELSE WRITE TO LBLMESS
            strFrom = "(SELECT CC.AN8_SHCD, CC.SHYR, " & _
                        "U.NAME_FIRST, U.NAME_LAST, " & _
                        "TO_CHAR(CC.CONFIRMDTTM, 'MON DD, YYYY') AS CONFDATE " & _
                        "FROM " & SRAEngCodeConf & " CC, " & IGLUser & " U " & _
                        "WHERE CC.CONFIRMSTATUS > 0 " & _
                        "AND CC.CONFIRMUSER = U.NAME_LOGON)"
            strSelect = "SELECT CR.ENGCODEID, CR.ENGCODEDESC, CR.UPDUSER, " & _
                        "TO_CHAR(CR.UPDDTTM, 'MON DD, YYYY') AS EDITDATE, " & _
                        "CC.NAME_FIRST, CC.NAME_LAST, CC.CONFDATE, CC.SHYR " & _
                        "FROM " & SRAEngCodeReq & " CR, " & strFrom & " CC " & _
                        "WHERE CR.ENGCODEID = " & Mid(Node.Key, 3) & " " & _
                        "AND CR.AN8_SHCD = CC.AN8_SHCD (+) " & _
                        "ORDER BY CC.SHYR DESC"
            Set rst = Conn.Execute(strSelect)
            If E_Mode Then
                cboCodeName.Text = Node.Text
                If Not rst.EOF Then
                    rtbCodeDesc.TextRTF = Trim(rst.Fields("ENGCODEDESC"))
                    lblCodeID = rst.Fields("ENGCODEID")
                    lblEditUser = "Last Edit: " & Trim(rst.Fields("EDITDATE")) & " by " & _
                                Trim(rst.Fields("UPDUSER"))
                    If Not IsNull(rst.Fields("NAME_LAST")) Then
                        lblConfUser = "Confirmed for " & rst.Fields("SHYR") & " Show on " & _
                                    Trim(rst.Fields("CONFDATE")) & " by " & Trim(rst.Fields("NAME_FIRST")) & _
                                    " " & Trim(rst.Fields("NAME_LAST"))
                    Else
                        lblConfUser = "No Confirmation has been done."
                    End If
                End If
                cmdCodeUpdate.Enabled = True
                cmdCodeSave.Enabled = False
                cmdCodeDelete.Enabled = True
            Else
                lblMess = Node.Text
                sMess = "Last Edit: " & Trim(rst.Fields("EDITDATE")) & " by " & _
                            Trim(rst.Fields("UPDUSER")) & vbNewLine & vbNewLine
                If Not IsNull(rst.Fields("NAME_LAST")) Then
                    sMess = sMess & "Confirmed for " & rst.Fields("SHYR") & " Show on " & _
                                Trim(rst.Fields("CONFDATE")) & " by " & Trim(rst.Fields("NAME_FIRST")) & _
                                " " & Trim(rst.Fields("NAME_LAST")) & vbNewLine & vbNewLine
                End If
                sMess = sMess & Trim(rst.Fields("ENGCODEDESC"))
                rtbMess.TextRTF = sMess
            End If
            rst.Close: Set rst = Nothing
            
        Case "ES"
            
        Case "EA"
            lSHCD = CLng(Mid(Node.Parent.Parent.Parent.Key, 3))
            sSHNM = Node.Parent.Parent.Parent.Text
            lblMess.Caption = ""
            Me.rtbMess = ""
            
        Case "LO"
            lSHCD = CLng(Mid(Node.Parent.Key, 3))
            sSHNM = Node.Parent.Text
            lblMess.Caption = ""
            Me.rtbMess = ""
            
    End Select
'    lErr = LockWindowUpdate(0)
    
End Sub

Private Sub volShowplan_MouseDown(Button As Integer, Shift As Integer, x As Double, y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuVolo
    End If
End Sub

Private Sub volShowplan_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            Call InitialView
            bViewSet = True
        End If
    End If
End Sub

'''''Private Sub vsc1_Change()
'''''    picInner.Top = (vsc1.Value * (-1 * flx1.RowHeight(0)))
'''''    flx1.Top = vsc1.Value * (-1 * flx1.RowHeight(0))
'''''End Sub
'''''
'''''Private Sub vsc1_Scroll()
'''''    picInner.Top = (vsc1.Value * (-1 * flx1.RowHeight(0)))
'''''    flx1.Top = vsc1.Value * (-1 * flx1.RowHeight(0))
'''''End Sub

'''''Public Function PopShows(sCUNO As String)
'''''    Dim i As Integer, iStart As Integer, iEnd As Integer, iLen As Integer
'''''    Dim l As Long, lCnt As Long, lMon As Long, lMonLen As Long, lDay As Long
'''''    Dim strSelect As String, sMon As String, sDay As String
'''''    Dim rst As ADODB.Recordset
'''''    Dim dStart As Date
'''''
'''''
'''''    flx1.Visible = False
'''''    flx1.Rows = 1
'''''    Pic1.Visible = False
'''''    Pic1.Cls
'''''    If sCUNO = "" Or sCUNO = "0" Then
'''''        strSelect = "SELECT COUNT(*) " & _
'''''                    "FROM " & F5601 & " " & _
'''''                    "WHERE SHY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                    "AND SHY56BEGDT <> 0 " & _
'''''                    "AND SHY56ENDDT <> 0"
'''''    Else
'''''        strSelect = "SELECT COUNT(DISTINCT CSY56SHCD) " & _
'''''                    "FROM " & F5611 & " " & _
'''''                    "WHERE CSY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                    "AND CSY56CUNO IN (" & sCUNO & ")"
'''''    End If
'''''    Set rst = Conn.Execute(strSelect)
'''''    lCnt = rst.FIELDS(0)
'''''    rst.Close: Set rst = Nothing
'''''
'''''    If lCnt > 0 Then
'''''        lblCount = lCnt & " Shows"
'''''        For i = 0 To cmdBlock.Count - 1
'''''            cmdBlock(i).Visible = False
'''''        Next i
'''''        If lCnt >= cmdBlock.Count - 1 Then
'''''    '''        For i = CInt(lCnt) To cmdBlock.Count - 1
'''''    '''            Unload cmdBlock(i)
'''''    '''        Next i
'''''    '''    Else
'''''            For i = cmdBlock.Count To lCnt - 1
'''''                Load cmdBlock(i)
'''''            Next i
'''''        End If
'''''
'''''
'''''        rDW = 80
'''''        flx1.ColWidth(0) = flx1.Width
'''''        flx1.ColWidth(1) = 0
'''''        flx1.ColAlignment(0) = 1
'''''        flx1.Rows = lCnt
'''''        flx1.Height = flx1.Rows * flx1.RowHeight(0)
'''''        Pic1.Height = flx1.Height
'''''        picInner.Height = flx1.Height
'''''        picTrack.ScaleWidth = Pic1.ScaleWidth
'''''        picTrack.ScaleHeight = Pic1.ScaleHeight
'''''        Debug.Print "PicX = " & Pic1.Width & ",    PicY = " & Pic1.Height
'''''        vsc1.Max = flx1.Rows - Int(picOuter.Height / flx1.RowHeight(0))
'''''        If vsc1.Max < 0 Then vsc1.Max = 1
'''''        vsc1.Value = 0
'''''        hsc1.Max = 54 - Int(picInner.Width / (rDW * 7))
'''''        hsc1.Value = 0
'''''
'''''
'''''        Pic1.AutoRedraw = True
'''''        Pic1.DrawStyle = 5
'''''
'''''        '///// SHADE WEEK BARS \\\\\
'''''        l = 0
'''''        Do While l < 365
'''''            Pic1.Line (l * rDW, 0)-((l * rDW) + (rDW * 7), Pic1.Height), 0, B
'''''            l = l + 14
'''''        Loop
'''''
'''''        '///// DRAW MONTH LINES AND WRITE MONTHS \\\\\
'''''        Pic1.DrawStyle = 0
'''''        Pic1.DrawWidth = 2
'''''        picHeader.DrawStyle = 0
'''''        picHeader.DrawWidth = 2
'''''        For l = 0 To 12
'''''            lMon = CLng(format(DateAdd("m", l, DateValue("01/01/" & iSHYR)), "y"))
'''''            If l = 12 Then lMon = lMon + 365
'''''            Pic1.Line (lMon * rDW, 0)-(lMon * rDW, Pic1.Height)
'''''            picHeader.Line (lMon * rDW, 0)-(lMon * rDW, picHeader.Height)
'''''            Select Case l + 1
'''''                Case Is > 0
'''''                    sMon = format(DateValue(CStr(l + 1) & "/1/2000"), "mmmm")
'''''                    lblTest = sMon
'''''                    lMonLen = lblTest.Width
'''''                    picHeader.CurrentX = (lMon * rDW) + (15 * rDW) - (lMonLen / 2)
'''''                    picHeader.CurrentY = 30
'''''                    picHeader.Print sMon
'''''            End Select
'''''        Next l
'''''
'''''        '///// WRITE FIRST DAY OF EACH WEEK \\\\\
'''''        dStart = DateValue("12/31/" & iSHYR - 1)
'''''        For l = 1 To 52
'''''            lDay = CLng(format(DateAdd("ww", l, dStart), "y"))
'''''            sDay = CStr(format(DateAdd("ww", l, dStart), "d"))
'''''            picHeader.CurrentX = lDay * rDW
'''''            picHeader.CurrentY = 240
'''''            picHeader.Print sDay
'''''        Next l
'''''
'''''        i = 0
'''''        If sCUNO = "" Or sCUNO = "0" Then
'''''            strSelect = "SELECT SM.SHY56SHCD, SM.SHY56NAMA, SM.SHY56BEGDT, SM.SHY56ENDDT, " & _
'''''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
'''''                        "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
'''''                        "OAS.SHOW_ID AS OASID, VIN.SHOW_ID AS VINID " & _
'''''                        "FROM " & F5601 & " SM, " & _
'''''                        "(SELECT DS.SHOW_ID, DS.SHYR, DS.AN8_SHCD FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
'''''                        "WHERE DS.SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND DS.DWGID = DM.DWGID " & _
'''''                        "AND DM.DWGTYPE = 3) OAS, " & _
'''''                        "(SELECT DS.SHOW_ID, DS.SHYR, DS.AN8_SHCD FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
'''''                        "WHERE DS.SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND DS.DWGID = DM.DWGID " & _
'''''                        "AND DM.DWGTYPE = 4) VIN " & _
'''''                        "WHERE SM.SHY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND SM.SHY56BEGDT <> 0 " & _
'''''                        "AND SM.SHY56ENDDT <> 0 " & _
'''''                        "AND SM.SHY56SHYR = OAS.SHYR (+) " & _
'''''                        "AND SM.SHY56SHCD = OAS.AN8_SHCD (+) " & _
'''''                        "AND SM.SHY56SHYR = VIN.SHYR (+) " & _
'''''                        "AND SM.SHY56SHCD = VIN.AN8_SHCD (+) "
'''''    '''        strOrderBy = "ORDER BY SHY56BEGDT, SHY56NAMA"
'''''        Else
'''''            strSelect = "SELECT DISTINCT SM.SHY56SHCD, SM.SHY56NAMA, SM.SHY56BEGDT, SM.SHY56ENDDT, " & _
'''''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
'''''                        "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
'''''                        "OAS.SHOW_ID AS OASID, VIN.SHOW_ID AS VINID " & _
'''''                        "FROM " & F5611 & " CS, " & F5601 & " SM, " & _
'''''                        "(SELECT DS.SHOW_ID, DS.SHYR, DS.AN8_SHCD FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
'''''                        "WHERE DS.SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND DS.DWGID = DM.DWGID " & _
'''''                        "AND DM.DWGTYPE = 3) OAS, " & _
'''''                        "(SELECT DS.SHOW_ID, DS.SHYR, DS.AN8_SHCD FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
'''''                        "WHERE DS.SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND DS.DWGID = DM.DWGID " & _
'''''                        "AND DM.DWGTYPE = 4) VIN " & _
'''''                        "WHERE CS.CSY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                        "AND CS.CSY56CUNO IN (" & sCUNO & ") " & _
'''''                        "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
'''''                        "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
'''''                        "AND SM.SHY56SHYR = OAS.SHYR (+) " & _
'''''                        "AND SM.SHY56SHCD = OAS.AN8_SHCD (+) " & _
'''''                        "AND SM.SHY56SHYR = VIN.SHYR (+) " & _
'''''                        "AND SM.SHY56SHCD = VIN.AN8_SHCD (+) "
'''''    '''        strOrderBy = "ORDER BY SHY56BEGDT, SHY56NAMA"
'''''        End If
'''''        Set rst = Conn.Execute(strSelect & strOrderBy)
'''''        Do While Not rst.EOF
'''''            flx1.TextMatrix(i, 0) = Trim(rst.FIELDS("SHY56NAMA"))
'''''            flx1.TextMatrix(i, 1) = rst.FIELDS("SHY56SHCD")
'''''            iStart = CInt(Right(rst.FIELDS("SHY56BEGDT"), 3))
'''''            iEnd = CInt(Right(rst.FIELDS("SHY56ENDDT"), 3))
'''''            If iEnd < iStart Then iEnd = iEnd + 365
'''''            iLen = iEnd - iStart + 1
'''''
'''''    '''''        Select Case i
'''''    ''''''''            Case 0
'''''    ''''''''                cmdBlock(0).Top = i * flx1.RowHeight(0) + 15
'''''    ''''''''                cmdBlock(0).Left = iStart * rDW
'''''    ''''''''                cmdBlock(0).Width = iLen * 80
'''''    ''''''''                cmdBlock(0).Height = flx1.RowHeight(0) - 30
'''''    ''''''''                cmdBlock(0).Visible = True
'''''    ''''''''                cmdBlock(0).ToolTipText = Trim(rst.Fields("BEG_DATE")) & " --> " & _
'''''    ''''''''                            Trim(rst.Fields("END_DATE"))
'''''    '''''            Case Is > 0
'''''    '''''                Load cmdBlock(i)
'''''    ''''''''                Call CreateCmdBlock(i, iStart, iLen)
'''''    ''''''''                cmdBlock(i).ToolTipText = Trim(rst.Fields("BEG_DATE")) & " --> " & _
'''''    ''''''''                            Trim(rst.Fields("END_DATE"))
'''''    '''''        End Select
'''''            cmdBlock(i).Top = i * flx1.RowHeight(0) + 15
'''''            cmdBlock(i).Left = iStart * rDW
'''''            cmdBlock(i).Width = iLen * 80
'''''            cmdBlock(i).Height = flx1.RowHeight(0) - 30
'''''
'''''            If IsNull(rst.FIELDS("OASID")) And IsNull(rst.FIELDS("VINID")) Then '/// NO SHOWPLAN OR VIGNETTE \\\
'''''                cmdBlock(i).BackColor = vbRed
'''''                flx1.Row = i: flx1.Col = 0: flx1.CellFontBold = False
'''''            ElseIf Not IsNull(rst.FIELDS("OASID")) And IsNull(rst.FIELDS("VINID")) Then '///SHOWPLAN, BUT NO VIGNETTE \\\
'''''                cmdBlock(i).BackColor = vbYellow
'''''                flx1.Row = i: flx1.Col = 0: flx1.CellFontBold = True
'''''            ElseIf IsNull(rst.FIELDS("OASID")) And Not IsNull(rst.FIELDS("VINID")) Then '/// VIGNETTE, BUT NO SHOWPLAN \\\
'''''                cmdBlock(i).BackColor = vbBlue
'''''                flx1.Row = i: flx1.Col = 0: flx1.CellFontBold = True
'''''            Else '/// BOTH SHOWPLAN AND VIGNETTE WERE FOUND
'''''                cmdBlock(i).BackColor = vbGreen
'''''                flx1.Row = i: flx1.Col = 0: flx1.CellFontBold = True
'''''            End If
'''''
'''''            cmdBlock(i).Visible = True
'''''            cmdBlock(i).ToolTipText = Trim(rst.FIELDS("BEG_DATE")) & " --> " & _
'''''                        Trim(rst.FIELDS("END_DATE"))
'''''            cmdBlock(i).Tag = rst.FIELDS("SHY56SHCD")
'''''            i = i + 1
'''''            rst.MoveNext
'''''        Loop
'''''        rst.Close: Set rst = Nothing
'''''        Pic1.Visible = True: flx1.Visible = True
'''''    Else
'''''        MsgBox "Selected Client has no current Shows.", vbExclamation, "No Shows Found..."
'''''    End If
'''''End Function

'''''Public Sub PopShowInfo(iD As Integer, lSHCD As Long)
'''''    Dim strSelect As String
'''''    Dim rst As ADODB.Recordset
'''''
''''''''    Debug.Print "Top = " & cmdBlock(Index).Top
''''''''
''''''''    lSHCD = CLng(cmdBlock(Index).Tag)
'''''
'''''    strSelect = "SELECT SM.SHY56NAMA, " & _
'''''                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
'''''                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
'''''                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DAY')BEG_DAY, " & _
'''''                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DAY')END_DAY, " & _
'''''                "AD.ALCTY1, AD.ALADDS " & _
'''''                "FROM " & F5601 & " SM, " & F0116 & " AD " & _
'''''                "WHERE SM.SHY56SHCD = " & lSHCD & " " & _
'''''                "AND SM.SHY56SHYR = " & cboSHYR(0).Text & " " & _
'''''                "AND SM.SHY56FCCDT = AD.ALAN8 (+)"
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''
'''''        mnuShowName.Caption = UCase(Trim(rst.FIELDS("SHY56NAMA")))
'''''        If IsNull(rst.FIELDS("ALCTY1")) Then
'''''            mnuShowLoc.Visible = False
''''''''                mnuShowLoc.Caption = "Show Location:  N/A"
'''''        Else
'''''            mnuShowLoc.Visible = True
'''''            mnuShowLoc.Caption = "Show Location:  " & UCase(Trim(rst.FIELDS("ALCTY1"))) & ", " & _
'''''                        UCase(Trim(rst.FIELDS("ALADDS")))
'''''        End If
'''''        mnuShowOpen.Caption = "Show Open:  " & UCase(Trim(rst.FIELDS("BEG_DAY"))) & _
'''''                    "  " & UCase(Trim(rst.FIELDS("BEG_DATE")))
'''''        mnuShowClose.Caption = "Show Close:  " & UCase(Trim(rst.FIELDS("END_DAY"))) & _
'''''                    "  " & UCase(Trim(rst.FIELDS("END_DATE")))
'''''        rst.Close: Set rst = Nothing
'''''
'''''        Select Case cmdBlock(iD).BackColor
'''''            Case vbRed
'''''                mnuShowPlan.Visible = False
'''''                mnuVignette.Visible = False
'''''            Case vbYellow
'''''                mnuShowPlan.Visible = True
'''''                mnuVignette.Visible = False
'''''            Case vbBlue
'''''                mnuShowPlan.Visible = False
'''''                mnuVignette.Visible = True
'''''            Case vbGreen
'''''                mnuShowPlan.Visible = True
'''''                mnuVignette.Visible = True
'''''        End Select
'''''
'''''        PopupMenu mnuShowData
'''''    Else
'''''        rst.Close: Set rst = Nothing
'''''    End If
'''''End Sub

Public Sub PopShowInfo2(lSHCD As Long)
    Dim strSelect As String, sMess As String, sHDR As String
    Dim rst As ADODB.Recordset
    Dim bFound1 As Boolean
    
    
'''    Debug.Print "Top = " & cmdBlock(Index).Top
'''
'''    lSHCD = CLng(cmdBlock(Index).Tag)
    sMess = "Show Dates:" & vbNewLine: bFound1 = False: sHDR = "Show Dates"
'''    strSelect = "SELECT SM.SHY56NAMA, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DAY')BEG_DAY, SM.SHY56BEGTT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DAY')END_DAY, SM.SHY56ENDTT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'DD-MON-YYYY')SUBEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'DD-MON-YYYY')SUEND_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'DAY')SUBEG_DAY, SM.SHY56SBEDT, SM.SHY56SBETT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'DAY')SUEND_DAY, SM.SHY56SENDT, SM.SHY56SENTT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'DD-MON-YYYY')TDBEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'DD-MON-YYYY')TDEND_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'DAY')TDBEG_DAY, SM.SHY56TBEDT, SM.SHY56TBETT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'DAY')TDEND_DAY, SM.SHY56TEDDT, SM.SHY56TENTT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'DD-MON-YYYY')PVBEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'DD-MON-YYYY')PVEND_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'DAY')PVBEG_DAY, SM.SHY56VBEDT, SM.SHY56VBETT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'DAY')PVEND_DAY, SM.SHY56VENDT, SM.SHY56VENTT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'DD-MON-YYYY')PRBEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'DD-MON-YYYY')PREND_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'DAY')PRBEG_DAY, SM.SHY56PBEDT, SM.SHY56PBETT, " & _
'''                "IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'DAY')PREND_DAY, SM.SHY56PENDT, SM.SHY56PENTT " & _
'''                "FROM " & F5601 & " SM, " & F0116 & " AD " & _
'''                "WHERE SM.SHY56SHCD = " & Mid(cSHNode, 3) & " " & _
'''                "AND SM.SHY56SHYR = " & cboSHYR(1).Text & " AND SM.SHY56FCCDT = AD.ALAN8 (+)"
        strSelect = "SELECT SM.SHY56NAMA, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MON DD, YYYY')BEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MON DD, YYYY')END_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DAY')BEG_DAY, SM.SHY56BEGTT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DAY')END_DAY, SM.SHY56ENDTT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MON DD, YYYY')SUBEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MON DD, YYYY')SUEND_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'DAY')SUBEG_DAY, SM.SHY56SBEDT, SM.SHY56SBETT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'DAY')SUEND_DAY, SM.SHY56SENDT, SM.SHY56SENTT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MON DD, YYYY')TDBEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MON DD, YYYY')TDEND_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'DAY')TDBEG_DAY, SM.SHY56TBEDT, SM.SHY56TBETT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'DAY')TDEND_DAY, SM.SHY56TEDDT, SM.SHY56TENTT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MON DD, YYYY')PVBEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MON DD, YYYY')PVEND_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'DAY')PVBEG_DAY, SM.SHY56VBEDT, SM.SHY56VBETT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'DAY')PVEND_DAY, SM.SHY56VENDT, SM.SHY56VENTT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'MON DD, YYYY')PRBEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'MON DD, YYYY')PREND_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'DAY')PRBEG_DAY, SM.SHY56PBEDT, SM.SHY56PBETT, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'DAY')PREND_DAY, SM.SHY56PENDT, SM.SHY56PENTT " & _
                "FROM " & F5601 & " SM, " & F0116 & " AD " & _
                "WHERE SM.SHY56SHCD = " & Mid(cSHNode, 3) & " " & _
                "AND SM.SHY56SHYR = " & cboSHYR(1).Text & " AND SM.SHY56FCCDT = AD.ALAN8 (+)"

    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
'''        If Not IsNull(rst.FIELDS("ALCTY1")) Then
'''            sMess = sMess & "Show Location:  " & UCase(Trim(rst.FIELDS("ALCTY1"))) & ", " & _
'''                        UCase(Trim(rst.FIELDS("ALADDS"))) & vbNewLine & vbNewLine
'''        End If
        
        If rst.Fields("SHY56PBEDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Press Begin:  " & Trim(rst.Fields("PRBEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("PRBEG_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56PBETT")) & ")" & vbNewLine
        End If
        If rst.Fields("SHY56PENDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Press End:  " & Trim(rst.Fields("PREND_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("PREND_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56PENTT")) & ")" & vbNewLine
        End If
        If bFound1 Then sMess = sMess & vbNewLine: bFound1 = False
        
        If rst.Fields("SHY56VBEDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Preview Begin:  " & Trim(rst.Fields("PVBEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("PVBEG_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56VBETT")) & ")" & vbNewLine
        End If
        If rst.Fields("SHY56VENDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Preview End:  " & Trim(rst.Fields("PVEND_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("PVEND_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56VENTT")) & ")" & vbNewLine
        End If
        If bFound1 Then sMess = sMess & vbNewLine: bFound1 = False
        
        If rst.Fields("SHY56BEGTT") > 0 Then
            sMess = sMess & Space(3) & "Show Open:  " & Trim(rst.Fields("BEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("BEG_DATE"))) & " @ " & _
                            ConvertTime(rst.Fields("SHY56BEGTT")) & vbNewLine
        Else
            sMess = sMess & Space(3) & "Show Open:  " & Trim(rst.Fields("BEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("BEG_DATE"))) & vbNewLine
        End If
        If rst.Fields("SHY56ENDTT") > 0 Then
            sMess = sMess & Space(3) & "Show Close:  " & Trim(rst.Fields("END_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("END_DATE"))) & " @ " & _
                            ConvertTime(rst.Fields("SHY56ENDTT")) & vbNewLine & vbNewLine
        Else
            sMess = sMess & Space(3) & "Show Close:  " & Trim(rst.Fields("END_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("END_DATE"))) & vbNewLine & vbNewLine
        End If
        
        sShowDates = Left(UCase(Trim(rst.Fields("BEG_DATE"))), Len(Trim(rst.Fields("BEG_DATE"))) - 6) & _
                    " - " & UCase(Trim(rst.Fields("END_DATE")))
        
        If rst.Fields("SHY56SBEDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Install Begin:  " & Trim(rst.Fields("SUBEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("SUBEG_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56SBETT")) & ")" & _
                        vbNewLine
        End If
        If rst.Fields("SHY56SENDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Install End:  " & Trim(rst.Fields("SUEND_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("SUEND_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56SENTT")) & ")" & _
                        vbNewLine
        End If
        If bFound1 Then sMess = sMess & vbNewLine: bFound1 = False
        
        If rst.Fields("SHY56TBEDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Dismantle Begin:  " & Trim(rst.Fields("TDBEG_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("TDBEG_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56TBETT")) & ")" & vbNewLine
        End If
        If rst.Fields("SHY56TEDDT") <> 0 Then
            bFound1 = True
            sMess = sMess & Space(3) & "Dismantle End:  " & Trim(rst.Fields("TDEND_DAY")) & _
                        "  " & UCase(Trim(rst.Fields("TDEND_DATE"))) & " (" & _
                        ConvertTime(rst.Fields("SHY56TENTT")) & ")" & vbNewLine
        End If
        If bFound1 Then sMess = sMess & vbNewLine: bFound1 = False
        rst.Close: Set rst = Nothing
    Else
        
        rst.Close: Set rst = Nothing
    End If
    
    If bClientAll_Enabled Then
        strSelect = "SELECT AB.ABALPH " & _
                    "FROM " & F5611 & " CS, " & F0101 & " AB " & _
                    "WHERE CS.CSY56SHYR = " & cboSHYR(1).Text & " " & _
                    "AND CS.CSY56SHCD = " & Mid(cSHNode, 3) & " " & _
                    "AND CS.CSY56CUNO = AB.ABAN8 " & _
                    "AND AB.ABAT1 = 'C' " & _
                    "ORDER BY UPPER(AB.ABALPH)"
    Else
        strSelect = "SELECT AB.ABALPH " & _
                    "FROM " & F5611 & " CS, " & F0101 & " AB " & _
                    "WHERE CS.CSY56SHYR = " & cboSHYR(1).Text & " " & _
                    "AND CS.CSY56SHCD = " & Mid(cSHNode, 3) & " " & _
                    "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                    "AND CS.CSY56CUNO = AB.ABAN8 " & _
                    "AND AB.ABAT1 = 'C' " & _
                    "ORDER BY UPPER(AB.ABALPH)"
    End If
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sMess = sMess & vbNewLine & "Attending Clients:" & vbNewLine
        sHDR = sHDR & " && Attending Clients"
        Do While Not rst.EOF
            sMess = sMess & Space(3) & Trim(rst.Fields("ABALPH")) & vbNewLine
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing

    rtbMess.TextRTF = sMess
    lblMess = sHDR & ":"
End Sub

Public Sub PopTree(sCUNO As String)
    Dim strSelect As String, sDesc As String, sHDR As String, sOrderBy As String
    Dim rst As ADODB.Recordset, rstN As ADODB.Recordset
    Dim nodX As Node
    Dim tSHCD As Long, tFCCD As Long, tHALL As Long, tSMGR As Long, tGCON As Long
    Dim SHNode As String, FCNode As String, HANode As String, CHNode As String, _
                HRNode As String, EANode As String, EDNode As String, SMNode As String, _
                GCNode As String, ACNode As String, ENNode As String, LONode As String, _
                ESNode As String, ECNode As String
    Dim iIcon As Integer, i As Integer, iShowCnt As Integer
    
    pSHNode = ""
    tvw1.Visible = False
    lblMess = "": rtbMess.Text = "": lblShow = ""
    tvw1.Nodes.Clear
    
    If optSort(0).Value = True Then sOrderBy = "ORDER BY UPPER(SM.SHY56NAMA), SM.SHY56SHCD, HM.HALLDESC" _
                Else sOrderBy = "ORDER BY SM.SHY56BEGDT, UPPER(SM.SHY56NAMA), SM.SHY56SHCD, HM.HALLDESC"
                
    tvw1.ImageList = ImageList1
    tSHCD = 0: tFCCD = 0: tSMGR = 0: tGCON = 0: iShowCnt = 0
    
    If sCUNO = "" Then
        strSelect = "SELECT SM.SHY56SHYR, SM.SHY56SHCD, SM.SHY56NAMA, SM.SHY56BEGDT, " & _
                    "SM.SHY56FCCDT, AB.ABALPH AS FACIL, SM.SHY56SMGRT, MG.ABALPH AS SMGR, " & _
                    "SM.SHY56GCONT, GC.ABALPH AS GCON, HM.HALLID, HM.HALLDESC " & _
                    "FROM " & F5601 & " SM, " & F0101 & " AB, " & SRAHallMas & " HM, " & _
                    "" & F0101 & " MG, " & F0101 & " GC " & _
                    "Where SM.SHY56SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                    "AND SM.SHY56FCCDT = AB.ABAN8 (+) " & _
                    "AND SM.SHY56FCCDT = HM.AN8_FCCD (+) " & _
                    "AND SM.SHY56SMGRT = MG.ABAN8 (+) " & _
                    "AND SM.SHY56GCONT = GC.ABAN8 (+) " & _
                    sOrderBy
    Else
        strSelect = "SELECT DISTINCT SM.SHY56SHYR, SM.SHY56SHCD, SM.SHY56NAMA, SM.SHY56BEGDT, " & _
                    "SM.SHY56FCCDT, AB.ABALPH AS FACIL, SM.SHY56SMGRT, MG.ABALPH AS SMGR, " & _
                    "SM.SHY56GCONT, GC.ABALPH AS GCON, HM.HALLID, HM.HALLDESC " & _
                    "FROM " & F5601 & " SM, " & F0101 & " AB, " & F5611 & " CS, " & _
                    "" & SRAHallMas & " HM, " & F0101 & " MG, " & F0101 & " GC " & _
                    "Where SM.SHY56SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                    "AND SM.SHY56FCCDT = AB.ABAN8 (+) " & _
                    "AND SM.SHY56FCCDT = HM.AN8_FCCD (+) " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND CS.CSY56CUNO IN (" & sCUNO & ") " & _
                    "AND SM.SHY56SMGRT = MG.ABAN8 (+) " & _
                    "AND SM.SHY56GCONT = GC.ABAN8 (+) " & _
                    sOrderBy
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If rst.Fields("SHY56SHCD") <> tSHCD Then
            SHNode = "SH" & rst.Fields("SHY56SHCD")
            sDesc = Trim(rst.Fields("SHY56NAMA"))
            Set nodX = tvw1.Nodes.Add(, , SHNode, sDesc, 1)
            tSHCD = rst.Fields("SHY56SHCD")
            tFCCD = 0: tSMGR = 0: tGCON = 0
            iShowCnt = iShowCnt + 1
        End If
        If rst.Fields("SHY56FCCDT") <> tFCCD Then
            strSelect = "SELECT ALCTY1, ALADDS " & _
                        "FROM " & F0116 & " " & _
                        "WHERE ALAN8 = " & rst.Fields("SHY56FCCDT") & " " & _
                        "AND ALEFTB IN " & _
                        "(SELECT MAX(ALEFTB) " & _
                        "FROM " & F0116 & " " & _
                        "WHERE ALAN8 = " & rst.Fields("SHY56FCCDT") & ")"
            Set rstN = Conn.Execute(strSelect)
            If Not rstN.EOF Then
                LONode = "LO" & rst.Fields("SHY56SHCD")
                sDesc = "Location:  " & Trim(rstN.Fields("ALCTY1")) & "  " & Trim(rstN.Fields("ALADDS"))
                Set nodX = tvw1.Nodes.Add(SHNode, tvwChild, LONode, sDesc, 2)
                FCNode = "FC" & rst.Fields("SHY56SHCD") & "-" & Trim(rst.Fields("SHY56FCCDT"))
                sDesc = "Facility:  " & Trim(rst.Fields("FACIL"))
                Set nodX = tvw1.Nodes.Add(SHNode, tvwChild, FCNode, sDesc, 3)
                nodX.Parent.Image = 3
                tFCCD = rst.Fields("SHY56FCCDT")
                rstN.Close: Set rstN = Nothing
                strSelect = "SELECT HM.HALLID, HM.HALLDESC, HM.HALLNOTE, " & _
                            "HM.CLGHGT, HM.CLGUNIT, HM.CLGNOTE, " & _
                            "SAW.RESID, SAW.HGTRES, SAW.RESUNIT, SAW.RESNOTE " & _
                            "FROM " & SRAHallMas & " HM, " & _
                            "(SELECT HALLID, RESID, HGTRES, RESUNIT, RESNOTE " & _
                            "FROM " & SRAHallRes & " " & _
                            "WHERE AN8_SHCD = " & tSHCD & ") SAW " & _
                            "WHERE HM.AN8_FCCD = " & tFCCD & " " & _
                            "AND HM.HALLID = SAW.HALLID (+) " & _
                            "ORDER BY HM.HALLDESC"
                Set rstN = Conn.Execute(strSelect)
                Do While Not rstN.EOF
                    HANode = "HA" & rst.Fields("SHY56SHCD") & "-" & rstN.Fields("HALLID")
                    sDesc = "Hall:  " & Trim(rstN.Fields("HALLDESC"))
                    If IsNull(rstN.Fields("HALLNOTE")) Or Trim(rstN.Fields("HALLNOTE")) = "" Then iIcon = 2 Else iIcon = 3
                    Set nodX = tvw1.Nodes.Add(FCNode, tvwChild, HANode, sDesc, iIcon)
                    
                    If Not IsNull(rstN.Fields("CLGHGT")) And rstN.Fields("CLGHGT") > 0 Then
                        CHNode = "CH" & rst.Fields("SHY56SHCD") & "-" & rstN.Fields("HALLID")
                        sDesc = "Ceiling Height:  " & CalcDim(rstN.Fields("CLGHGT"), rstN.Fields("CLGUNIT"))
                        If IsNull(rstN.Fields("CLGNOTE")) Or Trim(rstN.Fields("CLGNOTE")) = "" Then iIcon = 2 Else iIcon = 3
                        Set nodX = tvw1.Nodes.Add(HANode, tvwChild, CHNode, sDesc, iIcon)
                    End If
                    
                    If Not IsNull(rstN.Fields("HGTRES")) And rstN.Fields("HGTRES") > 0 Then
                        HRNode = "HR" & rst.Fields("SHY56SHCD") & "-" & rstN.Fields("RESID")
                        sDesc = "Height Restriction:  " & CalcDim(rstN.Fields("HGTRES"), rstN.Fields("RESUNIT"))
                        If IsNull(rstN.Fields("RESNOTE")) Or Trim(rstN.Fields("RESNOTE")) = "" Then iIcon = 2 Else iIcon = 3
                        Set nodX = tvw1.Nodes.Add(HANode, tvwChild, HRNode, sDesc, iIcon)
                    End If
                    
                    rstN.MoveNext
                Loop
                rstN.Close: Set rstN = Nothing
            Else
                rstN.Close: Set rstN = Nothing
            End If
        End If
        
        If rst.Fields("SHY56SMGRT") <> tSMGR Then
            SMNode = "SM" & rst.Fields("SHY56SHCD") & "-" & rst.Fields("SHY56SMGRT")
            sDesc = "Show Manager:  " & Trim(rst.Fields("SMGR"))
            Set nodX = tvw1.Nodes.Add(SHNode, tvwChild, SMNode, sDesc, 3)
            tSMGR = rst.Fields("SHY56SMGRT")
            nodX.Parent.Image = 3
        End If
            
        If rst.Fields("SHY56GCONT") <> tGCON Then
            GCNode = "GC" & rst.Fields("SHY56SHCD") & "-" & rst.Fields("SHY56GCONT")
            sDesc = "Show Contractor:  " & Trim(rst.Fields("GCON"))
            Set nodX = tvw1.Nodes.Add(SHNode, tvwChild, GCNode, sDesc, 3)
            tGCON = rst.Fields("SHY56GCONT")
            nodX.Parent.Image = 3
        End If
            
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
            
    '///// THE EASEMENTS \\\\\
    On Error Resume Next
    tSHCD = 0: i = 1
    strSelect = "SELECT * FROM " & SRAEase & " " & _
                "ORDER BY AN8_SHCD, HALLID, EASENAME"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        tSHCD = rst.Fields("AN8_SHCD")
        If tHALL <> rst.Fields("HALLID") Then
            tHALL = rst.Fields("HALLID")
            EANode = "EA" & i: i = i + 1
            HANode = "HA" & tSHCD & "-" & tHALL
            sDesc = "Easements"
            Set nodX = tvw1.Nodes.Add(HANode, tvwChild, EANode, sDesc, 2)
        End If
        EDNode = "ED" & tSHCD & "-" & rst.Fields("EASEID")
        sDesc = CalcDim(rst.Fields("EASEVAL"), rst.Fields("EASEUNIT")) & _
                    "  [" & UCase(Trim(rst.Fields("EASENAME"))) & "]"
        If IsNull(rst.Fields("EASEDESC")) Or Trim(rst.Fields("EASEDESC")) = "" Then iIcon = 2 Else iIcon = 3
        Set nodX = tvw1.Nodes.Add(EANode, tvwChild, EDNode, sDesc, iIcon)
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    '///// ATTENDING CLIENTS \\\\\
    sDesc = "Attending Clients in Hall..."
    strSelect = "SELECT DISTINCT AN8_SHCD, HALLID " & _
                "FROM " & SRACliHall & " " & _
                "WHERE SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                "ORDER BY AN8_SHCD, HALLID"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        HANode = "HA" & rst.Fields("AN8_SHCD") & "-" & rst.Fields("HALLID")
        ACNode = "AC" & rst.Fields("AN8_SHCD") & "-" & rst.Fields("HALLID")
        Set nodX = tvw1.Nodes.Add(HANode, tvwChild, ACNode, sDesc, 4)
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    '///// ENGINEERING CODE REQUIREMENTS \\\\\
    sHDR = "Engineering Code Requirements..."
    tSHCD = 0
    '''SELECT ALL ACTIVE CODE REQS, ORDER BY AN8_SHCD, CODENAME
    strSelect = "SELECT ENGCODEID, ENGCODENAME, AN8_SHCD " & _
                "FROM " & SRAEngCodeReq & " " & _
                "ORDER BY AN8_SHCD, ENGCODENAME"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If rst.Fields("AN8_SHCD") <> tSHCD Then
            tSHCD = rst.Fields("AN8_SHCD")
            SHNode = "SH" & tSHCD
            ESNode = "ES" & tSHCD
            Set nodX = tvw1.Nodes.Add(SHNode, tvwChild, ESNode, sHDR, 6)
        End If
        ECNode = "EC" & rst.Fields("ENGCODEID")
        sDesc = Trim(rst.Fields("ENGCODENAME"))
        Set nodX = tvw1.Nodes.Add(ESNode, tvwChild, ECNode, sDesc, 3)
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    '///// RESET ICON WHERE SHOWPLAN IS AVAILABLE \\\\\
    On Error Resume Next
    strSelect = "SELECT DISTINCT DS.AN8_SHCD " & _
                "FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
                "WHERE DS.SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                "AND DS.DWGID = DM.DWGID " & _
                "AND DM.DWGTYPE = 3"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        tvw1.Nodes("SH" & rst.Fields("AN8_SHCD")).Image = 5
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
     '///// RESET ICON WHERE COMPOSITE IS AVAILABLE \\\\\
    On Error Resume Next
    strSelect = "SELECT DISTINCT DS.AN8_SHCD " & _
                "FROM " & DWGShow & " DS, " & DWGMas & " DM " & _
                "WHERE DS.SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                "AND DS.DWGID = DM.DWGID " & _
                "AND DM.DWGTYPE = 4"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        tvw1.Nodes("SH" & rst.Fields("AN8_SHCD")).Image = 7
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    Select Case iShowCnt
        Case Is = 1: lblCount = iShowCnt & " Show"
        Case Else: lblCount = iShowCnt & " Shows"
    End Select
    tvw1.Visible = True
End Sub

Public Function CalcDim(Num As Single, iUnit As Integer) As String
    Dim Feet As Integer, Inch As Integer, Numer As Integer
    Dim Frac As Currency
    Dim strFrac As String
    
    Select Case iUnit
        Case 1, 2
            If iUnit = 2 Then Num = Num * 12
            Feet = Int(Num / 12)
            Inch = Int(Num - (Feet * 12))
            Frac = CCur((((Num / 12) - Feet) * 12) - Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1: strFrac = " 1/8"""
                    Case 2: strFrac = " 1/4"""
                    Case 3: strFrac = " 3/8"""
                    Case 4: strFrac = " 1/2"""
                    Case 5: strFrac = " 5/8"""
                    Case 6: strFrac = " 3/4"""
                    Case 7: strFrac = " 7/8"""
                    Case 8
                        Inch = Inch + 1
                        If Inch = 12 Then
                            Feet = Feet + 1
                            Inch = 0
                        End If
                        strFrac = Chr(34)
                End Select
            Else
                strFrac = Chr(34)
            End If
            CalcDim = Feet & "'-" & Inch & strFrac
        Case 5
            CalcDim = Format(Num, "#,##0.0") & " cm"
        Case 6
            CalcDim = Format(Num, "#,##0.0000") & " M"
    End Select
End Function

Function DeGlitch(sName As String) As String
    Dim iStr As Integer
    iStr = 1
    Do While InStr(iStr, sName, "'") <> 0
        sName = Left(sName, InStr(iStr, sName, "'")) & "'" & Mid(sName, InStr(iStr, sName, "'") + 1)
        iStr = InStr(iStr, sName, "'") + 2
    Loop
    DeGlitch = sName
End Function

Function DblAmp(sName As String) As String
    Dim iStr As Integer
    iStr = 1
    Do While InStr(iStr, sName, "&") <> 0
        sName = Left(sName, InStr(iStr, sName, "&")) & "&" & Mid(sName, InStr(iStr, sName, "&") + 1)
        iStr = InStr(iStr, sName, "&") + 2
    Loop
    DblAmp = sName
End Function

Public Sub CheckForShowplan(tSHYR As Integer, tSHCD As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bSFound As Boolean, bCFound As Boolean
    
    bSFound = False: bCFound = False
    '///// FIRST LOOK FOR SHOWPLAN \\\\\
    strSelect = "SELECT DD.DWFPATH " & _
                "FROM " & DWGShow & " DS, " & DWGMas & " DM, " & DWGDwf & " DD " & _
                "WHERE DS.SHYR = " & tSHYR & " " & _
                "AND DS.AN8_SHCD = " & tSHCD & " " & _
                "AND DS.AN8_CUNO IS NULL " & _
                "AND DS.DWGID = DM.DWGID " & _
                "AND DM.DWGTYPE = 3 " & _
                "AND DS.DWGID = DD.DWGID " & _
                "AND DD.DWFTYPE = 30"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        cmdViewShowplan.Tag = Trim(rst.Fields("DWFPATH"))
        cmdViewShowplan.Enabled = True
        cmdViewShowplan.Visible = True
        bSFound = True
    Else
        cmdViewShowplan.Tag = ""
        cmdViewShowplan.Enabled = False
        cmdViewShowplan.Visible = True
        bSFound = False
    End If
    rst.Close: Set rst = Nothing
    
    If bPerm(5) Then
        '///// NEXT LOOK FOR COMPOSITE PLAN \\\\\
        strSelect = "SELECT DD.DWFPATH " & _
                    "FROM " & DWGShow & " DS, " & DWGMas & " DM, " & DWGDwf & " DD " & _
                    "WHERE DS.SHYR = " & tSHYR & " " & _
                    "AND DS.AN8_SHCD = " & tSHCD & " " & _
                    "AND DS.AN8_CUNO IS NULL " & _
                    "AND DS.DWGID = DM.DWGID " & _
                    "AND DM.DWGTYPE = 4 " & _
                    "AND DS.DWGID = DD.DWGID " & _
                    "AND DD.DWFTYPE = 31"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            cmdViewComposite.Tag = Trim(rst.Fields("DWFPATH"))
            cmdViewComposite.Enabled = True
            cmdViewComposite.Visible = True
            bCFound = True
        Else
            cmdViewComposite.Tag = ""
            cmdViewComposite.Enabled = False
            cmdViewComposite.Visible = True
            bCFound = False
        End If
        rst.Close: Set rst = Nothing
    End If
    
    If bSFound And bCFound Then
        cmdViewShowplan.Width = 1005
        cmdViewShowplan.Caption = "Showplan"
    Else
        cmdViewShowplan.Width = 2055
        cmdViewShowplan.Caption = "View Overall Showplan"
    End If
        
'''    CheckForShowplan = bFound
End Sub

Public Sub ClearChecks()
    mnuVPan.Checked = False
    mnuVZoom.Checked = False
    mnuVZoomW.Checked = False
End Sub

Public Function InitialView()
    volShowplan.GetCurrentView dLeft, dRight, dBottom, dTop
End Function

Public Sub CodeSet(cbo1 As ComboBox, sText As String)
    Const CB_FINDSTRING = &H14C
    Dim x As Long
    Dim Pos As Integer
    x = SendMessage(cbo1.hwnd, CB_FINDSTRING, 0, ByVal sText)
    If x = -1 Then
        Pos = cbo1.SelStart
''''        If Pos > 0 Then Pos = cbo1.SelStart - 1 Else Pos = 0
''''        cbo1.Text = Left(cbo1.Text, Pos)
    Else
        Pos = cbo1.SelStart
        cbo1.Text = cbo1.List(x)
        cbo1.SelStart = Pos
        cbo1.SelLength = Len(cbo1.Text) - Pos
    End If
End Sub

Public Function CheckForConfirmation(sType As String) As Boolean
    Dim strSelect As String, sMess As String, sConfUser As String, strUpdate As String, sConfShort As String
    Dim rst As ADODB.Recordset, rstC As ADODB.Recordset
    Dim Resp As VbMsgBoxResult
    Dim E_Alert As Boolean
    Dim tSHYR As Integer
    
    strSelect = "SELECT CC.SHYR, CC.CONFIRMUSER, U.NAME_FIRST, U.NAME_LAST, " & _
                "TO_CHAR(CC.CONFIRMDTTM, 'MON DD, YYYY') AS CONFDATE " & _
                "FROM " & SRAEngCodeConf & " CC, " & IGLUser & " U " & _
                "WHERE CC.AN8_SHCD = " & Mid(cSHNode, 3) & " " & _
                "AND CC.CONFIRMSTATUS > 0 " & _
                "AND CC.CONFIRMUSER = U.NAME_LOGON " & _
                "ORDER BY CC.SHYR DESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If rst.Fields("SHYR") >= CInt(Format(Now, "YYYY")) Then
            '///// CONFIRMATION EXIST, CHECK IF FUTURE SHOW \\\\\
            sConfUser = Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST"))
            sConfShort = Trim(rst.Fields("CONFIRMUSER"))
            strSelect = "SELECT IGL_JDEDATE_TOCHAR(SHY56BEGDT, 'MON DD, YYYY')BEG_DATE " & _
                        "FROM " & F5601 & " " & _
                        "WHERE SHY56SHYR = " & rst.Fields("SHYR") & " " & _
                        "AND SHY56SHCD = " & Mid(cSHNode, 3) & " " & _
                        "AND SHY56BEGDT > " & IGLToJDEDate(Now)
            Set rstC = Conn.Execute(strSelect)
            If Not rstC.EOF Then
                If UCase(LogName) <> UCase(sConfUser) Then
                    sMess = "Show Code Requirements have been confirmed for Show Year " & _
                                rst.Fields("SHYR") & " by " & sConfUser & " on " & Trim(rst.Fields("CONFDATE")) & ".  " & _
                                "Show Open date for the confirmed Show is " & Trim(rstC.Fields("BEG_DATE")) & _
                                "." & vbNewLine & vbNewLine & _
                                "Do you want to continue " & sType & " this Code Requirement?  " & _
                                "Selecting 'YES', will inactivate the confirmation and prompt an " & _
                                "automated email to " & sConfUser & ", alerting him of the change."
                Else
                    sMess = "You confirmed Show Code Requirements for Show Year " & _
                                rst.Fields("SHYR") & " on " & Trim(rst.Fields("CONFDATE")) & ".  " & _
                                "Show Open date for the confirmed Show is " & Trim(rstC.Fields("BEG_DATE")) & _
                                "." & vbNewLine & vbNewLine & _
                                "Do you want to continue " & sType & " this Code Requirement?  " & _
                                "Selecting 'YES', will inactivate the confirmation."
                End If
                rstC.Close: Set rstC = Nothing
                Resp = MsgBox(sMess, vbExclamation + vbYesNo, "Existing Confirmation...")
                If Resp = vbYes Then
                    strUpdate = "UPDATE " & SRAEngCodeConf & " SET " & _
                                "CONFIRMSTATUS = 0, " & _
                                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                                "WHERE SHYR = " & rst.Fields("SHYR") & " " & _
                                "AND AN8_SHCD = " & Mid(cSHNode, 3)
                    Conn.Execute (strUpdate)
                    If UCase(LogName) <> UCase(sConfUser) Then
                        E_Alert = True
                        tSHYR = rst.Fields("SHYR")
                    Else
                        E_Alert = False
                    End If
                Else
'''                    rst.Close: Set rst = Nothing
                    bAbort = True
                End If
            Else
                rstC.Close: Set rstC = Nothing
            End If
        End If
    End If
    rst.Close: Set rst = Nothing
    If E_Alert Then
        If sType = "adding" Then sType = "added" Else sType = "edited"
        strEmailHdr = "Engineering Show Code Change for " & tvw1.Nodes(cSHNode).Text
        strEmailMsg = "Engineering Code Requirements have been " & sType & " for the following Show.  " & _
                    "Your confirmation has been INACTIVATED.  Please, review the " & sType & " regulations " & _
                    "and re-confirm their validity for Show Year " & tSHYR & "." & vbNewLine & vbNewLine & _
                    vbTab & tSHYR & " - " & UCase(tvw1.Nodes(cSHNode).Text) & vbNewLine & vbNewLine & _
                    "The Code Requirement was " & sType & " by " & LogName & " on " & Format(Now, "DD-MMM-YYYY") & _
                    " at " & Format(Now, "H:MM AMPM") & "."
        strEmailTo = LCase(Trim(sConfShort)) & "@gpjco.com"
    End If
    CheckForConfirmation = E_Alert
End Function

Public Function ConvertTime(lTime As Long) As String
    Dim lSub As Long
    
    Select Case lTime
        Case 0
            ConvertTime = " "
        Case 120000
            ConvertTime = "12:00 N"
        Case Is > 120000
            lTime = lTime - 120000
            If Len(CStr(lTime)) = 5 Then
                ConvertTime = Mid(lTime, 1, 1) & ":" & Mid(lTime, 2, 2) & " PM"
            ElseIf Len(CStr(lTime)) = 6 Then
                ConvertTime = Mid(lTime, 1, 2) & ":" & Mid(lTime, 3, 2) & " PM"
            End If
        Case Is < 10000
            If Len(CStr(lTime)) = 4 Then
                ConvertTime = "12:" & Mid(lTime, 1, 2) & " AM"
            ElseIf Len(CStr(lTime)) = 3 Then
                ConvertTime = "12:00 AM"
            Else
                ConvertTime = "12:0" & Left(lTime, 1) & " AM"
            End If
        Case Else
            If Len(CStr(lTime)) = 5 Then
                ConvertTime = Mid(lTime, 1, 1) & ":" & Mid(lTime, 2, 2) & " AM"
            ElseIf Len(CStr(lTime)) = 6 Then
                ConvertTime = Mid(lTime, 1, 2) & ":" & Mid(lTime, 3, 2) & " AM"
            End If
    End Select
End Function

Public Sub ConfirmAlert(sTo As String, sHDR As String, sMsg As String)
    '///// EXECUTE E-MAIL \\\\\
'''''    Dim myNotes As New Domino.NotesSession
'''''    Dim myDB As New Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    
''''''    myNotes.Initialize
'''''    On Error Resume Next
'''''    If sNOTESID = "GANNOTAT" Then
'''''        myNotes.Initialize (sNOTESPASSWORD)
'''''    Else
'''''        If sNOTESPASSWORD = "" Then
'''''            ''GET PASSWORD''
'''''TryPWAgain:
'''''            frmGetPassword.Show 1, Me
'''''            Select Case sNOTESPASSWORD
'''''                Case "_CANCEL"
'''''                    sNOTESPASSWORD = ""
'''''                    MsgBox "No email will be sent", vbExclamation, "User Canceled..."
'''''                    Set myNotes = Nothing
'''''                    Set myDB = Nothing
'''''                Case Else
'''''                    Err.Clear
'''''                    myNotes.Initialize (sNOTESPASSWORD)
'''''                    If Err Then
'''''                        Err.Clear
'''''                        GoTo TryPWAgain
'''''                    End If
'''''            End Select
'''''        Else
'''''            myNotes.Initialize (sNOTESPASSWORD)
'''''        End If
'''''    End If
    
    
    Dim MailMan As New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
    
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
    
    Dim Email As New ChilkatEmail2
    Email.FromAddress = LogAddress
    Email.fromName = LogName
    
    Email.subject = sHDR
    Email.Body = sMsg
    
    
    
    
    
'    If Not bCitrix Then
'        ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
'        Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
'        Dim myDB As Object '' LOTUS.NotesDatabase
'
'
'        On Error Resume Next
'        Set myNotes = GetObject(, "Notes.NotesSession")
'
'        If Err Then
'            Err.Clear
'            Set myNotes = CreateObject("Notes.NotesSession")
'            If Err Then
'                MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
'                GoTo GetOut
'            End If
'        End If
'        On Error GoTo 0
'        Set myDB = myNotes.GetDatabase("", "")
'        myDB.OPENMAIL
'        Set myDoc = myDB.CreateDocument
'
'    Else
'        ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
'        Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
'        Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
'
'
'        myDom.Initialize (sGAnnoPW)
'        'Set myDomDB = myDom.GetDatabase("detsrv1/det/GPJNotes", "mail\gannotat.nsf")
'        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'        Set myDoc = myDomDB.CreateDocument
'
'        Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'    End If
    
    Email.AddTo sTo, sTo
        
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
    
'    Set myItem = myDoc.AppendItemValue("Subject", sHDR)
'    Set myRichText = myDoc.CreateRichTextItem("Body")
'    myRichText.AppendText sMsg
'    myDoc.AppendItemValue "SENDTO", sTo
''''    myDoc.SaveMessageOnSend = True
'
'    On Error Resume Next
'    Call myDoc.Send(False, sTo)
'    If Err Then
'        MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                    vbExclamation, "Error Encountered"
'        Err = 0
'    End If
    
GetOut:
'    Set myReply = Nothing
'    Set myRichText = Nothing
'    Set myItem = Nothing
'    Set myDoc = Nothing
'
'    If bCitrix Then
'        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'        If Not myDom Is Nothing Then Set myDom = Nothing
'    Else
'        If Not myDB Is Nothing Then Set myDB = Nothing
'        If Not myNotes Is Nothing Then Set myNotes = Nothing
'    End If
        
    
End Sub



Public Sub CheckForPDFs(tPath As String)
    Dim sCheck As String
    
    sCheck = Left(tPath, Len(tPath) - 4) & ".pdf"
    If Dir(sCheck, vbNormal) = "" Then
        mnuEmailPDF.Enabled = False
        mnuDownloadPDF.Enabled = False
    Else
        mnuEmailPDF.Enabled = True
        mnuDownloadPDF.Enabled = True
    End If
End Sub
