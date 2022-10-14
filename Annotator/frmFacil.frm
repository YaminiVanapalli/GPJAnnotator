VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmFacil 
   BackColor       =   &H00000000&
   Caption         =   "Facilities"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer timRightClick 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   6060
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5355
      Left            =   0
      ScaleHeight     =   5355
      ScaleWidth      =   12015
      TabIndex        =   1
      Top             =   600
      Width           =   12015
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2160
         Top             =   3480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":1424
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picFP 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   30
         ScaleHeight     =   3135
         ScaleWidth      =   3615
         TabIndex        =   22
         Top             =   30
         Visible         =   0   'False
         Width           =   3615
         Begin VB.PictureBox picFPClose 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   1980
            ScaleHeight     =   435
            ScaleWidth      =   1635
            TabIndex        =   28
            Top             =   0
            Width           =   1635
            Begin VB.Label lblFPClose 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Close Floorplan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   195
               Left            =   60
               MouseIcon       =   "frmFacil.frx":1F7E
               MousePointer    =   99  'Custom
               TabIndex        =   29
               Top             =   30
               UseMnemonic     =   0   'False
               Width           =   1095
            End
            Begin VB.Image imgFPClose 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1260
               MouseIcon       =   "frmFacil.frx":2288
               MousePointer    =   99  'Custom
               Picture         =   "frmFacil.frx":2592
               Top             =   60
               Width           =   315
            End
         End
         Begin VOLOVIEWXLibCtl.AvViewX volFP 
            Height          =   1635
            Left            =   120
            TabIndex        =   23
            Top             =   660
            Width           =   3000
            _cx             =   5292
            _cy             =   2884
            Appearance      =   0
            BorderStyle     =   0
            BackgroundColor =   "16777215"
            Enabled         =   -1  'True
            UserMode        =   "Pan"
            HighlightLinks  =   0   'False
            src             =   ""
            LayersOn        =   ""
            LayersOff       =   ""
            SrcTemp         =   ""
            SupportPath     =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
            FontPath        =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
            NamedView       =   ""
            GeometryColor   =   "0"
            PrintBackgroundColor=   "16777215"
            PrintGeometryColor=   "0"
            ShadingMode     =   "Gouraud"
            ProjectionMode  =   "Parallel"
            EnableUIMode    =   "DisableRightClickMenu"
            Layout          =   ""
            DisplayMode     =   -1
         End
         Begin VB.Label lblFPS 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1260
            TabIndex        =   27
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   60
         End
         Begin VB.Label lblFPC 
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1260
            TabIndex        =   26
            Top             =   60
            UseMnemonic     =   0   'False
            Width           =   60
         End
         Begin VB.Label lblFPViewer 
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
            MouseIcon       =   "frmFacil.frx":2C28
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   165
            Width           =   540
         End
         Begin VB.Label lblDtl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   2280
            Width           =   540
         End
         Begin VB.Image imgFPViewer 
            Height          =   570
            Left            =   0
            MouseIcon       =   "frmFacil.frx":2F32
            MousePointer    =   99  'Custom
            Picture         =   "frmFacil.frx":323C
            Top             =   0
            Width           =   1080
         End
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Sort by Name"
         Height          =   360
         Index           =   1
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   1500
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Sort by Location"
         Height          =   360
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.ListBox lstFCCD 
         Height          =   1035
         Left            =   780
         TabIndex        =   14
         Top             =   3060
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1560
         Top             =   3480
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
               Picture         =   "frmFacil.frx":5290
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":582A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":5DC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":635E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":68F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":71D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFacil.frx":776C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraCS 
         Caption         =   "Future Shows at this Facility:"
         Height          =   2835
         Left            =   8400
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   3675
         Begin MSComctlLib.TreeView tvwCS 
            Height          =   1275
            Left            =   300
            TabIndex        =   6
            Top             =   540
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2249
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   265
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdGetShows 
            Caption         =   "Get List of Shows..."
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblCSExpand 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expand All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   1
            Left            =   120
            MouseIcon       =   "frmFacil.frx":7D06
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   2520
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblCSExpand 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collapse All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   0
            Left            =   1020
            MouseIcon       =   "frmFacil.frx":8010
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   2520
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.Frame fraInfo 
         Caption         =   "Facility Contact Information:"
         Height          =   2835
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   3675
         Begin SHDocVwCtl.WebBrowser web1 
            Height          =   1815
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   2595
            ExtentX         =   4577
            ExtentY         =   3201
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin VOLOVIEWXLibCtl.AvViewX volP 
         Height          =   1635
         Left            =   4620
         TabIndex        =   2
         Top             =   3300
         Visible         =   0   'False
         Width           =   7455
         _cx             =   13150
         _cy             =   2884
         Appearance      =   0
         BorderStyle     =   0
         BackgroundColor =   "16777215"
         Enabled         =   0   'False
         UserMode        =   "Pan"
         HighlightLinks  =   0   'False
         src             =   ""
         LayersOn        =   ""
         LayersOff       =   ""
         SrcTemp         =   ""
         SupportPath     =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
         FontPath        =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
         NamedView       =   ""
         GeometryColor   =   "8421504"
         PrintBackgroundColor=   "16777215"
         PrintGeometryColor=   "0"
         ShadingMode     =   "Gouraud"
         ProjectionMode  =   "Parallel"
         EnableUIMode    =   "DefaultUI"
         Layout          =   ""
         DisplayMode     =   -1
      End
      Begin MSComctlLib.TreeView tvw1 
         Height          =   4635
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   8176
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmFacil.frx":831A
      End
      Begin VB.Label lblExpand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collapse All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   0
         Left            =   1020
         MouseIcon       =   "frmFacil.frx":8634
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   5040
         Width           =   810
      End
      Begin VB.Label lblExpand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expand All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmFacil.frx":893E
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   5040
         Width           =   750
      End
      Begin VB.Label lblOpenDWF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click to Open Facility Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   4620
         MouseIcon       =   "frmFacil.frx":8C48
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label lblMess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   600
         Width           =   3915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Facilities:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   660
      End
   End
   Begin VOLOVIEWXLibCtl.AvViewX vol1 
      Height          =   1635
      Left            =   1200
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   7455
      _cx             =   13150
      _cy             =   2884
      Appearance      =   0
      BorderStyle     =   0
      BackgroundColor =   "0"
      Enabled         =   -1  'True
      UserMode        =   "ZoomToRect"
      HighlightLinks  =   0   'False
      src             =   ""
      LayersOn        =   ""
      LayersOff       =   ""
      SrcTemp         =   ""
      SupportPath     =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
      FontPath        =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
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
   Begin VB.Label lblReds 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Label lblOthers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Others..."
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
      Left            =   120
      MouseIcon       =   "frmFacil.frx":8F52
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   1335
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Image imgOthers 
      Height          =   570
      Left            =   0
      Picture         =   "frmFacil.frx":925C
      Top             =   1170
      Width           =   1080
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
      MouseIcon       =   "frmFacil.frx":B2B0
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   765
      Width           =   540
   End
   Begin VB.Image imgPhotos 
      Height          =   480
      Left            =   8520
      MouseIcon       =   "frmFacil.frx":B5BA
      MousePointer    =   99  'Custom
      Picture         =   "frmFacil.frx":B8C4
      ToolTipText     =   "Click to view associated Facility photos"
      Top             =   60
      Visible         =   0   'False
      Width           =   480
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
      Left            =   9660
      MouseIcon       =   "frmFacil.frx":C18E
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   180
      Width           =   510
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmFacil.frx":C498
      MousePointer    =   99  'Custom
      Picture         =   "frmFacil.frx":C7A2
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Width           =   720
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Facility from the list below..."
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
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   3735
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
   Begin VB.Image imgViewer 
      Height          =   570
      Left            =   0
      Picture         =   "frmFacil.frx":D2EC
      Top             =   600
      Width           =   1080
   End
   Begin VB.Menu mnuRC 
      Caption         =   "mnuRC"
      Visible         =   0   'False
      Begin VB.Menu mnuWebsite 
         Caption         =   "Go to Show Website..."
      End
   End
   Begin VB.Menu mnuOSPShow 
      Caption         =   "mnuOSPShow"
      Visible         =   0   'False
      Begin VB.Menu mnuOSP 
         Caption         =   "View Photo..."
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
      Begin VB.Menu mnuVDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuMainDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Default Colors"
            Index           =   0
         End
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Black on White"
            Index           =   1
         End
      End
      Begin VB.Menu mnuVPrint 
         Caption         =   "Print..."
         Index           =   0
      End
      Begin VB.Menu mnuVPrint 
         Caption         =   "Print without Camera Icons..."
         Index           =   1
      End
      Begin VB.Menu mnuVDash02 
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
      Begin VB.Menu mnuVDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuVCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuFP 
      Caption         =   "mnuFP"
      Visible         =   0   'False
      Begin VB.Menu mnuFPPan 
         Caption         =   "Pan"
      End
      Begin VB.Menu mnuFPZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuFPZoomW 
         Caption         =   "Zoom Window"
      End
      Begin VB.Menu mnuFPFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuFPDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuFPMainDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuFPDisplay 
            Caption         =   "Default Colors"
            Index           =   0
         End
         Begin VB.Menu mnuFPDisplay 
            Caption         =   "Black on White"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFPPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuFPDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmFacil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iTVWbutton As Integer
Dim iSHYR As Integer
Public lFCCD As Long
Dim bViewSet As Boolean, bFPViewSet As Boolean
Dim dLeft As Double, dRight As Double, dTop As Double, dBottom As Double
Dim dFPLeft As Double, dFPRight As Double, dFPTop As Double, dFPBottom As Double
Dim iRightClick As Integer


'''Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
'''            (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
'''            ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
            
            

Private Sub cmdGetShows_Click()
    Screen.MousePointer = 11
    Call GetShows(iSHYR, lFCCD)
    tvwCS.Visible = True
    lblCSExpand(0).Visible = True: lblCSExpand(1).Visible = True
    Screen.MousePointer = 0
End Sub



'Private Sub Form_Click()
'    vol1.LayersOff = "CAMERA"
'End Sub

Private Sub Form_Load()
    
    iSHYR = CInt(Format(Date, "YYYY"))
    
    tvwCS.Left = cmdGetShows.Left
    tvwCS.Top = cmdGetShows.Top
    
    Call GetFacilities(0) ''(iSHYR)
    
    lblMess.Caption = "Welcome to GPJ Facilities..." & vbNewLine & vbNewLine & _
                    "This interface allows you to view Facility drawings " & _
                    "and the information GPJ has stored in JDEdwards " & _
                    "and other interfacing applications." & _
                    vbNewLine & vbNewLine & _
                    "To access a Facility, expand the tree at left " & _
                    "and click on the Facility name."

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    shpHDR.Width = Me.ScaleWidth
    
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight - shpHDR.Height
    
    tvw1.Width = (picBack.ScaleWidth - 480) / 3
    fraInfo.Left = tvw1.Left + tvw1.Width + 120
    fraInfo.Width = tvw1.Width
    fraCS.Left = fraInfo.Left + fraInfo.Width + 120
    fraCS.Width = tvw1.Width
    
    volP.Left = fraCS.Left
    volP.Width = fraCS.Width
    
    tvw1.Height = picBack.ScaleHeight - tvw1.Top - 120 - 150 ''360
    fraInfo.Height = picBack.ScaleHeight - fraInfo.Top - 120 '' tvw1.Height + (tvw1.Top - fraInfo.Top) ''
    
    volP.Top = picBack.ScaleHeight * (2 / 3)
    volP.Height = picBack.ScaleHeight - volP.Top - 120
    
    lblOpenDWF.Left = volP.Left
    lblOpenDWF.Top = volP.Top - 60 - lblOpenDWF.Height
    
    fraCS.Height = volP.Top - fraInfo.Top - 300
    
    cmdGetShows.Width = fraCS.Width - (cmdGetShows.Left * 2)
    tvwCS.Width = cmdGetShows.Width
    tvwCS.Height = fraCS.Height - tvwCS.Top - tvwCS.Left - 300
    lblCSExpand(0).Top = tvwCS.Top + tvwCS.Height + 60
    lblCSExpand(1).Top = lblCSExpand(0).Top
    
    web1.Width = fraInfo.Width - (web1.Left * 2)
    web1.Height = fraInfo.Height - web1.Top - web1.Left
    
    lblMess.Left = fraInfo.Left + 60
    lblMess.Width = fraInfo.Width - 120
    
    vol1.Width = Me.ScaleWidth - vol1.Left - 120
    vol1.Height = Me.ScaleHeight - vol1.Top - 120
    
    lblClose.Left = Me.ScaleWidth - lblClose.Width - 180
    imgPhotos.Left = lblClose.Left - 180 - imgPhotos.Width
    
    optSort(1).Left = tvw1.Left + tvw1.Width - 250 - optSort(1).Width
    optSort(0).Left = optSort(1).Left - optSort(0).Width
    
    lblExpand(0).Top = tvw1.Top + tvw1.Height + 0
    lblExpand(1).Top = lblExpand(0).Top
    
    picFP.Height = picBack.Height - 60
    picFP.Width = fraInfo.Left + fraInfo.Width + 60
    volFP.Width = picFP.ScaleWidth - (volFP.Left * 2)
    volFP.Height = picFP.ScaleHeight - volFP.Top - 300
    lblDtl.Top = picFP.ScaleHeight - 60 - lblDtl.Height
    picFPClose.Left = picFP.ScaleWidth - picFPClose.Width
End Sub

Public Sub GetFacilities(pSort As Integer)
    Dim i As Integer
    Dim sList As String, strSelect As String, strNest As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim fNode As String, sDesc As String, cNode As String, sNode As String
    Dim sCountry As String
    
    
    lstFCCD.Clear
'    lstFCCD.AddItem 8723
'    lstFCCD.AddItem 4176
'    lstFCCD.AddItem 2750
'    lstFCCD.AddItem 3091
'    lstFCCD.AddItem 12368
'    lstFCCD.AddItem 25237
'    lstFCCD.AddItem 3115
'    lstFCCD.AddItem 3235
'    lstFCCD.AddItem 3478
'    lstFCCD.AddItem 3789
'

'    sList = ""
'    For i = 0 To lstFCCD.ListCount - 1
'        Select Case i
'            Case 0: sList = lstFCCD.List(0)
'            Case Else: sList = sList & ", " & lstFCCD.List(i)
'        End Select
'    Next i
    
    
    cNode = "": sNode = "": fNode = ""
    tvw1.Nodes.Clear
    tvw1.ImageList = ImageList1
    
    strNest = "SELECT AN8_CUNO " & _
                "From ANNOTATOR.DWG_MASTER " & _
                "Where DWGID > 0 " & _
                "AND DWGTYPE = 6 " & _
                "AND DSTATUS > 0"
                        
    Select Case pSort
        Case 0 ''SORT BY LOCATION''
            strSelect = "SELECT DISTINCT AB.ABAN8 AS FCCD, AB.ABALPH AS FACIL, " & _
                        "TRIM(DECODE(TRIM(AL.ALCTR), NULL, 'US', '', 'US', AL.ALCTR)) AS COUNTRY, " & _
                        "NVL(AL.ALADDS, '-') AS STATE " & _
                        "FROM " & F0101 & " AB, " & F0116 & " AL " & _
                        "WHERE AB.ABAN8 IN (" & strNest & ") " & _
                        "AND AB.ABAN8 = AL.ALAN8 " & _
                        "ORDER BY COUNTRY, STATE, FACIL"
                        
'            strSelect = "SELECT DISTINCT AB.ABAN8 AS FCCD, AB.ABALPH AS FACIL, " & _
'                        "DECODE(TRIM(AL.ALCTR), NULL, 'US', '', 'US', AL.ALCTR) AS COUNTRY, " & _
'                        "NVL(AL.ALADDS, '-') AS STATE " & _
'                        "FROM " & F0101 & " AB, " & F0116 & " AL " & _
'                        "WHERE AB.ABAN8 IN (" & strNest & ") " & _
'                        "AND AB.ABAN8 = AL.ALAN8 " & _
'                        "ORDER BY COUNTRY, STATE, FACIL"
                        
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
'                If Trim(rst.Fields("COUNTRY")) = "" Then
'                    sCountry = "US"
'                Else
'                    sCountry = Trim(rst.Fields("COUNTRY"))
'                End If
                If cNode <> "c" & Trim(rst.Fields("COUNTRY")) Then
                    cNode = "c" & Trim(rst.Fields("COUNTRY"))
                    sDesc = GetName("CN", Trim(rst.Fields("COUNTRY")))
                    Set nodX = tvw1.Nodes.Add(, , cNode, sDesc, 2)
                End If
                
                If Trim(rst.Fields("STATE")) <> "-" And Trim(rst.Fields("STATE")) <> "" Then
                    If sNode <> "s" & Trim(rst.Fields("COUNTRY")) & "-" & Trim(rst.Fields("STATE")) Then
                        sNode = "s" & Trim(rst.Fields("COUNTRY")) & "-" & Trim(rst.Fields("STATE"))
                        sDesc = GetName("S", Trim(rst.Fields("STATE")))
                        Set nodX = tvw1.Nodes.Add(cNode, tvwChild, sNode, sDesc, 2)
                    End If
                    
                    fNode = "f" & Trim(rst.Fields("FCCD"))
                    sDesc = Trim(rst.Fields("FACIL"))
                    Set nodX = tvw1.Nodes.Add(sNode, tvwChild, fNode, sDesc, 1)
                    
                    nodX.Tag = GetWebsite(CLng(Trim(rst.Fields("FCCD"))))
                    If nodX.Tag <> "" Then nodX.Image = 7
                    
'                    Select Case CLng(Trim(rst.Fields("FCCD")))
'                        Case 4176
'                            nodX.Tag = "http://www.cobocenter.com"
'                        Case 8723
'                            nodX.Tag = "http://www.mccormickplace.com"
'                    End Select
                    
                Else
                    fNode = "f" & Trim(rst.Fields("FCCD"))
                    sDesc = Trim(rst.Fields("FACIL"))
                    Set nodX = tvw1.Nodes.Add(cNode, tvwChild, fNode, sDesc, 1)
                End If
                
                rst.MoveNext
            Loop
        Case 1 ''SORT BY NAME''
            strSelect = "SELECT DISTINCT AB.ABAN8 AS FCCD, AB.ABALPH AS FACIL " & _
                        "FROM " & F0101 & " AB " & _
                        "WHERE AB.ABAN8 IN (" & strNest & ") " & _
                        "ORDER BY FACIL"
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                fNode = "f" & Trim(rst.Fields("FCCD"))
                sDesc = Trim(rst.Fields("FACIL"))
                Set nodX = tvw1.Nodes.Add(, , fNode, sDesc, 1)
                
                rst.MoveNext
            Loop
    End Select
    
    rst.Close: Set rst = Nothing
    
    
'    Set nodX = tvw1.Nodes.Add(, , "f8723", "Mc Cormick Place South", 1)
'    Set nodX = tvw1.Nodes.Add(, , "f4176", "Cobo Conference Exhibition Ctr", 1)
End Sub


Private Sub imgDirs_Click()
    picBack.Visible = Not picBack.Visible
    vol1.Visible = Not picBack.Visible
    Select Case picBack.Visible
        Case True ''SHOW PLAN ICON''
            imgDirs.Picture = ImageList2.ListImages(2).Picture
        Case False ''SHOW DETAIL ICON''
            imgDirs.Picture = ImageList2.ListImages(1).Picture
    End Select
End Sub

Private Sub imgFPClose_Click()
    picFP.Visible = False
End Sub

Private Sub imgPhotos_Click()
    frmPhoto.PassFCCD = lFCCD
    frmPhoto.Show 1, Me
    
'    Call GetPhotos(lFCCD)
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblCSExpand_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To tvwCS.Nodes.Count
        tvwCS.Nodes(i).Expanded = CBool(Index)
    Next i
End Sub

Private Sub lblExpand_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To tvw1.Nodes.Count
        tvw1.Nodes(i).Expanded = CBool(Index)
    Next i
End Sub

Private Sub lblFPClose_Click()
    Call imgFPClose_Click
End Sub

Private Sub lblFPViewer_Click()
    If volFP.Visible = True Then
        Me.PopupMenu mnuFP, 0, picBack.Left + picFP.Left + imgFPViewer.Left, _
                    picBack.Top + picFP.Top + imgFPViewer.Top + imgFPViewer.Height
    End If
End Sub

Private Sub lblOpenDWF_Click()
    picBack.Visible = False
    vol1.Visible = True
    imgDirs.Picture = ImageList2.ListImages(1).Picture
End Sub

Private Sub lblOthers_Click()
    frmOthers.PassFCCD = lFCCD
    frmOthers.Show 1, Me
    
End Sub

Private Sub lblViewer_Click()
    If vol1.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgViewer.Left, imgViewer.Top + imgViewer.Height
    End If
End Sub

Private Sub mnuDownloadDWF_Click()
    With frmBrowse
        .PassFrom = UCase(Me.Name) & "-DWF"
'        .PassBCC = BCC
'        .PassFBCN = FBCN
'        .PassSHYR = SHYR
'        .PassSHCD = SHCD
'        .PassSHNM = SHNM
        .PassFacil = Mid(lblWelcome.Caption, 12)
        .PassFCCD = lFCCD
        .PassFILETYPE = "DWF"
        .Show 1
    End With
End Sub

Private Sub mnuDownloadPDF_Click()
    With frmBrowse
        .PassFrom = UCase(Me.Name) & "-PDF"
'        .PassBCC = BCC
'        .PassFBCN = FBCN
'        .PassSHYR = SHYR
'        .PassSHCD = SHCD
'        .PassSHNM = SHNM
        .PassFacil = Mid(lblWelcome.Caption, 12)
        .PassFCCD = lFCCD
        .PassFILETYPE = "PDF"
        .Show 1
    End With
End Sub

Private Sub mnuEmailPDF_Click()
    frmEmailFile.PassHDR = Mid(lblWelcome.Caption, 12)
    frmEmailFile.PassFrom = UCase(Me.Name) & "-PDF"
    frmEmailFile.PassFCCD = lFCCD
    frmEmailFile.Show 1, Me
End Sub

Private Sub mnuOSP_Click()
    iRightClick = 1
    timRightClick.Enabled = True
    
    

End Sub

Private Sub mnuVDisplay_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 1 ''2
        If i = Index Then mnuVDisplay(i).Checked = True Else mnuVDisplay(i).Checked = False
    Next i
    Select Case Index
    Case 0
        vol1.GeometryColor = "DefaultColors"
        vol1.BackgroundColor = "DefaultColors"
    Case 1
        vol1.GeometryColor = vbBlack
        vol1.BackgroundColor = vbWhite
'    Case 2
'        volShowplan.GeometryColor = "ClearScale"
'        volShowplan.BackgroundColor = "ClearScale"
    End Select
End Sub

Private Sub mnuVFullView_Click()
    vol1.SetCurrentView dLeft, dRight, dBottom, dTop
End Sub

Private Sub mnuVLayers_Click()
    vol1.ShowLayersDialog
End Sub

Private Sub mnuVPan_Click()
    ClearChecks
    mnuVPan.Checked = True
    vol1.UserMode = "Pan"
End Sub

Private Sub mnuVPrint_Click(Index As Integer)
    If Index = 1 Then
        On Error Resume Next
        vol1.LayersOff = "CAMERA"
    End If
    vol1.ShowPrintDialog
    If Index = 1 Then
        On Error Resume Next
        vol1.LayersOn = "CAMERA"
    End If
End Sub

Private Sub mnuVZoom_Click()
    ClearChecks
    mnuVZoom.Checked = True
    vol1.UserMode = "Zoom"
End Sub

Private Sub mnuVZoomW_Click()
    ClearChecks
    mnuVZoomW.Checked = True
    vol1.UserMode = "ZoomToRect"
End Sub

Public Sub ClearChecks()
    mnuVPan.Checked = False
    mnuVZoom.Checked = False
    mnuVZoomW.Checked = False
End Sub

Public Function InitialView()
    vol1.GetCurrentView dLeft, dRight, dBottom, dTop
End Function

Private Sub mnuWebsite_Click()
    Dim lVal As Long
    Select Case mnuWebsite.Tag
        Case "F": lVal = ShellExecute(0, "open", tvw1.SelectedItem.Tag, 0, 0, 2)
        Case "S": lVal = ShellExecute(0, "open", tvwCS.SelectedItem.Tag, 0, 0, 2) '', 0, 0, 1)
    End Select
End Sub

Private Sub optSort_Click(Index As Integer)
    Call GetFacilities(Index)
    If lFCCD > 0 Then
        tvw1.Nodes("f" & lFCCD).Selected = True
    End If
    lblExpand(0).Enabled = Not CBool(Index)
    lblExpand(1).Enabled = Not CBool(Index)
End Sub

Private Sub timRightClick_Timer()
    timRightClick.Enabled = False
    Select Case iRightClick
        Case 1 ''MNUOSP''
            frmOSP.PassFile = Mid(mnuOSP.Tag, 5)
            frmOSP.PassType = 1
            frmOSP.PassHDR = lblWelcome.Caption
            frmOSP.Show 1, Me
        
    End Select
    iRightClick = 0
End Sub

Private Sub tvw1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iTVWbutton = Button
End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
'    MsgBox Node.Key
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sDWFPath As String
    
    If UCase(Left(Node.Key, 1)) = "F" Then
        If iTVWbutton = vbRightButton Then
            If Node.Tag <> "" Then
                mnuWebsite.Caption = "Go to Facility Website..."
                mnuWebsite.Tag = "F"
                Me.PopupMenu mnuRC
            End If
        Else
            If CLng(Mid(Node.Key, 2)) = lFCCD Then Exit Sub
            cmdGetShows.Enabled = True
            tvwCS.Visible = False
            lblCSExpand(0).Visible = False: lblCSExpand(1).Visible = False
            lFCCD = CLng(Mid(Node.Key, 2))
            lblWelcome.Caption = "Facility:  " & Node.Text
            
            fraInfo.Visible = True
            fraCS.Visible = True
            
            ''CHECK FOR DWF''
            strSelect = "SELECT DF.DWFID, DF.DWFPATH, DF.DWFSTATUS, DF.DWFDESC " & _
                        "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                        "Where DM.AN8_CUNO = " & lFCCD & " " & _
                        "AND DM.DWGTYPE = 6 " & _
                        "AND DM.DWGID = DS.DWGID " & _
                        "AND DS.DWGID = DF.DWGID " & _
                        "AND DS.SHTID = DF.SHTID " & _
                        "AND DF.DWFSTATUS > 0 " & _
                        "ORDER BY DWFSTATUS DESC, DWFID"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sDWFPath = Trim(rst.Fields("DWFPATH"))
                If Dir(sDWFPath, vbNormal) <> "" Then
                    volP.src = sDWFPath
                    volP.Update
                    volP.Visible = True
                    vol1.src = volP.src
                    vol1.Tag = rst.Fields("DWFID")
                    vol1.Update
                    lblReds.Caption = Trim(rst.Fields("DWFDESC"))
                    lblOpenDWF.Visible = volP.Visible
                    imgDirs.Picture = ImageList2.ListImages(2).Picture
                Else
                    volP.Visible = False
                    volP.src = ""
                    vol1.Visible = False
                    vol1.src = ""
                    vol1.Tag = ""
                    lblReds.Caption = ""
                    lblOpenDWF.Visible = False
                    imgDirs.Picture = ImageList2.ListImages(1).Picture
                End If
                rst.Close
            Else
                rst.Close
                volP.Visible = False
                volP.src = ""
                vol1.Visible = False
                vol1.src = ""
                vol1.Tag = ""
                lblReds.Caption = ""
                lblOpenDWF.Visible = False
                imgDirs.Picture = ImageList2.ListImages(1).Picture
            End If
            Set rst = Nothing
            
            imgPhotos.Visible = True '' CheckForPhotos(lFCCD)
            
            Call PopFacilInfo(lFCCD)
            
'            web1.Navigate2 (App.Path & "\Test.htm")
        End If
    End If
End Sub

Public Sub GetShows(pSHYR As Integer, pFCCD As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim tSHYR As Integer
    Dim tSHCD As Long
    Dim YNode As String, sNode As String, cNode As String, _
                DNode As String, ANode As String
    Dim sDesc As String
    Dim nodX As Node
    Dim iIcon As Integer
    
    
    tvwCS.Nodes.Clear
    tSHYR = 0: tSHCD = 0
    tvwCS.ImageList = ImageList1
    Set nodX = tvwCS.Nodes.Add(, , "archive", "Archive of Past Show Years...", 2)
    
    
    
    
    If bClientAll_Enabled Then
        strSelect = "SELECT SH.SHY56SHYR AS SHYR, SH.SHY56SHCD AS SHCD, S.ABALPH AS SHOWNAME, " & _
                    "CS.CSY56CUNO AS CUNO, C.ABALPH AS CLIENT, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56ENDDT, 'DD-MON-YYYY')END_DATE " & _
                    "FROM " & F5601 & " SH, " & F5611 & " CS, " & F0101 & " S, " & F0101 & " C " & _
                    "Where SH.SHY56SHYR >= " & pSHYR & " " & _
                    "AND SH.SHY56FCCDT = " & pFCCD & " " & _
                    "AND SH.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SH.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SH.SHY56SHCD = S.ABAN8 " & _
                    "AND CS.CSY56CUNO = C.ABAN8 " & _
                    "ORDER BY SHYR, SHOWNAME, CLIENT"
    Else
        strSelect = "SELECT SH.SHY56SHYR AS SHYR, SH.SHY56SHCD AS SHCD, S.ABALPH AS SHOWNAME, " & _
                    "CS.CSY56CUNO AS CUNO, C.ABALPH AS CLIENT, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56ENDDT, 'DD-MON-YYYY')END_DATE " & _
                    "FROM " & F5601 & " SH, " & F5611 & " CS, " & F0101 & " S, " & F0101 & " C " & _
                    "Where SH.SHY56SHYR >= " & pSHYR & " " & _
                    "AND SH.SHY56FCCDT = " & pFCCD & " " & _
                    "AND SH.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SH.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                    "AND CS.CSY56CUNO = C.ABAN8 " & _
                    "AND SH.SHY56SHCD = S.ABAN8 " & _
                    "ORDER BY SHYR, SHOWNAME, CLIENT"
    End If
                    
                    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If tSHYR = 0 Or tSHYR <> rst.Fields("SHYR") Then
            tSHYR = rst.Fields("SHYR")
            YNode = "y" & tSHYR
            sDesc = rst.Fields("SHYR")
            Set nodX = tvwCS.Nodes.Add(, , YNode, sDesc, 2)
            nodX.Bold = True
            tSHCD = 0
        End If
        
        If tSHCD = 0 Or tSHCD <> rst.Fields("SHCD") Then
            tSHCD = rst.Fields("SHCD")
            sNode = "s" & tSHYR & "-" & tSHCD
            sDesc = Trim(rst.Fields("SHOWNAME"))
            Set nodX = tvwCS.Nodes.Add(YNode, tvwChild, sNode, sDesc, 2)
'            If sDesc = "Chicago Auto Show" Then
'                nodX.Tag = "http://www.chicagoautoshow.com"
'            ElseIf sDesc = "North American Intl Auto Show" Then
'                nodX.Tag = "http://www.naias.com"
'            End If
            DNode = "d" & tSHYR & "-" & tSHCD
            sDesc = "Show Dates: " & Format(Trim(rst.Fields("BEG_DATE")), "DDD, MMM D, YYYY") & _
                        " - " & Format(Trim(rst.Fields("END_DATE")), "DDD, MMM D")
            Set nodX = tvwCS.Nodes.Add(sNode, tvwChild, DNode, sDesc, 4)
            
            ANode = "a" & tSHYR & "-" & tSHCD
            sDesc = "Attending Clients:"
            Set nodX = tvwCS.Nodes.Add(sNode, tvwChild, ANode, sDesc, 5)
        End If
        
        cNode = "c" & tSHYR & "-" & tSHCD & "-" & rst.Fields("CUNO")
        sDesc = Trim(rst.Fields("CLIENT"))
        iIcon = CheckForFPIcon(tSHYR, tSHCD, rst.Fields("CUNO"))
        Set nodX = tvwCS.Nodes.Add(ANode, tvwChild, cNode, sDesc, iIcon)
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
        
    Dim i As Integer
    For i = tvwCS.Nodes.Count To 1 Step -1
        If UCase(Left(tvwCS.Nodes(i).Key, 1)) = "Y" Then
            tvwCS.Nodes(i).Expanded = True
        End If
    Next i
    
    
End Sub


Public Function GetName(pType As String, pVal As String) As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT DRDL01 as VAL " & _
                "From " & F0005 & " " & _
                "WHERE DRSY = '00' " & _
                "AND DRRT = '" & pType & "' " & _
                "AND DRKY = '" & Space(7) & pVal & "'"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If Trim(rst.Fields("VAL")) <> "" Then
            GetName = Trim(rst.Fields("VAL"))
        Else
            GetName = ""
        End If
    Else
        GetName = ""
    End If
    rst.Close: Set rst = Nothing
    
End Function

Private Sub tvwCS_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim RetVal
    If UCase(Left(Node.Key, 1)) = "A" And Node.Children = 0 Then
        ''POP PAST SHOW YEARS ''
        Screen.MousePointer = 11
        Call PopPastYears(lFCCD, Year(Date))
        Screen.MousePointer = 0
    
    ElseIf UCase(Left(Node.Key, 1)) = "C" Then
        ''cSHYR-SHCD-CUNO''
        picFP.Visible = CheckForFloorplan(Node.Key)
        lblDtl.Visible = picFP.Visible
        
'        lblFPC.Caption = Node.Key
'        lblFPS.Caption = Node.Key
'        picFP.Visible = True
    ElseIf Node.Tag <> "" Then
        mnuWebsite.Caption = "Go to Show Website..."
        mnuWebsite.Tag = "S"
        Me.PopupMenu mnuRC
        
    End If

End Sub

Public Sub PopFacilInfo(pFCCD As Long)
    Dim rst As ADODB.Recordset
    Dim strSelect As String, sHTML As String, sDate1 As String, sDate2 As String, tFile1 As String
    Dim i As Integer
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, tdOa As String, tdOb As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    
    
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY>": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=2><DIV align=left><FONT SIZE=2 COLOR=""#FF0000 "" FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdOb = "%"" VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & pFCCD & tiC & hdC & vbNewLine
    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & bolO & lblWelcome.Caption & bolC & fC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & trO & tdc2O & "Facility" & tdcC & trC & vbNewLine
    sHTML = sHTML & tblO & vbNewLine
    strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
                "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
                "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
                "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
                "WHERE AB.ABAN8 = " & pFCCD & " " & _
                "AND AB.ABAN8 = AL.ALAN8 " & _
                "AND AL.ALAN8 = WP.WPAN8"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Facility Address:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "70" & tdOb & vbNewLine
        sHTML = sHTML & UCase(Trim(rst.Fields("ABALPH"))) & br & vbNewLine
        If Trim(rst.Fields("ALADD1")) <> "" Then _
                    sHTML = sHTML & UCase(Trim(rst.Fields("ALADD1"))) & br & vbNewLine
        If Trim(rst.Fields("ALADD2")) <> "" Then _
                    sHTML = sHTML & UCase(Trim(rst.Fields("ALADD2"))) & br & vbNewLine
        If Trim(rst.Fields("ALADD3")) <> "" Then _
                    sHTML = sHTML & UCase(Trim(rst.Fields("ALADD3"))) & br & vbNewLine
        If Trim(rst.Fields("ALADD4")) <> "" Then _
                    sHTML = sHTML & UCase(Trim(rst.Fields("ALADD4"))) & br & vbNewLine
        If Trim(rst.Fields("ALCTY1")) <> "" Then _
                    sHTML = sHTML & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                    UCase(Trim(rst.Fields("ALADDS"))) & "  " & _
                    Trim(rst.Fields("ALADDZ")) & br & vbNewLine
        sHTML = sHTML & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        sHTML = sHTML & br & vbNewLine
        sHTML = sHTML & tblO & vbNewLine
                
        Do While Not rst.EOF
            Select Case Trim(rst.Fields("WPPHTP"))
                Case ""
                    sHTML = sHTML & trO & vbNewLine
                    sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Facility Phone:" & bolC & tdC & vbNewLine
                    sHTML = sHTML & tdOa & "70" & tdOb & Trim(rst.Fields("WPAR1")) & _
                                " " & Trim(rst.Fields("WPPH1")) & tdC & vbNewLine
                    sHTML = sHTML & trC & vbNewLine
                Case "FAX"
                    sHTML = sHTML & trO & vbNewLine
                    sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Facility Fax:" & bolC & tdC & vbNewLine
                    sHTML = sHTML & tdOa & "70" & tdOb & Trim(rst.Fields("WPAR1")) & _
                            " " & Trim(rst.Fields("WPPH1")) & tdC & vbNewLine
                    sHTML = sHTML & trC & vbNewLine
            End Select
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    
    
    '///// GET SHOW REG ABSTRACT DATA \\\\\
    strSelect = "SELECT HM.HALLDESC, HM.CLGHGT, HM.CLGUNIT, HM.CLGNOTE, HM.HALLNOTE " & _
                "FROM IGLPROD.SRA_HALLMASTER HM " & _
                "Where HM.AN8_FCCD = " & pFCCD & " " & _
                "ORDER BY HALLDESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
'        sHTML = sHTML & br & vbNewLine
        sHTML = sHTML & trO & tdc2O & "Hall Information" & tdcC & trC & vbNewLine
        Do While Not rst.EOF
            sHTML = sHTML & tblO & vbNewLine
            sHTML = sHTML & trO & vbNewLine
            sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Hall:" & bolC & tdC & vbNewLine
            sHTML = sHTML & tdOa & "70" & tdOb & UCase(Trim(rst.Fields("HALLDESC"))) & tdC & vbNewLine
            sHTML = sHTML & trC & vbNewLine
            If Not IsNull(rst.Fields("HALLNOTE")) Then
                sHTML = sHTML & trO & vbNewLine
                sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Hall Comment:" & bolC & tdC & vbNewLine
                sHTML = sHTML & tdOa & "70" & tdOb & UCase(Trim(rst.Fields("HALLNOTE"))) & tdC & vbNewLine
                sHTML = sHTML & trC & vbNewLine
            End If
            
            sHTML = sHTML & trO & vbNewLine
            sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Hall Ceiling Hgt:" & bolC & tdC & vbNewLine
            sHTML = sHTML & tdOa & "70" & tdOb & ConvertDims(CDbl(rst.Fields("CLGHGT")), rst.Fields("CLGUNIT")) & _
                        tdC & vbNewLine
            sHTML = sHTML & trC & vbNewLine
            If Not IsNull(rst.Fields("CLGNOTE")) Then
                sHTML = sHTML & trO & vbNewLine
                sHTML = sHTML & tdOa & "30" & tdOb & bolO & "Ceiling Note:" & bolC & tdC & vbNewLine
                sHTML = sHTML & tdOa & "70" & tdOb & UCase(Trim(rst.Fields("CLGNOTE"))) & tdC & vbNewLine
                sHTML = sHTML & trC & vbNewLine
            End If
            
            sHTML = sHTML & tblC & vbNewLine
            
            rst.MoveNext
            If Not rst.EOF Then sHTML = sHTML & br & vbNewLine
        Loop
'        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & hr & vbNewLine
    End If
    rst.Close: Set rst = Nothing
        
    
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = App.Path & "\Facility.html"
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    web1.Navigate2 tFile1
    web1.Visible = True
End Sub

Public Function ConvertDims(Num As Double, iUnit As Integer) As String
    Dim Feet As Integer, Inch As Integer, Numer As Integer
    Dim Frac As Currency
    Dim strFrac As String
    Select Case iUnit
        Case 1
            Feet = Int(Num / 12)
            Inch = Int(Num - (Feet * 12))
            Frac = CCur((((Num / 12) - Feet) _
                    * 12) - Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case 2
            Feet = Int(Num)
            Inch = (Num - Feet) * 12
            Frac = Inch - Int(Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case Else
            ConvertDims = "Soon!"
    End Select
End Function


Public Sub GetPhotos(pFCCD As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT GID, GDESC, GPATH " & _
                "From ANNOTATOR.GFX_MASTER " & _
                "Where GID > 0 " & _
                "AND AN8_CUNO = " & pFCCD & " " & _
                "AND GTYPE = 66 " & _
                "AND GSTATUS = 66 " & _
                "ORDER BY GDESC"
End Sub

Private Sub vol1_DoNavigateToURL(ByVal URL As String, ByVal window_name As String, enable_default As Boolean)
    enable_default = False
    mnuOSP.Tag = URL
    Me.PopupMenu mnuOSPShow
'    MsgBox url
'    If Left(sLinkID, 3) = "OSP" Then
'        Me.PopupMenu mnuOSPpop
'    Else
'        mnuEng.Visible = CheckForENG(sLinkID)
'        Me.PopupMenu mnuRC
'    End If

    
End Sub

Private Sub vol1_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuVolo
    End If
End Sub

Private Sub vol1_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            Call InitialView
            bViewSet = True
        End If
    End If
End Sub

Public Sub PopPastYears(pFCCD As Long, pSHYR As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim tSHYR As Integer
    Dim tSHCD As Long
    Dim YNode As String, sNode As String, cNode As String, _
                DNode As String, ANode As String
    Dim sDesc As String
    Dim nodX As Node
    Dim iIcon As Integer
    

'''    strSelect = "SELECT SH.SHY56SHYR AS SHYR, SH.SHY56SHCD AS SHCD, S.ABALPH AS SHOWNAME, " & _
'''                "CS.CSY56CUNO AS CUNO, C.ABALPH AS CLIENT, " & _
'''                "IGL_JDEDATE_TOCHAR(SH.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
'''                "IGL_JDEDATE_TOCHAR(SH.SHY56ENDDT, 'DD-MON-YYYY')END_DATE " & _
'''                "FROM " & F5601 & " SH, " & F5611 & " CS, " & F0101 & " S, " & F0101 & " C " & _
'''                "Where SH.SHY56SHYR < " & pSHYR & " " & _
'''                "AND SH.SHY56FCCDT = " & pFCCD & " " & _
'''                "AND SH.SHY56SHYR = CS.CSY56SHYR " & _
'''                "AND SH.SHY56SHCD = CS.CSY56SHCD " & _
'''                "AND SH.SHY56SHCD = S.ABAN8 " & _
'''                "AND CS.CSY56CUNO = C.ABAN8 " & _
'''                "ORDER BY SHYR, SHOWNAME, CLIENT"
                
    If bClientAll_Enabled Then
        strSelect = "SELECT SH.SHY56SHYR AS SHYR, SH.SHY56SHCD AS SHCD, S.ABALPH AS SHOWNAME, " & _
                    "CS.CSY56CUNO AS CUNO, C.ABALPH AS CLIENT, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56ENDDT, 'DD-MON-YYYY')END_DATE " & _
                    "FROM " & F5601 & " SH, " & F5611 & " CS, " & F0101 & " S, " & F0101 & " C " & _
                    "Where SH.SHY56SHYR < " & pSHYR & " " & _
                    "AND SH.SHY56FCCDT = " & pFCCD & " " & _
                    "AND SH.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SH.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SH.SHY56SHCD = S.ABAN8 " & _
                    "AND CS.CSY56CUNO = C.ABAN8 " & _
                    "ORDER BY SHYR, SHOWNAME, SHCD, CLIENT"
    Else
        strSelect = "SELECT SH.SHY56SHYR AS SHYR, SH.SHY56SHCD AS SHCD, S.ABALPH AS SHOWNAME, " & _
                    "CS.CSY56CUNO AS CUNO, C.ABALPH AS CLIENT, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SH.SHY56ENDDT, 'DD-MON-YYYY')END_DATE " & _
                    "FROM " & F5601 & " SH, " & F5611 & " CS, " & F0101 & " S, " & F0101 & " C " & _
                    "Where SH.SHY56SHYR < " & pSHYR & " " & _
                    "AND SH.SHY56FCCDT = " & pFCCD & " " & _
                    "AND SH.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SH.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                    "AND CS.CSY56CUNO = C.ABAN8 " & _
                    "AND SH.SHY56SHCD = S.ABAN8 " & _
                    "ORDER BY SHYR, SHOWNAME, SHCD, CLIENT"
    End If
                
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If tSHYR = 0 Or tSHYR <> rst.Fields("SHYR") Then
            tSHYR = rst.Fields("SHYR")
            YNode = "y" & tSHYR
            sDesc = rst.Fields("SHYR")
            Set nodX = tvwCS.Nodes.Add("archive", tvwChild, YNode, sDesc, 2)
            nodX.Bold = True
            tSHCD = 0
        End If
        
        If tSHCD = 0 Or tSHCD <> rst.Fields("SHCD") Then
            tSHCD = rst.Fields("SHCD")
            sNode = "s" & tSHYR & "-" & tSHCD
            sDesc = Trim(rst.Fields("SHOWNAME"))
            Set nodX = tvwCS.Nodes.Add(YNode, tvwChild, sNode, sDesc, 2)
            If sDesc = "Chicago Auto Show" Then
                nodX.Tag = "http://www.chicagoautoshow.com"
            ElseIf sDesc = "North American Intl Auto Show" Then
                nodX.Tag = "http://www.naias.com"
            End If
            DNode = "d" & tSHYR & "-" & tSHCD
            sDesc = "Show Dates: " & Format(Trim(rst.Fields("BEG_DATE")), "DDD, MMM D, YYYY") & _
                        " - " & Format(Trim(rst.Fields("END_DATE")), "DDD, MMM D")
            Set nodX = tvwCS.Nodes.Add(sNode, tvwChild, DNode, sDesc, 4)
            
            ANode = "a" & tSHYR & "-" & tSHCD
            sDesc = "Attending Clients:"
            Set nodX = tvwCS.Nodes.Add(sNode, tvwChild, ANode, sDesc, 5)
        End If
        
        cNode = "c" & tSHYR & "-" & tSHCD & "-" & rst.Fields("CUNO")
        sDesc = Trim(rst.Fields("CLIENT"))
        iIcon = CheckForFPIcon(tSHYR, tSHCD, rst.Fields("CUNO"))
        Set nodX = tvwCS.Nodes.Add(ANode, tvwChild, cNode, sDesc, iIcon)
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
        
'    Dim i As Integer
'    For i = tvwCS.Nodes.Count To 1 Step -1
'        If UCase(Left(tvwCS.Nodes(i).Key, 1)) = "Y" Then
'            tvwCS.Nodes(i).Expanded = True
'        End If
'    Next i
End Sub

Public Function CheckForFloorplan(pKey As String) As Boolean
    Dim tSHYR As Integer
    Dim tCUNO As Long, tSHCD As Long
    Dim iDash As Integer, iUCnt As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bFPfound As Boolean
    
    bFPfound = False
    
    ''cSHYR-SHCD-CUNO''
    tSHYR = CInt(Mid(pKey, 2, 4))
    iDash = InStr(7, pKey, "-")
    tSHCD = CLng(Mid(pKey, 7, iDash - 7))
    tCUNO = CLng(Mid(pKey, iDash + 1))
    
    iUCnt = 1
    strSelect = "SELECT DWFID, DWFDESC, DWFPATH, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT " & _
                "From ANNOTATOR.DWG_DWF " & _
                "WHERE DWGID IN (" & _
                    "SELECT DWGID " & _
                    "From ANNOTATOR.DWG_SHOW " & _
                    "Where SHYR = " & tSHYR & " " & _
                    "AND AN8_CUNO = " & tCUNO & " " & _
                    "AND AN8_SHCD = " & tSHCD & ") " & _
                "AND DWFTYPE = 0 " & _
                "AND DWFSTATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If Dir(Trim(rst.Fields("DWFPATH")), vbNormal) <> "" Then
            volFP.src = Trim(rst.Fields("DWFPATH"))
            volFP.Update
            volFP.Visible = True
            lblDtl.Caption = "Posted by " & Trim(rst.Fields("ADDUSER")) & " on " & _
                        Format(Trim(rst.Fields("ADDDTTM")), "ddd, mmm d, yyyy (h:nn ampm)")
            If rst.Fields("UPDCNT") > 1 Then
                lblDtl.Caption = lblDtl.Caption & ".  [Last update: " & _
                            Format(Trim(rst.Fields("UPDDTTM")), "ddd, mmm d, yyyy (h:nn ampm)") & _
                            " by " & Trim(rst.Fields("UPDUSER")) & "]"
            End If
            bFPfound = True
        Else
            volFP.Visible = False
            volFP.src = ""
            volFP.Update
        End If
    Else
        volFP.Visible = False
        volFP.src = ""
        volFP.Update
    End If
    rst.Close: Set rst = Nothing
    
    lblFPC.Caption = tvwCS.Nodes(pKey).Text
    If bFPfound Then
        lblFPS.Caption = tvwCS.Nodes(pKey).Parent.Parent.Parent.Text & " - " & _
                    tvwCS.Nodes(pKey).Parent.Parent.Text
    Else
        lblFPS.Caption = "Sorry...  No Floorplan has been posted"
    End If
    
    CheckForFloorplan = Not CBool(volFP.src = "")
    
End Function

Public Function CheckForFPIcon(pSHYR, pSHCD, pCUNO) As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT DWFID " & _
                "From ANNOTATOR.DWG_DWF " & _
                "WHERE DWGID IN (" & _
                    "SELECT DWGID " & _
                    "From ANNOTATOR.DWG_SHOW " & _
                    "Where SHYR = " & pSHYR & " " & _
                    "AND AN8_CUNO = " & pCUNO & " " & _
                    "AND AN8_SHCD = " & pSHCD & ") " & _
                "AND DWFTYPE = 0 " & _
                "AND DWFSTATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        CheckForFPIcon = 3
    Else
        CheckForFPIcon = 6
    End If
    rst.Close: Set rst = Nothing

End Function

Private Sub volFP_DoNavigateToURL(ByVal URL As String, ByVal window_name As String, enable_default As Boolean)
    enable_default = False
    MsgBox "You will need to open this floorplan in the 'Space Plan Viewer'" & vbNewLine & _
                "module of this Annotator to access the embedded hyperlinks", _
                vbInformation, "Sorry..."
End Sub

Private Sub volFP_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bFPViewSet = False Then
        If StatusCode = 42 Then
            Call InitialFPView
            bFPViewSet = True
        End If
    End If
End Sub

Private Sub volFP_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuFP
    End If
End Sub

Private Sub mnuFPDisplay_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 1 ''2
        If i = Index Then mnuFPDisplay(i).Checked = True Else mnuFPDisplay(i).Checked = False
    Next i
    Select Case Index
    Case 0
        volFP.GeometryColor = "DefaultColors"
        volFP.BackgroundColor = "DefaultColors"
    Case 1
        volFP.GeometryColor = vbBlack
        volFP.BackgroundColor = vbWhite
'    Case 2
'        volShowplan.GeometryColor = "ClearScale"
'        volShowplan.BackgroundColor = "ClearScale"
    End Select
End Sub

Private Sub mnuFPFullView_Click()
    volFP.SetCurrentView dFPLeft, dFPRight, dFPBottom, dFPTop
End Sub

Private Sub mnuFPLayers_Click()
    volFP.ShowLayersDialog
End Sub

Private Sub mnuFPPan_Click()
    ClearFPChecks
    mnuFPPan.Checked = True
    volFP.UserMode = "Pan"
End Sub

Private Sub mnuFPPrint_Click()
    volFP.ShowPrintDialog
End Sub

Private Sub mnuFPZoom_Click()
    ClearFPChecks
    mnuFPZoom.Checked = True
    volFP.UserMode = "Zoom"
End Sub

Private Sub mnuFPZoomW_Click()
    ClearFPChecks
    mnuFPZoomW.Checked = True
    volFP.UserMode = "ZoomToRect"
End Sub

Public Sub ClearFPChecks()
    mnuFPPan.Checked = False
    mnuFPZoom.Checked = False
    mnuFPZoomW.Checked = False
End Sub

Public Function InitialFPView()
    volFP.GetCurrentView dFPLeft, dFPRight, dFPBottom, dFPTop
End Function


Public Function GetWebsite(pFID As Long) As String
    Dim strSelect As String, sURL As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT NVL(EAEMAL, '') AS URL " & _
                "FROM " & F0115 & "1 " & _
                "WHERE EAAN8 = " & pFID & " " & _
                "AND EAETP = 'I'"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sURL = Trim(rst.Fields("URL"))
    Else
        sURL = ""
    End If
    rst.Close: Set rst = Nothing
    
    If InStr(1, sURL, "@") > 0 Then sURL = ""
    GetWebsite = sURL
    
End Function
