VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDIL 
   BackColor       =   &H00000000&
   Caption         =   "Digital Image Library"
   ClientHeight    =   9555
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDIL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDLCart 
      BackColor       =   &H00666666&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6660
      MouseIcon       =   "frmDIL.frx":08CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   600
      ScaleWidth      =   1200
      TabIndex        =   72
      ToolTipText     =   "Click to access your Download Cart"
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      Begin VB.Label lblDLCart 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Files"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   728
         MouseIcon       =   "frmDIL.frx":0BD4
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   360
         Width           =   315
      End
      Begin VB.Line lin 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   1185
         X2              =   1185
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line lin 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   570
      End
      Begin VB.Image imgDLCart 
         Height          =   600
         Left            =   120
         MouseIcon       =   "frmDIL.frx":0EDE
         MousePointer    =   99  'Custom
         Picture         =   "frmDIL.frx":11E8
         ToolTipText     =   "Click to access your Download Cart"
         Top             =   -15
         Width           =   450
      End
      Begin VB.Label lblDLCnt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "222"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   660
         MouseIcon       =   "frmDIL.frx":1BD2
         MousePointer    =   99  'Custom
         TabIndex        =   73
         ToolTipText     =   "Click to access your Download Cart"
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin MSComctlLib.ImageList imlSkins 
      Left            =   5100
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1600
      ImageHeight     =   75
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":1EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":59D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":B1C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":126F58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlDirs 
      Left            =   1980
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":127FF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":1288CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":1291A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":129A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":12A35B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":12AC35
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":12B50F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":12BDE9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picWait 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3420
      ScaleHeight     =   555
      ScaleWidth      =   2415
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...Retrieving Images..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.PictureBox picDirs 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6435
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   11535
      TabIndex        =   5
      Top             =   600
      Width           =   11535
      Begin VB.PictureBox picIconSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         ScaleHeight     =   315
         ScaleWidth      =   1260
         TabIndex        =   68
         Top             =   6060
         Width           =   1260
         Begin VB.OptionButton optIconSize 
            Height          =   315
            Index           =   2
            Left            =   840
            Picture         =   "frmDIL.frx":12C6C3
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Display Large Thumbnails"
            Top             =   0
            Width           =   420
         End
         Begin VB.OptionButton optIconSize 
            Height          =   315
            Index           =   1
            Left            =   420
            Picture         =   "frmDIL.frx":12C7DD
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Display Medium Thumbnails"
            Top             =   0
            Value           =   -1  'True
            Width           =   420
         End
         Begin VB.OptionButton optIconSize 
            Height          =   315
            Index           =   0
            Left            =   0
            Picture         =   "frmDIL.frx":12C8D3
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Display Small Thumbnails"
            Top             =   0
            Width           =   420
         End
      End
      Begin VB.Frame fraBatch 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   4500
         TabIndex        =   49
         Top             =   6060
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Label lblBatch 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   600
            MouseIcon       =   "frmDIL.frx":12C9B1
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   23
            Width           =   615
         End
         Begin VB.Label lblBatch 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1500
            MouseIcon       =   "frmDIL.frx":12CCBB
            MousePointer    =   99  'Custom
            TabIndex        =   56
            Top             =   23
            Width           =   345
         End
         Begin VB.Label lblBatch 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   0
            MouseIcon       =   "frmDIL.frx":12CFC5
            MousePointer    =   99  'Custom
            TabIndex        =   55
            Top             =   23
            Width           =   315
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   420
            MouseIcon       =   "frmDIL.frx":12D2CF
            TabIndex        =   54
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1320
            MouseIcon       =   "frmDIL.frx":12D5D9
            TabIndex        =   53
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   1920
            MouseIcon       =   "frmDIL.frx":12D8E3
            TabIndex        =   52
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text List..."
            Height          =   195
            Left            =   2880
            MouseIcon       =   "frmDIL.frx":12DBED
            MousePointer    =   99  'Custom
            TabIndex        =   51
            Top             =   23
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblBatch 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   2100
            MouseIcon       =   "frmDIL.frx":12DEF7
            MousePointer    =   99  'Custom
            TabIndex        =   50
            Top             =   23
            Width           =   300
         End
      End
      Begin VB.Frame fraCnt 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   7620
         TabIndex        =   64
         Top             =   6120
         Width           =   3735
         Begin VB.Label lblCnt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3660
            TabIndex        =   65
            Top             =   23
            Width           =   45
         End
      End
      Begin VB.Frame fraBack 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         TabIndex        =   63
         Top             =   60
         Width           =   1035
      End
      Begin VB.PictureBox picMulti 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3435
         Left            =   1680
         ScaleHeight     =   3435
         ScaleWidth      =   2955
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   2955
         Begin VB.Image imgPopClose 
            Height          =   240
            Left            =   2640
            MouseIcon       =   "frmDIL.frx":12E201
            MousePointer    =   99  'Custom
            Picture         =   "frmDIL.frx":12E50B
            ToolTipText     =   "Click to Close"
            Top             =   60
            Width           =   240
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
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   2685
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   8100
         TabIndex        =   33
         Top             =   180
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email Copy Mode..."
            Height          =   195
            Left            =   1740
            MouseIcon       =   "frmDIL.frx":12E655
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   30
            Width           =   1395
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1500
            TabIndex        =   35
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblDownload 
            AutoSize        =   -1  'True
            Caption         =   "Download Mode..."
            Height          =   195
            Left            =   60
            MouseIcon       =   "frmDIL.frx":12E95F
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   30
            Width           =   1320
         End
      End
      Begin VB.HScrollBar hsc1 
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5820
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.PictureBox picOuter 
         BackColor       =   &H80000005&
         Height          =   5535
         Index           =   0
         Left            =   4500
         ScaleHeight     =   5475
         ScaleWidth      =   6795
         TabIndex        =   6
         Top             =   480
         Width           =   6855
         Begin VB.PictureBox picInner 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawMode        =   6  'Mask Pen Not
            DrawStyle       =   2  'Dot
            ForeColor       =   &H80000008&
            Height          =   5475
            Index           =   0
            Left            =   0
            ScaleHeight     =   5475
            ScaleWidth      =   6375
            TabIndex        =   7
            Top             =   0
            Width           =   6375
            Begin VB.CheckBox chkMulti 
               BackColor       =   &H0000FFFF&
               Height          =   195
               Index           =   0
               Left            =   120
               MaskColor       =   &H0000FFFF&
               TabIndex        =   37
               Top             =   120
               Visible         =   0   'False
               Width           =   195
            End
            Begin IMAGXPR5LibCtl.ImagXpress imx0 
               Height          =   1200
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   120
               Visible         =   0   'False
               Width           =   1600
               _ExtentX        =   2831
               _ExtentY        =   2117
               ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
               ErrCode         =   1288381336
               ErrInfo         =   -275179512
               Persistence     =   -1  'True
               _cx             =   132055552
               _cy             =   1
               FileName        =   ""
               MouseIcon       =   "frmDIL.frx":12EC69
               MousePointer    =   99
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               BackColor       =   -2147483643
               AutoSize        =   4
               BorderType      =   0
               ShowDithered    =   1
               ScrollBars      =   0
               ScrollBarLargeChangeH=   10
               ScrollBarSmallChangeH=   1
               Multitask       =   0   'False
               CancelMode      =   0
               CancelLoad      =   0   'False
               CancelRemove    =   0   'False
               Palette         =   0
               ShowHourglass   =   0   'False
               LZWPassword     =   ""
               PlaceHolder     =   ""
               PFileName       =   ""
               PICPassword     =   ""
               PrinterBanding  =   0   'False
               UndoEnabled     =   0   'False
               Update          =   -1  'True
               CropX           =   0
               CropY           =   0
               SaveGIFType     =   0
               SaveTIFCompression=   0
               SavePNGInterlaced=   0   'False
               SaveGIFInterlaced=   0   'False
               SaveGIFTransparent=   0   'False
               SaveJPGProgressive=   0   'False
               SaveJPGGrayscale=   0   'False
               SaveGIFTColor   =   0
               TwainProductName=   ""
               TwainProductFamily=   ""
               TwainManufacturer=   ""
               TwainVersionInfo=   ""
               Notify          =   0   'False
               NotifyDelay     =   0
               SavePBMType     =   0
               SavePGMType     =   0
               SavePPMType     =   0
               PageNbr         =   0
               ProgressEnabled =   0   'False
               ManagePalette   =   -1  'True
               PictureEnabled  =   -1  'True
               SaveJPGLumFactor=   25
               SaveJPGChromFactor=   35
               DisplayMode     =   0
               DrawStyle       =   1
               DrawWidth       =   1
               DrawFillColor   =   0
               DrawFillStyle   =   1
               DrawMode        =   13
               PICThumbnail    =   0
               PICCropEnabled  =   0   'False
               PICCropX        =   0
               PICCropY        =   0
               PICCropWidth    =   1
               PICCropHeight   =   1
               Antialias       =   0
               SaveJPGSubSampling=   2
               OLEDropMode     =   0
               CompressInMemory=   0
            End
            Begin VB.Label lbl0 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   705
               TabIndex        =   9
               Top             =   1080
               UseMnemonic     =   0   'False
               Visible         =   0   'False
               Width           =   75
            End
         End
      End
      Begin MSComctlLib.TreeView tvwGraphics 
         Height          =   5535
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   9763
         _Version        =   393217
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
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
         MouseIcon       =   "frmDIL.frx":12EF83
      End
      Begin VB.CommandButton cmdDirs 
         Enabled         =   0   'False
         Height          =   3075
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Width           =   6795
      End
      Begin VB.Label Label1 
         BackColor       =   &H001CAF6F&
         Caption         =   "GPJ Digital Image Library"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   180
         TabIndex        =   12
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   3450
      End
   End
   Begin VB.PictureBox picMenu2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   1260
      TabIndex        =   58
      Top             =   1170
      Visible         =   0   'False
      Width           =   1260
      Begin VB.Label lblResize 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Resize"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmDIL.frx":12F29D
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblFullSize 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmDIL.frx":12F5A7
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblKeyEdit 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmDIL.frx":12F8B1
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   780
         Width           =   735
      End
      Begin VB.Image imgResize 
         Height          =   360
         Left            =   0
         Picture         =   "frmDIL.frx":12FBBB
         Top             =   0
         Width           =   1260
      End
      Begin VB.Image imgFullSize 
         Height          =   360
         Left            =   0
         Picture         =   "frmDIL.frx":12FEDD
         Top             =   360
         Width           =   1260
      End
      Begin VB.Image imgKeyEdit 
         Height          =   360
         Left            =   0
         Picture         =   "frmDIL.frx":1301FF
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3060
      ScaleHeight     =   1095
      ScaleWidth      =   4695
      TabIndex        =   13
      Top             =   7380
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing in process... Depending on Graphic File size, Printing may take a few moments..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   4275
      End
   End
   Begin VB.PictureBox picResult 
      ForeColor       =   &H80000009&
      Height          =   2055
      Left            =   7920
      ScaleHeight     =   1995
      ScaleWidth      =   2535
      TabIndex        =   18
      Top             =   7140
      Visible         =   0   'False
      Width           =   2595
      Begin VB.ListBox lstResultPath 
         Height          =   255
         ItemData        =   "frmDIL.frx":130521
         Left            =   1320
         List            =   "frmDIL.frx":130523
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox lstResult 
         Height          =   255
         ItemData        =   "frmDIL.frx":130525
         Left            =   60
         List            =   "frmDIL.frx":130527
         TabIndex        =   19
         Top             =   300
         Width           =   2415
      End
      Begin VB.Image imgResult 
         Height          =   240
         Left            =   2280
         Picture         =   "frmDIL.frx":130529
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblResult 
         BackColor       =   &H80000002&
         Caption         =   " Search Result..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox picJPG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6930
      Left            =   1260
      ScaleHeight     =   6930
      ScaleMode       =   0  'User
      ScaleWidth      =   10440
      TabIndex        =   0
      Top             =   1140
      Width           =   10440
      Begin VB.Label lblByGeorge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digital Image Library"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   825
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   5700
         Visible         =   0   'False
         Width           =   6165
      End
      Begin VB.Label lblByGeorge 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "By George!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   96
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   2175
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   4860
         Visible         =   0   'False
         Width           =   8985
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   8010
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
            Picture         =   "frmDIL.frx":13096B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":130F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":13149F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":131A39
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":131FD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":1328AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDIL.frx":132E47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   1380
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRedline 
      Caption         =   "Redline..."
      Height          =   315
      Left            =   60
      TabIndex        =   67
      Top             =   7500
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   795
      Left            =   420
      TabIndex        =   66
      Top             =   7860
      Visible         =   0   'False
      Width           =   675
      ExtentX         =   1191
      ExtentY         =   1402
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
   Begin VB.PictureBox picViewer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   3960
      Picture         =   "frmDIL.frx":1333E1
      ScaleHeight     =   6015
      ScaleWidth      =   6015
      TabIndex        =   40
      Top             =   1980
      Visible         =   0   'False
      Width           =   6015
      Begin IMAGXPR5LibCtl.ImagXpress imxViewer 
         Height          =   3300
         Left            =   600
         TabIndex        =   45
         Top             =   960
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   5821
         ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
         ErrCode         =   1288381336
         ErrInfo         =   -275179512
         Persistence     =   -1  'True
         _cx             =   132055184
         _cy             =   1
         FileName        =   ""
         MouseIcon       =   "frmDIL.frx":13446C
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         BackColor       =   0
         AutoSize        =   4
         BorderType      =   0
         ShowDithered    =   1
         ScrollBars      =   0
         ScrollBarLargeChangeH=   10
         ScrollBarSmallChangeH=   1
         Multitask       =   0   'False
         CancelMode      =   0
         CancelLoad      =   0   'False
         CancelRemove    =   0   'False
         Palette         =   0
         ShowHourglass   =   0   'False
         LZWPassword     =   ""
         PlaceHolder     =   ""
         PFileName       =   ""
         PICPassword     =   ""
         PrinterBanding  =   0   'False
         UndoEnabled     =   0   'False
         Update          =   -1  'True
         CropX           =   0
         CropY           =   0
         SaveGIFType     =   0
         SaveTIFCompression=   0
         SavePNGInterlaced=   0   'False
         SaveGIFInterlaced=   0   'False
         SaveGIFTransparent=   0   'False
         SaveJPGProgressive=   0   'False
         SaveJPGGrayscale=   0   'False
         SaveGIFTColor   =   0
         TwainProductName=   ""
         TwainProductFamily=   ""
         TwainManufacturer=   ""
         TwainVersionInfo=   ""
         Notify          =   0   'False
         NotifyDelay     =   0
         SavePBMType     =   0
         SavePGMType     =   0
         SavePPMType     =   0
         PageNbr         =   0
         ProgressEnabled =   0   'False
         ManagePalette   =   -1  'True
         PictureEnabled  =   -1  'True
         SaveJPGLumFactor=   25
         SaveJPGChromFactor=   35
         DisplayMode     =   0
         DrawStyle       =   1
         DrawWidth       =   1
         DrawFillColor   =   0
         DrawFillStyle   =   1
         DrawMode        =   13
         PICThumbnail    =   0
         PICCropEnabled  =   0   'False
         PICCropX        =   0
         PICCropY        =   0
         PICCropWidth    =   1
         PICCropHeight   =   1
         Antialias       =   0
         SaveJPGSubSampling=   2
         OLEDropMode     =   0
         CompressInMemory=   0
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "...File Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4905
         MouseIcon       =   "frmDIL.frx":134786
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   4755
         Width           =   900
      End
      Begin VB.Label lblDisclaimer 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDIL.frx":134A90
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   1725
         TabIndex        =   43
         Top             =   5160
         Width           =   4065
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename..."
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
         Left            =   240
         MouseIcon       =   "frmDIL.frx":134B2A
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label lblPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview File..."
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
         Left            =   195
         MouseIcon       =   "frmDIL.frx":134E34
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   4755
         Width           =   1335
      End
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Left            =   10320
      Picture         =   "frmDIL.frx":13513E
      ToolTipText     =   "Click to access the Search Tool"
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSupDoc 
      Height          =   480
      Left            =   120
      Picture         =   "frmDIL.frx":135A08
      ToolTipText     =   "Click to view the Supoort Document for the current Image"
      Top             =   2340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   300
      MouseIcon       =   "frmDIL.frx":1362D2
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   750
      Visible         =   0   'False
      Width           =   465
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
      Left            =   11085
      MouseIcon       =   "frmDIL.frx":1365DC
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   180
      Width           =   510
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H000E5838&
      Height          =   195
      Left            =   11100
      MouseIcon       =   "frmDIL.frx":1368E6
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   540
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmDIL.frx":136BF0
      MousePointer    =   99  'Custom
      Picture         =   "frmDIL.frx":136EFA
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Width           =   720
   End
   Begin VB.Label lblViewAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All"
      Height          =   195
      Index           =   0
      Left            =   2460
      MouseIcon       =   "frmDIL.frx":137A44
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "181-200"
      Height          =   195
      Index           =   9
      Left            =   9180
      MouseIcon       =   "frmDIL.frx":137D4E
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   9300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "161-180"
      Height          =   195
      Index           =   8
      Left            =   8400
      MouseIcon       =   "frmDIL.frx":138058
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   9300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "141-160"
      Height          =   195
      Index           =   7
      Left            =   7620
      MouseIcon       =   "frmDIL.frx":138362
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   9300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "121-140"
      Height          =   195
      Index           =   6
      Left            =   6840
      MouseIcon       =   "frmDIL.frx":13866C
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   9300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "101-120"
      Height          =   195
      Index           =   5
      Left            =   6060
      MouseIcon       =   "frmDIL.frx":138976
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   9300
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81-100"
      Height          =   195
      Index           =   4
      Left            =   5400
      MouseIcon       =   "frmDIL.frx":138C80
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   9300
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "61-80"
      Height          =   195
      Index           =   3
      Left            =   4800
      MouseIcon       =   "frmDIL.frx":138F8A
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   9300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41-60"
      Height          =   195
      Index           =   2
      Left            =   4200
      MouseIcon       =   "frmDIL.frx":139294
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   9300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-40"
      Height          =   195
      Index           =   1
      Left            =   3600
      MouseIcon       =   "frmDIL.frx":13959E
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   9300
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1-20"
      Height          =   195
      Index           =   0
      Left            =   3120
      MouseIcon       =   "frmDIL.frx":1398A8
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   9300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   8340
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgSize 
      Height          =   495
      Left            =   1980
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Library has loaded..."
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
      Left            =   900
      TabIndex        =   4
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   2505
   End
   Begin VB.Label lblGraphic 
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
      Left            =   1260
      TabIndex        =   3
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Image imgClose 
      Height          =   945
      Left            =   10800
      Picture         =   "frmDIL.frx":139BB2
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgMenu 
      Height          =   570
      Left            =   0
      Picture         =   "frmDIL.frx":13A15B
      Top             =   600
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgBack 
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgBack 
      Height          =   735
      Index           =   1
      Left            =   8700
      Top             =   0
      Width           =   2235
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuResizeGraphic 
         Caption         =   "Resize Graphic to Actual Size"
      End
      Begin VB.Menu mnuMaxGraphic 
         Caption         =   "Maximize Graphic"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendALink 
         Caption         =   "Send-A-Link..."
      End
      Begin VB.Menu mnuEmailSel 
         Caption         =   "Email Image File..."
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download Image File..."
      End
      Begin VB.Menu mnuDownloadAdd 
         Caption         =   "Add Image to Download Cart"
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGFXData 
         Caption         =   "View Graphic Data..."
      End
      Begin VB.Menu mnuGPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuDownloadMulti 
      Caption         =   "mnuDownloadMulti"
      Visible         =   0   'False
      Begin VB.Menu mnuDownloadMode 
         Caption         =   "Activate Download Mode..."
      End
      Begin VB.Menu mnuDownloadSels 
         Caption         =   "Download Selections..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownloadSelsAdd 
         Caption         =   "Add Selections to Download Cart"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEmailMulti 
      Caption         =   "mnuEmailMulti"
      Visible         =   0   'False
      Begin VB.Menu mnuEmailMode 
         Caption         =   "Activate Email Copy Mode..."
      End
      Begin VB.Menu mnuEmailSels 
         Caption         =   "Email Copy of Selections..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOptArray 
      Caption         =   "mnuOptArray"
      Visible         =   0   'False
      Begin VB.Menu mnuOptSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuOptClearAll 
         Caption         =   "Clear All"
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmailSels2 
         Caption         =   "Email Copy of Selections..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownloadSels2 
         Caption         =   "Download Selections..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownloadSelsAdd2 
         Caption         =   "Add Selections to Download Cart"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmDIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim maxX As Double, maxY As Double, dTop As Double, dLeft As Double, _
            dGTop As Double, dGLeft As Double, _
            dMTop As Double, dMLeft As Double, dSTop As Double, dSLeft As Double
Dim rAsp As Double, rFAsp As Double, rX As Double, rY As Double, rXO As Double, rYO As Double, _
            rMX As Double, rMY As Double, rSX As Double, rSY As Double
Dim sGPath As String
Dim sInType As String
Dim bPicLoaded As Boolean, bPopped As Boolean
Dim iImageState As Integer ''0=Small:1=Max''
Dim TNode As String
Dim sOrder As String
Dim CurrParNode As String, CurrParText As String
Public sTable As String, CurrFile As String
Public CurrSelect As String
Dim xStr As Long, yStr As Long
Dim iListStart As Integer, iGFXCount As Integer
Dim bEMode As Boolean, bDMode As Boolean
Dim iIMXIndex As Integer
Public bDirsOpen As Boolean, bFirst As Boolean

Dim xStart As Single, yStart As Single, bMouseDown As Boolean
Dim xs, ys


Public pDownloadPath As String
Public Property Get PassDLPath() As String
    PassDLPath = pDownloadPath
End Property
Public Property Let PassDLPath(ByVal vNewValue As String)
    pDownloadPath = vNewValue
End Property



Private Sub chkMulti_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bFound As Boolean
    Dim i As Integer
    
    bFound = False
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Value = 1 Then
            bFound = True
            Exit For
        End If
    Next i
    
    If bEMode Then
        mnuEmailSels.Enabled = bFound
        mnuEmailSels2.Enabled = bFound
    End If
    If bDMode Then
        mnuDownloadSels.Enabled = bFound
        mnuDownloadSelsAdd.Enabled = bFound
        mnuDownloadSels2.Enabled = bFound
        mnuDownloadSelsAdd2.Enabled = bFound
    End If
    
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub cmdRedline_Click()
    frmPPTRedline.PassHeight = web1.Height
    frmPPTRedline.PassLeft = web1.Left
    frmPPTRedline.PassTop = web1.Top
    frmPPTRedline.PassWidth = CLng(web1.Height * (4 / 3))
    frmPPTRedline.Show 1, Me
End Sub

Private Sub Form_Click()
    Dim lErr As Long
    
    lErr = LockWindowUpdate(Me.hwnd)
    If Me.BackColor = vbWhite Then
'        Set Me.Picture = imlSkins.ListImages(1).Picture
        Me.BackColor = vbBlack
'        picJPG.BackColor = vbBlack
        lblStatus.ForeColor = vbWhite
        lblGraphic.ForeColor = vbWhite
        picMenu2.BackColor = vbBlack
        picViewer.BackColor = vbBlack
    Else
'        Set Me.Picture = imlSkins.ListImages(2).Picture
        Me.BackColor = vbWhite
        picMenu2.BackColor = vbWhite
'        picJPG.BackColor = vbWhite
        lblStatus.ForeColor = vbBlack
        lblGraphic.ForeColor = vbBlack
        picViewer.BackColor = vbWhite
    End If
    lErr = LockWindowUpdate(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Resp As VbMsgBoxResult
    
    If picDLCart.Visible And CInt(lblDLCnt.Caption) > 0 Then
        Resp = MsgBox("You have files in your Download Cart.  " & _
                    "Would you like to review them before closing " & _
                    "the Digital Image Library?", vbQuestion + vbYesNoCancel, "Download Cart...")
        Select Case Resp
            Case vbNo
                Resp = MsgBox("Would you like to clear your Download Cart's Digital Image Library files?", _
                            vbQuestion + vbYesNo, "Empty your Cart?")
                If Resp = vbYes Then Call ClearDownloadCart("DIL")
                bPopped = False
            Case vbYes
                Call imgDLCart_Click
                bPopped = False
            Case vbCancel
                Cancel = 1
        End Select
    Else
        bPopped = False
    End If
    
End Sub

'''Private Sub fraBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case bDirsOpen
'''        Case True: Set imgDirs.Picture = imlDirs.ListImages(3).Picture
'''        Case False: Set imgDirs.Picture = imlDirs.ListImages(1).Picture
'''    End Select
'''End Sub

Private Sub imgBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
'''        Case 0
'''            Select Case bDirsOpen
'''                Case True: Set imgDirs.Picture = imlDirs.ListImages(3).Picture
'''                Case False: Set imgDirs.Picture = imlDirs.ListImages(1).Picture
'''            End Select
        Case 1
            Set imgSearch.Picture = imlDirs.ListImages(5).Picture
            Set imgSupDoc.Picture = imlDirs.ListImages(7).Picture
    End Select
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = lGeo_Back '' vbWhite
    lblClose.ForeColor = lGeo_Back '' vbWhite
End Sub

Private Sub imgDirs_Click()
    If picDirs.Visible = False Then
        picDirs.Visible = True
        bDirsOpen = True
        imgDirs.ToolTipText = "Click to Close File Index"
'        Set imgDirs.Picture = imlDirs.ListImages(4).Picture
    Else
        picDirs.Visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
'        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
    End If
End Sub


Private Sub imgDLCart_Click()
    Screen.MousePointer = 11
    
    frmDownloadCart.PassDLType = "DIL"
    frmDownloadCart.Show 1
    
    ''CHECK FOR FILES AND SET VISIBILITY''
    Call CheckDownloadCart("DIL")
    
    Screen.MousePointer = 0
End Sub

'''Private Sub imgDirs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case bDirsOpen
'''        Case True: Set imgDirs.Picture = imlDirs.ListImages(4).Picture
'''        Case False: Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'''    End Select
'''End Sub

Private Sub imgSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSearch.Picture = imlDirs.ListImages(6).Picture
End Sub

Private Sub imgSupDoc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSupDoc.Picture = imlDirs.ListImages(8).Picture
End Sub

Private Sub lblDLCart_Click()
    Call imgDLCart_Click
End Sub

Private Sub lblDLCnt_Click()
    Call imgDLCart_Click
End Sub

Private Sub lblfullsize_Click()
    frmHTMLViewer.PassFile = CurrFile
    frmHTMLViewer.PassFrom = Me.Name
    frmHTMLViewer.PassHDR = lblWelcome.Caption
    frmHTMLViewer.Show 1, Me
End Sub

Private Sub lblkeyedit_Click()
    frmKeywordEdit.PassGID = lGID
    frmKeywordEdit.PassFrom = "DIL"
    frmKeywordEdit.Show 1, Me
End Sub

Private Sub lblMenu_Click()
    Me.PopupMenu mnuRightClick, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
End Sub

Private Sub lblResize_Click()
    Select Case lblResize.Caption
        Case "Resize"
            mnuResizeGraphic_Click
        Case "Maximize"
            mnuMaxGraphic_Click
    End Select
End Sub

Private Sub imgSearch_Click()
    lOpenInViewer = 0
    frmSearch.PassFrom = "DIL"
    frmSearch.Show 1
    If lOpenInViewer > 0 Then
        MsgBox "Open file# " & lOpenInViewer
    End If
    
    Set imgSearch.Picture = imlDirs.ListImages(5).Picture
End Sub

Private Sub imgSupdoc_Click()
    
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If imgSupDoc.Tag <> "" Then
        strSelect = "SELECT SUPDOC_ID, SUPDOCDESC, SUPDOCFORMAT " & _
                    "FROM ANNOTATOR.GFX_SUPDOC " & _
                    "WHERE SUPDOC_ID = " & imgSupDoc.Tag
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            frmHTMLViewer.PassFile = sSupDocPath & rst.Fields("SUPDOC_ID") & _
                        "." & LCase(Trim(rst.Fields("SUPDOCFORMAT")))
            frmHTMLViewer.PassHDR = "Case Study:  " & Trim(rst.Fields("SUPDOCDESC"))
            frmHTMLViewer.PassDFile = Trim(rst.Fields("SUPDOCDESC")) & _
                        "." & LCase(Trim(rst.Fields("SUPDOCFORMAT")))
            rst.Close: Set rst = Nothing
            frmHTMLViewer.Show 1, Me
        Else
            rst.Close: Set rst = Nothing
            MsgBox "Support Document not found", vbExclamation, "Sorry..."
        End If
        
        Set imgSupDoc.Picture = imlDirs.ListImages(7).Picture
    End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Drag 2
    Source.Move CSng(X - xStr), CSng(Y - yStr)
'''    Source.Move CSng(Me.Left + (X - xStr)), CSng(Me.Top - (Me.Height - Me.ScaleHeight) + (Y - yStr))
End Sub

Private Sub Form_Load()
    Dim i As Integer, iCol As Integer, iRow As Integer
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lOpenID As Long
    
    bDirsOpen = True
    web1.Navigate2 "about:Loading..."
    
    bFirst = True
    optIconSize(iIconSize).Value = True
    bFirst = False
    
    maxX = Me.ScaleWidth - 1260 - 120 '' - 240
    maxY = Me.ScaleHeight - 1140 - 120 '' 960 - 120 '' 795 - 120
    
    sGPath = "\\DETMSFS01\GPJAnnotator\Graphics\"
    sInType = "1, 2, 3, 4"
    
    If UCase(Shortname) = UCase(sOUser) Then
        '///// WRITE DIL_OPEN TO ANO_LOCKLOG \\\\\'
        Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
        lOpenID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        strInsert = "INSERT INTO " & ANOLockLog & " " & _
                    "(LOCKID, LOCKREFID, LOCKREFSOURCE, USER_SEQ_ID, " & _
                    "LOCKOPENDTTM, LOCKSTATUS, ADDUSER, ADDDTTM, " & _
                    "UPDUSER, UPDDTTM, UPDCNT) VALUES " & _
                    "(" & lOpenID & ", 1002, 'DIL_OPEN', " & UserID & ", " & _
                    "SYSDATE, 1, '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
    End If
    
    '///// GRAPHIC TYPES \\\\\
    GfxType(1) = "Digital Photos"
    GfxType(2) = "Graphic Files"
    GfxType(3) = "Graphic Layouts"
    GfxType(4) = "Presentation Files"
    
    Me.WindowState = AppWindowState
    
    dTop = 1140 ''960 ''
    dLeft = 1260 '' 120
    picJPG.Top = dTop
    picJPG.Left = dLeft
    web1.Top = dTop
    web1.Left = dLeft
    cmdRedline.Left = web1.Left + 120
    cmdRedline.Top = web1.Top + 120
    
    lblStatus.Left = 120
    
    picResult.Top = picMenu2.Top
    
            
    lblByGeorge(0).ForeColor = lGeo_Back '' RGB(30, 30, 21)
    lblByGeorge(1).ForeColor = lGeo_Fore '' RGB(100, 100, 68)
    
    picDirs.Left = 0
    picDirs.Top = 600
    picDirs.Width = Me.ScaleWidth
    picDirs.Height = Me.ScaleHeight - picDirs.Top
    tvwGraphics.Height = picDirs.ScaleHeight - tvwGraphics.Top - 420
    picOuter(0).Height = tvwGraphics.Height
    hsc1(0).Top = picOuter(0).Top + picOuter(0).Height - hsc1(0).Height
    picInner(0).Height = picOuter(0).ScaleHeight
    fraBatch.Top = picDirs.ScaleHeight - fraBatch.Height - 60
    fraCnt.Top = fraBatch.Top
    
    picOuter(0).Width = picDirs.ScaleWidth - picOuter(0).Left - tvwGraphics.Left
    picInner(0).Width = picOuter(0).ScaleWidth
    
'''''    iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
'''''    iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
'''''
'''''    imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
'''''    spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
'''''    imageX = CLng((picInner(0).ScaleWidth - 240) / iCols) '' (imageY / 3) * 4
'''''    spaceX = imageX ''CLng(picInner(0).ScaleWidth / iCols) '' imageX + 720 ''240
'''''
'''''    For i = 0 To (iCols * iRows) - 1 ''19
'''''        If i >= imx0.Count Then
'''''            Load imx0(i)
'''''            Set imx0(i).Container = picInner(0)
'''''        End If
'''''        iCol = Int(i / iRows): iRow = i Mod iRows
'''''        imx0(i).Width = lIconX
'''''        imx0(i).Height = lIconY
''''''''        imx0(i).Left = 480 + (iCol * spaceX)
'''''        imx0(i).Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
'''''        imx0(i).Top = 120 + (iRow * spaceY)
'''''
'''''        If i >= lbl0.Count Then Load lbl0(i)
'''''
'''''        If i >= chkMulti.Count Then Load chkMulti(i)
'''''    Next i
    
'    picInner(0).Width = ((iCol + 1) * spaceX) + 240
    sOrder = "ORDER BY GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
    
    Call GetGraphicList
    
    ''SET VIEW BASED ON PERMISSIONS''
    If bPerm(58) Then
        imgKeyEdit.Visible = True
        lblKeyEdit.Visible = True
    Else
        imgKeyEdit.Visible = False
        lblKeyEdit.Visible = False
        picMenu2.Height = imgFullSize.Top + imgFullSize.Height
    End If
    
    If bPassIn Then
        strSelect = "SELECT GM.GID, GF.FLRDESC, GM.GDESC, GM.GPATH " & _
                    "FROM ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                    "Where GM.GID > 0 " & _
                    "AND GM.GID = " & sPassInValue & " " & _
                    "AND GM.AN8_CUNO = 40579 " & _
                    "AND GM.FLR_ID = GF.FLR_ID"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            picDirs.Visible = False
            bDirsOpen = False
            imgDirs.ToolTipText = "Click to Open File Index..."
'''            Set imgDirs.Picture = imlDirs.ListImages(1).Picture
            '///// TIME TO LOAD THE GRAPHIC \\\\\
            Call LoadGraphic(0, sPassInValue, Trim(rst.Fields("GDESC")), _
                        Trim(rst.Fields("FLRDESC")))
        End If
        rst.Close: Set rst = Nothing
            
    End If
    
    Call CheckDownloadCart("DIL")
End Sub

Private Sub Form_Resize()
    Dim i As Integer, iTab As Integer, iCol As Integer, iRow As Integer
    
    If Me.WindowState <> 1 Then
        picOuter(0).Visible = False
        If Me.Width > 8400 And Me.Height > 4000 Then
            maxX = Me.ScaleWidth - 1260 - 120 '' 240
            maxY = Me.ScaleHeight - 1140 - 120 '' 960 - 120 '' 795 - 120
            
            web1.Width = maxX
            web1.Height = maxY
            
            rFAsp = maxX / maxY
            Select Case rAsp
                Case Is = rFAsp, 0
                    With picJPG
                        .Width = maxX
                        .Height = maxY
                        .Top = dTop: .Left = dLeft
                    End With
                Case Is > rFAsp
                    With picJPG
                        .Width = maxX
                        .Height = .Width / rAsp
                        .Top = dTop + ((maxY - .Height) / 2)
                        .Left = dLeft
                    End With
                Case Is < rFAsp
                    With picJPG
                        .Height = maxY
                        .Width = .Height * rAsp ''''' / rFAsp)
                        .Top = dTop
                        .Left = (Me.ScaleWidth - .Width) / 2
                    End With
            End Select
            
            If bPicLoaded Then
                If imgSize.Picture <> 0 Then
                    picJPG.PaintPicture imgSize.Picture, 0, 0, picJPG.Width, picJPG.Height
                    Call SetImageState
                End If
            End If
            
            lblStatus.Top = Me.ScaleHeight - lblStatus.Height - 60
            
            picResult.Left = 0 '' Me.ScaleWidth - picMenu2.Left - picResult.Width
            
            shpHDR.Width = Me.ScaleWidth
            
'''            '///// POSITION LBLBYGEORGEs \\\\\
'''            lblByGeorge(0).Left = maxX - 240 - lblByGeorge(0).Width
'''            lblByGeorge(0).Top = maxY - 240 - lblByGeorge(0).Height
'''            lblByGeorge(1).Left = 240
'''            lblByGeorge(1).Top = lblByGeorge(0).Top + 840
            
            AppWindowState = Me.WindowState
            
            imgClose.Left = Me.ScaleWidth - imgClose.Width
            lblHelp.Left = imgClose.Left + (imgClose.Width / 2) - (lblHelp.Width / 2)
            lblClose.Left = imgClose.Left + (imgClose.Width / 2) - (lblClose.Width / 2)
            
            imgSearch.Left = imgClose.Left - 120 - imgSearch.Width
'''            imgSupDoc.Left = imgSearch.Left - 60 - imgSupDoc.Width
'''            imgBack(1).Left = imgSupDoc.Left - 300
            imgBack(1).Left = imgSearch.Left - 300
'''            picDL.Left = imgSupDoc.Left - 60 - picDL.Width
            picDLCart.Left = imgSearch.Left - 120 - picDLCart.Width
            
            picPrint.Left = (Me.Width - picPrint.Width) / 2
            picPrint.Top = (Me.Height - picPrint.Height) / 2
            
            
            If bPicLoaded Then
                '///// CHECK IF IMAGE IS STRETCHED \\\\\
                If picJPG.Width > imgSize.Width Then
'                    mnuResizeGraphic.Visible = True
'                    mnuMaxGraphic.Visible = True
'                    mnuResizeGraphic.Enabled = True
'                    mnuMaxGraphic.Enabled = False
                    lblResize.Caption = "Resize"
                    lblResize.Enabled = True
                Else
'                    mnuResizeGraphic.Visible = False
'                    mnuMaxGraphic.Visible = False
                    lblResize.Enabled = False
                End If
            End If
            
            picDirs.Width = Me.ScaleWidth '' maxX
            cmdDirs.Width = picDirs.Width
            picDirs.Height = Me.ScaleHeight - picDirs.Top
            cmdDirs.Height = picDirs.Height
            
            tvwGraphics.Height = picDirs.ScaleHeight - tvwGraphics.Top - 420
            picOuter(0).Height = tvwGraphics.Height
            hsc1(0).Top = picOuter(0).Top + picOuter(0).Height - hsc1(0).Height
            picInner(0).Height = picOuter(0).ScaleHeight
            fraBatch.Top = picDirs.ScaleHeight - fraBatch.Height - 60
            fraCnt.Top = fraBatch.Top
            
            picIconSize.Top = tvwGraphics.Top + tvwGraphics.Height + 45
            
            picOuter(0).Width = picDirs.ScaleWidth - picOuter(0).Left - tvwGraphics.Left
            picInner(0).Width = picOuter(0).ScaleWidth
    
            iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
            iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
            
            imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
            spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
            imageX = CLng((picInner(0).ScaleWidth) / iCols) '' (imageY / 3) * 4
            spaceX = imageX ''CLng(picInner(0).ScaleWidth / iCols) '' imageX + 720 ''240
    
            For i = 0 To (iCols * iRows) - 1 ''19
                If i >= imx0.Count Then Load imx0(i)
                iCol = Int(i / iRows): iRow = i Mod iRows
        '''        imx0(i).Left = 480 + (iCol * spaceX)
                imx0(i).Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
                imx0(i).Top = 120 + (iRow * spaceY)
                
                If i >= lbl0.Count Then Load lbl0(i)
                
                If i >= chkMulti.Count Then Load chkMulti(i)
            Next i
            
            If bPopped Then Call tvwGraphics_NodeClick(tvwGraphics.SelectedItem)
            
            picViewer.Left = (Me.ScaleWidth - picViewer.Width) / 2
            picViewer.Top = dTop + (((Me.ScaleHeight - dTop - 300) - picViewer.Height) / 2)
            
            picOuter(0).Left = 4500
            fraBatch.Left = picOuter(0).Left
            If picDirs.ScaleWidth - picOuter(0).Left - 180 > 0 Then
                picOuter(0).Width = picDirs.ScaleWidth - picOuter(0).Left - 180
            End If
            hsc1(0).Width = picOuter(0).Width
'            fraBatch.Width = picOuter(0).Width
            fraCnt.Left = picOuter(0).Left + picOuter(0).Width - fraCnt.Width
'            lblCnt.Left = fraBatch.Width - lblCnt.Width
            
            picWait.Left = picDirs.Left + picOuter(0).Left + _
                        ((picOuter(0).Width - picWait.Width) / 2)
            picWait.Top = picDirs.Top + picOuter(0).Top + _
                        ((picOuter(0).Height - picWait.Height) / 2)
                        
'            lblCnt.Left = picOuter(0).Left + picOuter(0).Width - lblCnt.Width
            
            fraMulti.Left = picOuter(0).Left + picOuter(0).Width - fraMulti.Width
            
            picOuter(0).Visible = True
            AppWindowState = Me.WindowState
        End If
    End If
End Sub

Public Sub SetImageState()
    If picJPG.Width = maxX Or picJPG.Height = maxY Then
        rMX = picJPG.Width: rMY = picJPG.Height
        rSX = imgSize.Width: rSY = imgSize.Height
        dMTop = picJPG.Top: dMLeft = picJPG.Left
        dSTop = picJPG.Top: dSLeft = picJPG.Left
        iImageState = 1
    ElseIf picJPG.Width < maxX And picJPG.Height < maxY Then
        rSX = picJPG.Width: rSY = picJPG.Height
        dSTop = picJPG.Top: dSLeft = picJPG.Left
        Select Case rAsp
            Case Is = rFAsp
                rMX = maxX: rMY = maxY
                dMTop = dTop: dMLeft = dLeft
            Case Is > rFAsp 'X IS DETERMINING FACTOR'
                rMX = maxX: rMY = picJPG.Height * (maxX / picJPG.Width)
                dMLeft = dLeft: dMTop = dTop + (maxY - rMY) / 2
            Case Is < rFAsp 'Y IS DETERMINING FACTOR'
                rMY = maxY: rMX = picJPG.Width * (maxY / picJPG.Height)
                dMTop = dTop: dMLeft = dLeft + (maxX - rMX) / 2
        End Select
        iImageState = 0
    End If
    
End Sub

Public Sub GetGraphicList()
    Dim strSelect As String, sList As String, sStat As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sDesc As String, sCNode As String, sDNode As String, sTNode As String, _
                sGNode As String, sSNode As String, sFNode As String, SHNode As String, _
                sPNode As String
    Dim iType As Integer
'''    Dim sGStatus(10 To 30) As String
'''    Dim iGStatus(0 To 30) As Integer
    
'''    '///// FILE STATUS VARIABLES \\\\\
'''    sGStatus(10) = "INTERNAL"
'''    sGStatus(20) = "CLIENT DRAFT"
'''    sGStatus(30) = "APPROVED"
    
'''    iGStatus(0) = 10
'''    iGStatus(10) = 7
'''    iGStatus(20) = 8
'''    iGStatus(30) = 9
    
    
    tvwGraphics.Visible = False
    tvwGraphics.Nodes.Clear
    tvwGraphics.ImageList = ImageList1
'''    sCNode = "": sDNode = "": sTNode = "": sGNode = "": sSNode = ""
    
    sList = "10, 20, 30"

        
    ''FIRST CHECK FOR CLIENT FOLDERS'' ''HARD CODED TO 27812 FOR NOW''
'    strSelect = "SELECT DISTINCT GM.AN8_CUNO, " & _
'                "GM.FLR_ID, GF.FLRDESC, GF.CLIENTRESTRICT_FLAG AS FLAG " & _
'                "FROM GFX_MASTER GM, GFX_FOLDER GF " & _
'                "Where GM.AN8_CUNO = 40579 " & _
'                "AND GM.FLR_ID > 0 " & _
'                "AND GM.GSTATUS IN (10, 20, 30) " & _
'                "AND GM.FLR_ID  = GF.FLR_ID " & _
'                "ORDER BY GF.FLRDESC"
                
    strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC, " & _
                "GF.FLRLEVEL, GF.FLRPARENT, GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                "FROM ANNOTATOR.GFX_FOLDER GF " & _
                "Where GF.AN8_CUNO = 40579 " & _
                "AND GF.FLR_ID > 0 " & _
                "ORDER BY GF.FLRLEVEL, UPPER(GF.FLRDESC)"
                
'    strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC, " & _
'                "GF.FLRLEVEL, GF.FLRPARENT, GF.CLIENTRESTRICT_FLAG AS FLAG " & _
'                "FROM GFX_FOLDER GF " & _
'                "Where GF.AN8_CUNO = 40579 " & _
'                "AND GF.FLR_ID > 0 " & _
'                "ORDER BY GF.FLRLEVEL, UPPER(GF.FLRDESC)"
                
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        iType = 1
'''        SHNode = "h0"
'''        sDesc = "Client Folders"
'''        Set nodX = tvwGraphics(3).Nodes.Add(, , SHNode, sDesc, iType)
        Do While Not rst.EOF
            Debug.Print "Flag " & rst.Fields("FLAG")
'''            iType = rst.Fields("FLAG")
'''            If iType > 1 Then iType = 2
            sFNode = "f" & rst.Fields("FLR_ID")
            sDesc = Trim(rst.Fields("FLRDESC"))
            If rst.Fields("FLRLEVEL") = "A" Then
                Set nodX = tvwGraphics.Nodes.Add(, , sFNode, sDesc, iType)
            Else
                If IsNull(rst.Fields("FLRPARENT")) Then
                    Set nodX = tvwGraphics.Nodes.Add(, , sFNode, sDesc, iType)
                Else
                    sPNode = "f" & Trim(Mid(rst.Fields("FLRPARENT"), 2))
                    Set nodX = tvwGraphics.Nodes.Add(sPNode, tvwChild, sFNode, sDesc, iType)
                End If
            End If
            
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    
    
'''    If bPerm(29) Then
'''        If iView = 0 Then
'''            strSelect = "SELECT GM.AN8_CUNO, C.ABALPH, GM.GID, GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
'''                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GSTATUS IN (" & sList & ") " & _
'''                        "AND GM.AN8_CUNO = C.ABAN8 " & _
'''                        "ORDER BY Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
''''''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
'''        Else
'''            strSelect = "SELECT DISTINCT " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1 " & _
'''                        "FROM GFX_MASTER GM " & _
'''                        "Where GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GSTATUS IN (" & sList & ") " & _
'''                        "ORDER BY Y4, M2"
'''
''''''            strSelect = "SELECT DISTINCT GM.AN8_CUNO, C.ABALPH, GM.GTYPE, GM.GSTATUS, " & _
''''''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
''''''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
''''''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
''''''                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
''''''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
''''''                        "AND GM.GSTATUS IN (" & sList & ") " & _
''''''                        "AND GM.AN8_CUNO = C.ABAN8 " & _
''''''                        "ORDER BY Y4, M2, GM.GTYPE"
'''''''''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
'''        End If
'''    Else
'''        If iView = 0 Then
'''            strSelect = "SELECT GM.AN8_CUNO, C.ABALPH, GM.GID, GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
'''                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GSTATUS IN (" & sList & ") " & _
'''                        "AND GM.GTYPE <> 3 " & _
'''                        "AND GM.AN8_CUNO = C.ABAN8 " & _
'''                        "ORDER BY Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
''''''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
'''        Else
'''            strSelect = "SELECT DISTINCT GM.AN8_CUNO, C.ABALPH, GM.GTYPE, GM.GSTATUS, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
'''                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GSTATUS IN (" & sList & ") " & _
'''                        "AND GM.GTYPE <> 3 " & _
'''                        "AND GM.AN8_CUNO = C.ABAN8 " & _
'''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
''''''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
'''        End If
'''    End If
'''
'''
'''
'''    Set rst = Conn.Execute(strSelect)
'''    Do While Not rst.EOF
''''''        Select Case Len(rst.Fields("GSTATUS"))
''''''            Case 1
''''''                sStat = "0" & CStr(rst.Fields("GSTATUS"))
''''''            Case 2
''''''                sStat = CStr(rst.Fields("GSTATUS"))
''''''        End Select
''''''        If sSNode <> "S" & sStat Then
''''''            sDesc = sGStatus(rst.Fields("GSTATUS")) & " Files"
''''''            sSNode = "S" & sStat
''''''            Set nodX = tvwGraphics(3).Nodes.Add(, , sSNode, sDesc, iGStatus(rst.Fields("GSTATUS")))
''''''        End If
'''
''''''        If sDNode <> "D" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") Then
'''        If sDNode <> "D" & rst.Fields("M2") & rst.Fields("Y4") Then
'''            sDesc = "Posted:  " & UCase(Trim(rst.Fields("M1"))) & " " & Trim(rst.Fields("Y4"))
'''            sDNode = "D" & rst.Fields("M2") & rst.Fields("Y4")
''''''            sDNode = "D" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4")
'''            Set nodX = tvwGraphics(3).Nodes.Add(, , sDNode, sDesc, 5)
'''        End If
'''
''''''        iType = rst.Fields("GTYPE")
'''''''''        If sTNode <> "T" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") & _
'''''''''                        "-" & rst.Fields("GTYPE") Then
''''''        If sTNode <> "T" & rst.Fields("M2") & rst.Fields("Y4") & _
''''''                        "-" & rst.Fields("GTYPE") Then
''''''            sDesc = GfxType(rst.Fields("GTYPE"))
'''''''''            sTNode = "T" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") & _
'''''''''                        "-" & rst.Fields("GTYPE")
''''''            sTNode = "T" & rst.Fields("M2") & rst.Fields("Y4") & _
''''''                        "-" & rst.Fields("GTYPE")
''''''            Set nodX = tvwGraphics(3).Nodes.Add(sDNode, tvwChild, sTNode, sDesc, iType)
''''''        End If
'''
''''''        If sCNode <> "C" & rst.Fields("AN8_CUNO") Then
''''''            sDesc = UCase(Trim(rst.Fields("ABALPH")))
''''''            sCNode = "C" & rst.Fields("AN8_CUNO")
''''''            Set nodX = tvwGraphics(3).Nodes.Add(, , sCNode, sDesc, 5)
''''''        End If
''''''
''''''        If sDNode <> "D" & rst.Fields("M2") & rst.Fields("Y4") Then
''''''            sDesc = UCase(Trim(rst.Fields("M1"))) & " " & Trim(rst.Fields("Y4"))
''''''            sDNode = "D" & rst.Fields("M2") & rst.Fields("Y4")
''''''            Set nodX = tvwGraphics(3).Nodes.Add(sCNode, tvwChild, sDNode, sDesc, 5)
''''''        End If
'''
''''''        iType = rst.Fields("GTYPE")
''''''        If sTNode <> "T" & rst.Fields("GTYPE") & rst.Fields("M2") & rst.Fields("Y4") Then
''''''            sDesc = GfxType(rst.Fields("GTYPE"))
''''''            sTNode = "T" & rst.Fields("GTYPE") & rst.Fields("M2") & rst.Fields("Y4")
''''''            Set nodX = tvwGraphics(3).Nodes.Add(sDNode, tvwChild, sTNode, sDesc, iType)
''''''        End If
'''
''''''        If iView = 0 Then
''''''            sGNode = "g" & rst.Fields("GID")
''''''            sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
''''''            Set nodX = tvwGraphics(3).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, iType) ''' iGStatus(rst.Fields("GSTATUS")))
''''''        End If
''''''        iType = rst.Fields("GTYPE")
''''''        Set nodX = tvwGraphics(3).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, iGStatus(rst.Fields("GSTATUS"))) '' iType)
'''        rst.MoveNext
'''    Loop
'''    rst.Close
'''    Set rst = Nothing
    tvwGraphics.Visible = True
End Sub


Private Sub hsc1_Change(Index As Integer)
    picInner(Index).Left = CLng(hsc1(Index).Value) * (-100)
End Sub

Private Sub hsc1_Scroll(Index As Integer)
    picInner(Index).Left = CLng(hsc1(Index).Value) * (-100)
End Sub

Private Sub imgPopClose_Click()
    picMulti.Visible = False
End Sub

Private Sub imgResult_Click()
    picResult.Visible = False
End Sub

Private Sub imx0_Click(Index As Integer)
    Dim sHDR As String
    
    iIMXIndex = Index
    
    picDirs.Visible = False
    bDirsOpen = False
    imgDirs.ToolTipText = "Click to Open File Index..."
'''    Set imgDirs.Picture = imlDirs.ListImages(1).Picture
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    sHDR = GetHeader(tvwGraphics.SelectedItem)
'''    Call LoadGraphic(0, imx0(Index).Tag, lbl0(Index).Caption, sHDR)
    Call LoadGraphic(0, imx0(Index).Tag, lbl0(Index).Tag, sHDR)
    Call WhatSupDoc(CLng(imx0(Index).Tag))
End Sub

Private Sub lblBatch_Click(Index As Integer)
    Select Case Index
        Case 0 ''PREVIOUS''
            iListStart = iListStart - (iCols * iRows) ''20
            If iListStart < 1 Then iListStart = 1
            
        Case 1 ''NEXT''
            iListStart = iListStart + (iCols * iRows) ''20
            If iListStart > iGFXCount Then iListStart = (Int((iGFXCount - 1) / (iCols * iRows)) * (iCols * iRows)) + 1
            
        Case 2 ''FIRST''
            iListStart = 1
            
        Case 3 ''LAST''
            iListStart = (Int((iGFXCount - 1) / (iCols * iRows)) * (iCols * iRows)) + 1
            
    End Select
    
    Me.MousePointer = 11
    picWait.Visible = True
    picWait.Refresh
    
'''    For i = 0 To 9
'''        If i = Index Then
'''            lblCount(i).ForeColor = vbRed
'''            lblCount(i).Refresh
''''''            lblCount(i).FontBold = True
'''        Else
'''            If lblCount(i).Visible Then
'''                lblCount(i).ForeColor = vbBlack
''''''                lblCount(i).FontBold = False
'''            End If
'''        End If
'''    Next i
    TNode = tvwGraphics.SelectedItem.Key
    picInner(0).Visible = False
    Call GetGraphics(0, CurrSelect, iListStart, TNode)
    picInner(0).Visible = True
    
    picWait.Visible = False
    Me.MousePointer = 0
    
    Call ResetBatch(iListStart, iGFXCount, 0)
    
'''    If iListStart - 20 < 1 Then
'''        lblBatch(0).Enabled = False
'''        lblBatch(2).Enabled = False
'''    Else
'''        lblBatch(0).Enabled = True
'''        lblBatch(2).Enabled = True
'''    End If
'''    If iListStart + 20 > iGFXCount Then _
'''        lblBatch(1).Enabled = False
'''        lblBatch(3).Enabled = False
'''    Else
'''        lblBatch(1).Enabled = True
'''        lblBatch(3).Enabled = True
'''    End If
    
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblClose.ForeColor = vbWhite
End Sub

Private Sub lblDownload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xval As Single, Yval As Single
    
'    lblDownload.BackColor = lGeo_Back
    lblDownload.ForeColor = vbRed
    lblEmail.ForeColor = vbButtonText
    
    Xval = picDirs.Left + fraMulti.Left + lblDownload.Left
    Yval = picDirs.Top + fraMulti.Top + lblDownload.Top + lblDownload.Height
    Me.PopupMenu mnuDownloadMulti, , Xval, Yval
    
End Sub

Private Sub lblDownload_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblDownload.ForeColor = vbButtonText
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xval As Single, Yval As Single
    
    
'    lblDownload.BackColor = lGeo_Back
    lblDownload.ForeColor = vbButtonText
    lblEmail.ForeColor = vbRed
    
    Xval = picDirs.Left + fraMulti.Left + lblEmail.Left
    Yval = picDirs.Top + fraMulti.Top + lblEmail.Top + lblEmail.Height
    Me.PopupMenu mnuEmailMulti, , Xval, Yval
End Sub

Private Sub lblHelp_Click()
    lblHelp.ForeColor = vbWhite '' lColor
    frmHelp.Show 1
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = vbWhite
End Sub

'''Private Sub lblCount_Click(Index As Integer)
'''    Dim i As Integer
'''
'''    Me.MousePointer = 11
'''    picWait.Visible = True
'''    picWait.Refresh
'''
'''    For i = 0 To 9
'''        If i = Index Then
'''            lblCount(i).ForeColor = vbRed
'''            lblCount(i).Refresh
''''''            lblCount(i).FontBold = True
'''        Else
'''            If lblCount(i).Visible Then
'''                lblCount(i).ForeColor = vbBlack
''''''                lblCount(i).FontBold = False
'''            End If
'''        End If
'''    Next i
'''    TNode = tvwGraphics.SelectedItem.key
'''    picInner(0).Visible = False
'''    Call GetGraphics(Index, CurrSelect, lblCount(Index).Caption, TNode)
'''    picInner(0).Visible = True
'''
'''    picWait.Visible = False
'''    Me.MousePointer = 0
'''
'''End Sub

Private Sub lblList_Click()
    frmGfxList.PassFrom = Me.Name
    frmGfxList.PassSQL = CurrSelect
    frmGfxList.PassSize = (iRows * iCols)
    frmGfxList.Show 1, Me
End Sub

Private Sub lblMess_Click()
    Call picMulti_Click
End Sub

Private Sub lblPreview_Click()
''    frmViewer.PassFile = CurrFile
''    frmViewer.PassLeft = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + picViewer.Left + 240 ''imxViewer.Left
''    frmViewer.PassTop = Me.Top + ((Me.Height - Me.ScaleHeight) - ((Me.Width - Me.ScaleWidth) / 2)) _
''                + picViewer.Top + 600 ''imxViewer.Top
''    frmViewer.Show 1, Me
End Sub

Private Sub lblResult_DblClick()
    picResult.Top = picMenu2.Top + picMenu2.Height
    picResult.Left = 0 '' Me.ScaleWidth - 240 - picResult.Width
End Sub

Private Sub lblResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "xStr=" & X & "    yStr=" & Y
    xStr = X: yStr = Y
    picResult.Drag 1
End Sub

'''Private Sub lblViewAll_Click(Index As Integer)
''''    Dim i As Integer
''''    Me.MousePointer = 11
''''    lblViewAll(Index).ForeColor = vbRed
''''    lblViewAll(Index).Refresh
''''    For i = 0 To 9
''''        If lblCount(i).Visible Then
''''            lblCount(i).ForeColor = vbBlack
''''            lblCount(i).Refresh
''''        End If
''''    Next i
''''    TNode = tvwGraphics.SelectedItem.key
''''    picInner(0).Visible = False
''''    Call GetGraphics(99, CurrSelect, "1-1000", TNode)
''''    picInner(0).Visible = True
''''    Me.MousePointer = 0
'''
'''    frmGfxList.PassFrom = Me.Name
'''    frmGfxList.PassSQL = CurrSelect
'''    frmGfxList.Show 1, Me
'''End Sub

Private Sub mnuKeywordEditor_Click()
     
    frmKeywordEdit.PassGID = lGID
    frmKeywordEdit.Show 1, Me
'''    MsgBox "This would launch the Keyword Editor", vbInformation, "Coming Soon..."
End Sub

Private Sub lstResult_Click()
    Call LoadGraphic(0, lstResult.ItemData(lstResult.ListIndex), _
                lstResult.List(lstResult.ListIndex), " Search Result...")
End Sub

Private Sub mnuDownload_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, _
                sChk As String, sPath As String, sFile As String
    Dim rst As ADODB.Recordset
    
    
    pDownloadPath = ""
    frmBrowse.PassFrom = Me.Name
    frmBrowse.Show 1, Me
    
'''    If shlShell Is Nothing Then
'''        Set shlShell = New Shell32.Shell
'''    End If
'''
'''    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
'''                "Select Folder to download Graphic into:", _
'''                BIF_RETURNONLYFSDIRS)
'''
'''    If shlFolder Is Nothing Then

    If pDownloadPath = "" Then
        Exit Sub
    Else
        Me.Refresh
        Screen.MousePointer = 11
        
        On Error GoTo BadFile
        sFolder = pDownloadPath '' shlFolder.Items.Item.Path
        
        If UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        On Error GoTo ErrorTrap
        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                    "FROM " & GFXMas & " " & _
                    "WHERE GID = " & lGID
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sFile = Trim(rst.Fields("GPATH"))
            sPath = sFolder & "\" & Trim(rst.Fields("GDESC")) & _
                        "." & Trim(rst.Fields("GFORMAT"))
        Else
            rst.Close: Set rst = Nothing
            Screen.MousePointer = 0
            MsgBox "Error:  File Not Found", vbExclamation, "File Not Copied..."
            Exit Sub
        End If
        rst.Close: Set rst = Nothing
        
        If FileLen(sFile) > 500000 Then
            frmDownloadProgress.PassSRCFILE = sFile
            frmDownloadProgress.PassDESFILE = sPath
            frmDownloadProgress.Show 1, Me
        Else
            FileCopy sFile, sPath
            Screen.MousePointer = 0
            MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
        End If
        
        Screen.MousePointer = 0
'''        MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
    End If
    
Exit Sub
ErrorTrap:
    rst.Close: Set rst = Nothing
    Screen.MousePointer = 0
    MsgBox "Error:  " & Err.Description, vbExclamation, "File Not Copied..."
Exit Sub
BadFile:
    Screen.MousePointer = 0
    MsgBox "The location you chose is not valid.  If you chose " & _
                "'DESKTOP', be aware the desktop in this interface is the " & _
                "Citrix Servers Desktop and you do not have rights to " & _
                "download files there." & vbNewLine & vbNewLine & _
                "Please, choose another folder.", vbExclamation, "Invalid Location..."
    Err.Clear
End Sub

Private Sub mnuDownloadAdd_Click()
    Dim sMess As String, strInsert As String, strSelect As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lDLID As Long
    
'''    sMess = "Digital Image Library file" & vbNewLine
'''    sMess = sMess & vbNewLine & _
'''                "Location = " & lblWelcome.Caption
'''    sMess = sMess & vbNewLine & _
'''                "Source Path = " & CurrFile
'''    sMess = sMess & vbNewLine & _
'''                "File Name = " & lbl0(iIMXIndex).Tag & Right(imx0(iIMXIndex).FileName, 4)
'''    sMess = sMess & vbNewLine & _
'''                "File Size = " & Format(FileLen(CurrFile) / 1000, "#,##0") & "KB"
'''
'''    MsgBox sMess, vbInformation, "Add to Cart..."
    
    ''CHECK IF GID EXISTS IN USER'S CART''
    strSelect = "SELECT DLID FROM ANNOTATOR.ANO_DOWNLOAD " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND FILE_TYPE = 'DIL' " & _
                "AND FILE_ID = " & lGID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        rst.Close: Set rst = Nothing
        MsgBox "This file already exists in your Download Cart", vbExclamation, "FYI..."
        Exit Sub
    End If
    rst.Close: Set rst = Nothing
    
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    lDLID = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing
    
    strInsert = "INSERT INTO ANNOTATOR.ANO_DOWNLOAD " & _
                "(DLID, USER_SEQ_ID, FILE_TYPE, DLSTATUS, " & _
                "FILE_ID, SOURCE_PATH, FILE_NAME, SOURCE_DESC, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                "VALUES " & _
                "(" & lDLID & ", " & UserID & ", 'DIL', 1, " & lGID & ", '" & DeGlitch(CurrFile) & "', " & _
                "'" & lbl0(iIMXIndex).Tag & Right(imx0(iIMXIndex).FileName, 4) & "', " & _
                "'" & DeGlitch(lblWelcome.Caption) & "', " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
    
    lblDLCnt.Caption = CInt(lblDLCnt.Caption) + 1
    picDLCart.Visible = True
End Sub

Private Sub mnuDownloadMode_Click()
    Dim i As Integer
    Dim bFound As Boolean
    
    ''MAKE CERTAIN NOT IN EMAIL MODE''
    mnuEmailMode.Checked = 0
    lblEmail.ForeColor = vbButtonText
    bEMode = False
    bFound = False
    
    Select Case mnuDownloadMode.Checked
        Case Is = True ''END DOWNLOAD MODE''
            bDMode = False
            mnuDownloadMode.Checked = False
            lblDownload.ForeColor = vbButtonText
            For i = 0 To imx0.Count - 1
                imx0(i).Enabled = True
                chkMulti(i).Visible = False
                chkMulti(i).Value = 0
            Next i
            mnuDownloadSels.Enabled = False
            mnuDownloadSelsAdd.Visible = False
            mnuDownloadSels2.Visible = False
            mnuDownloadSelsAdd2.Visible = False
            lblDownload.ForeColor = vbButtonText
            picMulti.Visible = False
            
            
        Case Is = False ''START DOWNLOAD MODE''
            bDMode = True
            bEMode = False
            mnuDownloadMode.Checked = True
            For i = 0 To imx0.Count - 1
                If imx0(i).Visible Then
                    imx0(i).Enabled = False
                    chkMulti(i).Visible = True
                    chkMulti(i).ZOrder
                    If Not bFound And chkMulti(i).Value = 1 Then bFound = True
                End If
            Next i
            mnuDownloadSels.Visible = True
            mnuDownloadSelsAdd.Visible = True
            mnuDownloadSels2.Visible = True
            mnuDownloadSelsAdd2.Visible = True
            mnuDownloadSels.Enabled = bFound
            mnuDownloadSelsAdd.Enabled = bFound
            mnuDownloadSels2.Enabled = bFound
            mnuDownloadSelsAdd2.Enabled = bFound
            mnuEmailSels.Enabled = False
            mnuEmailSels2.Visible = False
            lblDownload.ForeColor = vbRed
            
            
            lblMess.Caption = "Checkboxes have been " & vbNewLine & _
                        "placed adjacent to each of the Thumbnail " & _
                        "Images in this page." & vbNewLine & vbNewLine & _
                        "Please, check the Images you would like to Download, " & _
                        "then return to the Download Popup to execute."
            picMulti.Height = (lblMess.Top * 2) + lblMess.Height
            picMulti.Visible = True
            picMulti.SetFocus
            
'''            MsgBox "Checkboxes have been placed adjacent to each of the Thumbnail " & _
'''                        "Images in this page." & vbNewLine & vbNewLine & _
'''                        "Please, check the Images you would like to Download, " & _
'''                        "then return to the Download Popup to execute.", _
'''                        vbInformation, "Entering Download Mode..."
    End Select
End Sub

Private Sub mnuDownloadSels_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, sChk As String, sPath As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    
    pDownloadPath = ""
    frmBrowse.PassFrom = Me.Name
    frmBrowse.Show 1, Me
    
'''    If shlShell Is Nothing Then
'''        Set shlShell = New Shell32.Shell
'''    End If
'''
'''    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
'''                "Select Folder to download Image File(s) into:", _
'''                BIF_RETURNONLYFSDIRS)
'''
'''    If shlFolder Is Nothing Then
    
    If pDownloadPath = "" Then
        Exit Sub
    Else
        Me.Refresh
        Screen.MousePointer = 11
        
        On Error GoTo BadFile
        sFolder = pDownloadPath ''shlFolder.Items.Item.Path
        
        If UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        On Error GoTo ErrorTrap
        
        For i = 0 To chkMulti.Count - 1
            If chkMulti(i).Value = 1 Then
                strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                            "FROM " & GFXMas & " " & _
                            "WHERE GID = " & imx0(i).Tag
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    sPath = sFolder & "\" & Trim(rst.Fields("GDESC")) & _
                                "." & Trim(rst.Fields("GFORMAT"))
                Else
                    rst.Close: Set rst = Nothing
                    Screen.MousePointer = 0
                    MsgBox "Error:  File Not Found", vbExclamation, "File Not Copied..."
                    Exit Sub
                End If
                
                FileCopy Trim(rst.Fields("GPATH")), sPath
                rst.Close: Set rst = Nothing
            End If
        Next i
        
        Screen.MousePointer = 0
        MsgBox "File(s) Copied to " & sFolder, vbInformation, "File Download Successful..."
    End If
    
Exit Sub
ErrorTrap:
    rst.Close: Set rst = Nothing
    Screen.MousePointer = 0
    MsgBox "Error:  " & Err.Description, vbExclamation, "File(s) Not Copied..."
Exit Sub
BadFile:
    Screen.MousePointer = 0
    MsgBox "The location you chose is not valid.  If you chose " & _
                "'DESKTOP', be aware the desktop in this interface is the " & _
                "Citrix Servers Desktop and you do not have rights to " & _
                "download files there." & vbNewLine & vbNewLine & _
                "Please, choose another folder.", vbExclamation, "Invalid Location..."
    Err.Clear
End Sub

Private Sub mnuDownloadSels2_Click()
    Call mnuDownloadSels_Click
End Sub


Private Sub mnuDownloadSelsAdd_Click()
    Dim i As Integer, iLoc As Integer, iCnt As Integer
    Dim sMess As String, sLoc As String, sPath As String, sFile As String
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lDLID As Long
    
    iCnt = 0
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Visible And chkMulti(i).Value = 1 Then
            ''CHECK IF GID EXISTS IN USER'S CART''
            strSelect = "SELECT DLID FROM ANNOTATOR.ANO_DOWNLOAD " & _
                        "WHERE USER_SEQ_ID = " & UserID & " " & _
                        "AND FILE_TYPE = 'DIL' " & _
                        "AND FILE_ID = " & imx0(i).Tag
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                rst.Close: Set rst = Nothing
                MsgBox "'" & lbl0(i).Tag & "' already exists in your Download Cart", _
                            vbExclamation, "Skipping Selection..."
                GoTo SkipIt
            End If
            rst.Close: Set rst = Nothing
            
            Err = 0: sLoc = ""
            On Error Resume Next
            sLoc = tvwGraphics.SelectedItem.Text
            If Err = 0 Then sLoc = tvwGraphics.SelectedItem.Parent.Text & " (" & sLoc & ")"
            If Err = 0 Then sLoc = tvwGraphics.SelectedItem.Parent.Parent.Text & " (" & sLoc & ")"
            If Err = 0 Then sLoc = tvwGraphics.SelectedItem.Parent.Parent.Parent.Text & " (" & sLoc & ")"
            If Err = 0 Then sLoc = tvwGraphics.SelectedItem.Parent.Parent.Parent.Parent.Text & " (" & sLoc & ")"
            If Err = 0 Then sLoc = tvwGraphics.SelectedItem.Parent.Parent.Parent.Parent.Parent.Text & " (" & sLoc & ")"
'''GotEm:
            On Error Resume Next
            Err = 0
            strSelect = "SELECT GPATH, GFORMAT FROM ANNOTATOR.GFX_MASTER WHERE GID = " & imx0(i).Tag
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sPath = Trim(rst.Fields("GPATH"))
                sFile = lbl0(i).Tag & "." & LCase(rst.Fields("GFORMAT"))
            Else
                rst.Close: Set rst = Nothing
                MsgBox "The source file for '" & lbl0(i).Tag & "' " & _
                            "could not be found", vbExclamation, "Skipping file..."
                GoTo SkipIt
            End If
            rst.Close: Set rst = Nothing
            
            
            
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lDLID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            
            strInsert = "INSERT INTO ANNOTATOR.ANO_DOWNLOAD " & _
                        "(DLID, USER_SEQ_ID, FILE_TYPE, DLSTATUS, FILE_ID, " & _
                        "SOURCE_PATH, FILE_NAME, SOURCE_DESC, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lDLID & ", " & UserID & ", 'DIL', 1, " & imx0(i).Tag & ", " & _
                        "'" & DeGlitch(sPath) & "', " & _
                        "'" & DeGlitch(sFile) & "', " & _
                        "'" & DeGlitch(sLoc) & "', " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            
            lblDLCnt.Caption = CInt(lblDLCnt.Caption) + 1
            picDLCart.Visible = True
            iCnt = iCnt + 1
        End If
SkipIt:
    Next i
    
    If iCnt = 1 Then
        MsgBox "A new file has been added to your Download Cart.  It is available " & _
                    "to download by clicking on the Cart icon above.", _
                    vbInformation, "New Download Cart file..."
    ElseIf iCnt <> 0 Then
        MsgBox "( " & iCnt & " ) new files have been added to your Download Cart.  They are available " & _
                    "to download by clicking on the Cart icon above.", _
                    vbInformation, "New Download Cart files..."
    End If
End Sub

Private Sub mnuDownloadSelsAdd2_Click()
    Call mnuDownloadSelsAdd_Click
End Sub

Private Sub mnuEmailMode_Click()
    Dim i As Integer
    Dim bFound As Boolean
    
    ''MAKE CERTAIN NOT IN DOWNLOAD MODE''
    mnuDownloadMode.Checked = 0
    lblDownload.ForeColor = vbButtonText
    bDMode = False
    bFound = False
    
    Select Case mnuEmailMode.Checked
        Case Is = True ''END EMAIL MODE''
            bEMode = False
            mnuEmailMode.Checked = False
            For i = 0 To imx0.Count - 1
                imx0(i).Enabled = True
                chkMulti(i).Visible = False
                chkMulti(i).Value = 0
            Next i
            mnuEmailSels.Enabled = False
            mnuEmailSels2.Visible = False
            lblEmail.ForeColor = vbButtonText
            
            
        Case Is = False ''START EMAIL MODE''
            bEMode = True
            bDMode = False
            mnuEmailMode.Checked = True
            For i = 0 To imx0.Count - 1
                If imx0(i).Visible Then
                    imx0(i).Enabled = False
                    chkMulti(i).Visible = True
                    chkMulti(i).ZOrder
                    If Not bFound And chkMulti(i).Value = 1 Then bFound = True
                End If
            Next i
            mnuEmailSels.Visible = True
            mnuEmailSels2.Visible = True
            mnuEmailSels.Enabled = bFound
            mnuEmailSels2.Enabled = bFound
            mnuDownloadSelsAdd.Enabled = False
            mnuDownloadSels.Enabled = False
            mnuDownloadSelsAdd2.Visible = False
            mnuDownloadSels2.Visible = False
            lblEmail.ForeColor = vbRed
            
            
            lblMess.Caption = "Checkboxes have been " & vbNewLine & _
                        "placed adjacent to each of the Thumbnail " & _
                        "Images in this page." & vbNewLine & vbNewLine & _
                        "Please, check the Images you would like to Email Copies of, " & _
                        "then return to the Email Copy Popup to execute."
            picMulti.Height = (lblMess.Top * 2) + lblMess.Height
            picMulti.Visible = True
            picMulti.SetFocus
                        
'''            MsgBox "Checkboxes have been placed adjacent to each of the Thumbnail " & _
'''                        "Images in this page." & vbNewLine & vbNewLine & _
'''                        "Please, check the Images you would like to Email Copies of, " & _
'''                        "then return to the Email Copy Popup to execute.", _
'''                        vbInformation, "Entering Email Copy Mode..."
    End Select

End Sub

Private Sub mnuEmailSel_Click()
    frmEmailFile.PassFrom = Me.Name & "-single"
    frmEmailFile.PassBCC = "00040579"
    frmEmailFile.PassFBCN = "GPJ Digital Image Library"
    frmEmailFile.PassHDR = GetHeader(tvwGraphics.SelectedItem)
    frmEmailFile.Show 1, Me
    Call ClearModes
End Sub

Private Sub mnuEmailSels_Click()
    frmEmailFile.PassFrom = Me.Name & "-multi"
    frmEmailFile.PassBCC = "00040579"
    frmEmailFile.PassFBCN = "GPJ Digital Image Library"
    frmEmailFile.PassHDR = GetHeader(tvwGraphics.SelectedItem)
    frmEmailFile.Show 1, Me
    Call ClearModes
End Sub

Private Sub mnuEmailSels2_Click()
    Call mnuEmailSels_Click
End Sub

Private Sub mnuGFXData_Click()
    Dim strSelect As String
    
    strSelect = "SELECT GM.* " & _
                "FROM " & GFXMas & " GM " & _
                "WHERE GM.GID > 0 " & _
                "AND GM.GID = " & lGID
    Call GetGFXData(strSelect, "msgbox")
End Sub

Private Sub mnuGPrint_Click()
    Dim pScaleX As Long, pScaleY As Long, lXStart As Long, lYStart As Long
    Dim pAspect As Single
    Dim sMsg As String, sDef As String
    Dim iFind As Integer
    
    
    On Error GoTo ErrorTrap
    Screen.MousePointer = 11
    sMsg = "You are attempting to print to " & Printer.DeviceName & "." & vbCr & vbCr & _
                "An error has occured.  The Annotator server might not have the " & _
                "required printer driver to print to this device.  Either temporarily set " & _
                "another printer as your default printer and try again, or contact the " & _
                "GPJ Help Desk to arrange for the correct driver to be installed.  " & _
                "(NOTE: Before placing a Help Desk Call, please know the make and " & _
                "model of your default printer.)"
    On Error Resume Next
    
    cdl1.Flags = cdlPDPrintSetup
    cdl1.ShowPrinter
    
    pAspect = imgSize.Width / imgSize.Height
    If pAspect >= 1 Then '''LANDSCAPE'''
'''                cdl1.Orientation = cdlLandscape
        Printer.Orientation = 2
        If pAspect > (10 / 6.75) Then
            pScaleX = 10 * 1440
            pScaleY = imgSize.Height / imgSize.Width * pScaleX
        Else
            pScaleY = 6.75 * 1440
            pScaleX = imgSize.Width / imgSize.Height * pScaleY
        End If
        lXStart = (Printer.ScaleWidth - pScaleX) / 2
        lYStart = ((Printer.ScaleHeight - pScaleY) / 2) + (0.5 * 1440)
        
    Else '''PORTRAIT'''
'''                cdl1.Orientation = cdlPortrait
        Printer.Orientation = 1
        If pAspect < (7 / 9.25) Then
            pScaleY = 9.25 * 1440
            pScaleX = imgSize.Width / imgSize.Height * pScaleY
        Else
            pScaleX = 7 * 1440
            pScaleY = imgSize.Height / imgSize.Width * pScaleX
        End If
        lXStart = ((Printer.ScaleWidth - pScaleX) / 2) + (0.25 * 1440)
        lYStart = ((Printer.ScaleHeight - pScaleY) / 2) + (0.5 * 1440)
    End If
'''            cdl1.ShowPrinter
'''            Printer.Orientation = cdl1.Orientation
    
'''            cdl1.PrinterDefault = True
    picPrint.Visible = True
    picPrint.Refresh
    Debug.Print Printer.DeviceName
    Printer.PaintPicture imgSize.Picture, lXStart, lYStart, pScaleX, pScaleY
        
    
    If Err Then
        MsgBox sMsg, vbExclamation, "Printer Error..."
        Err = 0
    Else
'''        Printer.CurrentY = 0.5 * Printer.TwipsPerPixelY
        Printer.CurrentY = 650
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.Print "George P. Johnson Company"
        Printer.FontSize = 8
        Printer.FontBold = False
        iFind = InStr(1, lblGraphic, "[A Redline")
        If iFind = 0 Then
            Printer.Print lblWelcome.Caption & vbNewLine & _
                        lblGraphic.Caption & vbNewLine & _
                        "Printed:  " & Format(Now, "mmmm d, yyyy")
        Else
            Printer.Print lblWelcome.Caption & vbNewLine & _
                        Trim(Left(lblGraphic.Caption, iFind - 1)) & vbNewLine & _
                        "Printed:  " & Format(Now, "mmmm d, yyyy")
        End If
        Printer.EndDoc
        
        Screen.MousePointer = 0
        
'        If Err Then
'            Screen.MousePointer = 0
'            MsgBox sMsg, vbExclamation, "Printer Error..."
'            Err = 0
'        Else
'            Screen.MousePointer = 0
'            MsgBox "Image printed to default printer:  " & Printer.DeviceName, vbInformation, "Print Sent..."
'        End If
'''        Printer.Orientation = 1
    End If
'''    If Index = 0 Then imgSize.Picture = LoadPicture()
    picPrint.Visible = False
'''    Screen.MousePointer = 0
Exit Sub
ErrorTrap:
    picPrint.Visible = False
    Screen.MousePointer = 0
    MsgBox "Error:  " & Err.Description, vbExclamation, "Printer Error Encountered..."
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub mnuMaxGraphic_Click()
    With picJPG
        .Visible = False
        .Width = rMX
        .Height = rMY
        .Top = dMTop
        .Left = dMLeft
        .PaintPicture imgSize.Picture, 0, 0, .Width, .Height
        .Visible = True
    End With
    mnuResizeGraphic.Enabled = True
    mnuMaxGraphic.Enabled = False
    lblResize.Caption = "Resize"
    iImageState = 1

End Sub

Private Sub mnuOptClearAll_Click()
    Dim i As Integer
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Visible Then chkMulti(i).Value = 0
    Next i
    Select Case bDMode
        Case True
            mnuDownloadSels.Enabled = False
            mnuDownloadSels2.Enabled = False
            mnuDownloadSelsAdd.Enabled = False
            mnuDownloadSelsAdd2.Enabled = False
    End Select
    Select Case bEMode
        Case True
            mnuEmailSels.Enabled = False
            mnuEmailSels2.Enabled = False
    End Select
End Sub

Private Sub mnuOptSelAll_Click()
    Dim i As Integer
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Visible Then chkMulti(i).Value = 1
    Next i
    Select Case bDMode
        Case True
            mnuDownloadSels.Enabled = True
            mnuDownloadSels2.Enabled = True
            mnuDownloadSelsAdd.Enabled = True
            mnuDownloadSelsAdd2.Enabled = True
    End Select
    Select Case bEMode
        Case True
            mnuEmailSels.Enabled = True
            mnuEmailSels2.Enabled = True
    End Select
End Sub

Private Sub mnuResizeGraphic_Click()
    With picJPG
        .Visible = False
        .Width = rSX
        .Height = rSY
        .Top = dSTop ''' dGTop + ((rY - rYO) / 2): .Left = dGLeft + ((rX - rXO) / 2)
        .Left = dSLeft
'''        .Left = (maxX - rXO) / 2
        .PaintPicture imgSize.Picture, 0, 0, .Width, .Height
        .Visible = True
    End With
    mnuResizeGraphic.Enabled = False
    mnuMaxGraphic.Enabled = True
    lblResize.Caption = "Maximize"
    iImageState = 0
End Sub

Private Sub mnuSendALink_Click()
    frmSendALink.PassBCC = 40579
    frmSendALink.PassFrom = "DIL"
    frmSendALink.PassGID = lGID
    frmSendALink.PassSub = "AnnoLink: DIL - " & _
            Trim(lblWelcome.Caption) & "  (" & Trim(lblGraphic.Caption) & ")"
    frmSendALink.Show 1, Me
End Sub

Private Sub optIconSize_Click(Index As Integer)
    Dim i As Integer, iCol As Integer, iRow As Integer
'''    Dim lX As Long, lY As Long
    
    On Error GoTo ErrorTrap
    Screen.MousePointer = 11
    
    iIconSize = Index
    
    picOuter(0).Visible = False
    Select Case iIconSize
        Case 0
            lIconX = IconSize_0x: lIconY = IconSize_0y
        Case 1
            lIconX = IconSize_1x: lIconY = IconSize_1y
        Case 2
            lIconX = IconSize_2x: lIconY = IconSize_2y
    End Select
    imx0(0).Width = lIconX: imx0(0).Height = lIconY
    
    If bFirst Then
        picOuter(0).Visible = True
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
    iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
    
    imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
    spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
    imageX = CLng((picInner(0).ScaleWidth - 240) / iCols) '' (imageY / 3) * 4
    spaceX = imageX ''CLng(picInner(0).ScaleWidth / iCols) '' imageX + 720 ''240
    
    For i = 0 To (iCols * iRows) - 1 ''19
        If i >= imx0.Count Then Load imx0(i)
        imx0(i).Width = lIconX: imx0(0).Height = lIconY
        iCol = Int(i / iRows): iRow = i Mod iRows
'''        imx0(i).Left = 480 + (iCol * spaceX)
        imx0(i).Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
        imx0(i).Top = 120 + (iRow * spaceY)
        imx0(i).Update = True
        imx0(i).Refresh
        
        If i >= lbl0.Count Then Load lbl0(i)
        
        If i >= chkMulti.Count Then Load chkMulti(i)
    Next i
        
    If bPopped Then Call tvwGraphics_NodeClick(tvwGraphics.SelectedItem)
    
    picOuter(0).Visible = True
    Screen.MousePointer = 0
Exit Sub
ErrorTrap:
    Screen.MousePointer = 0
End Sub

Private Sub picDLCart_Click()
    Call imgDLCart_Click
End Sub

Private Sub picGraphic_DragDrop(Source As Control, X As Single, Y As Single)
    Debug.Print "Dragging " & Source.Name
    Source.Drag 2
    Source.Move CSng(picJPG.Left + (X - xStr)), CSng(picJPG.Top + (Y - yStr))

End Sub

Private Sub picGraphic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuRightClick
End Sub

Private Sub picInner_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bDMode And Not bEMode Then Exit Sub
    picInner(Index).Cls
    xStart = X
    yStart = Y
    bMouseDown = True
End Sub

Private Sub picInner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMouseDown Then
        movelines X, Y
    End If
End Sub

Private Sub picInner_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bMouseDown Then Exit Sub
    
    bMouseDown = False
    picInner(Index).Line (xStart, yStart)-(xs, ys), , B
    Debug.Print "xStart = " & xStart & " : yStart = " & yStart
    Debug.Print "xEnd = " & X & " : yEnd = " & Y
    
    picInner(Index).Cls
    xs = 0: ys = 0
    
    Call GetWindowSels(xStart, yStart, X, Y)
    
    Dim bFound As Boolean
    Dim i As Integer
    
    bFound = False
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Value = 1 Then
            bFound = True
            Exit For
        End If
    Next i
    
    If bEMode Then
        mnuEmailSels.Enabled = bFound
        mnuEmailSels2.Enabled = bFound
    End If
    If bDMode Then
        mnuDownloadSels.Enabled = bFound
        mnuDownloadSelsAdd.Enabled = bFound
        mnuDownloadSels2.Enabled = bFound
        mnuDownloadSelsAdd2.Enabled = bFound
    End If
    
End Sub

Private Sub picMulti_Click()
    picMulti.Visible = False
End Sub

Private Sub picMulti_LostFocus()
    picMulti.Visible = False
End Sub

Private Sub tvwGraphics_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lFID As Long
    Dim strSelect As String
    Dim i As Integer
    
    picWait.Visible = True
    picWait.Refresh
    
    TNode = Node.Key
    
    Call ResetNodeIcons(TNode)
    
'    CurrParNode = Node.Parent.key
'    CurrParText = Node.Parent.Text
    lFID = Mid(Node.Key, 2)
    strSelect = "SELECT GM.GID, GM.GDESC, " & _
                "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                "FROM ANNOTATOR.GFX_MASTER GM " & _
                "WHERE GM.FLR_ID = " & lFID & " " & _
                "AND GM.GTYPE IN (" & sInType & ") " & _
                "ORDER BY GM.GDESC"
    CurrSelect = strSelect
    
    iListStart = 1
    Call GetGraphics(0, strSelect, iListStart, TNode)
    
    If bDMode Or bEMode Then ClearModes
    
    picWait.Visible = False
End Sub

Public Sub GetGraphics(CntIndex As Integer, strSelect As String, iStart As Integer, sNode As String)
    Dim rst As ADODB.Recordset
    Dim i1 As Integer, i2 As Integer, iDash As Integer, iCnt As Integer
    Dim i As Integer, iLock As Integer, iCol As Integer, iRow As Integer
    Dim imxCon As ImagXpress
    Dim lblCon As Label
    Dim chkCon As CheckBox
    Dim sFile As String
    
    
    '''iDash = InStr(1, sSpan, "-")
    i1 = iStart - 1 '''CInt(Left(sSpan, iDash - 1)) - 1
    i2 = i1 + (iCols * iRows - 1) '' 19 ''CInt(Mid(sSpan, iDash + 1)) - 1
    
    Set rst = Conn.Execute(strSelect)
    i = 0: iCnt = 0
'''''        picInner(Index).Visible = False
    
    Do While Not rst.EOF
        Do While iCnt < i1
            rst.MoveNext
            iCnt = iCnt + 1
        Loop
        iCol = Int(i / iRows): iRow = i Mod iRows
        
        If i >= imx0.Count Then Load imx0(i)
        Set imxCon = imx0(i)
        
        With imxCon
            .Width = lIconX
            .Height = lIconY
            .Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
'            .Left = 360 + (iCol * spaceX)
            .Top = 120 + (iRow * spaceY)
            .Update = False
            .PICThumbnail = iRes
            
            If UCase(Trim(rst.Fields("GFORMAT"))) = "JPG" _
                        Or UCase(Trim(rst.Fields("GFORMAT"))) = "BMP" Then
                sFile = sGPath & "Thumbs\thb_" & rst.Fields("GID") & ".jpg"
                If Dir(sFile, vbNormal) = "" Then ''OPEN FULL FILE''
                    Select Case FileLen(Trim(rst.Fields("GPATH")))
                        Case Is < 10000: .PICThumbnail = THUMB_None
                        Case Is < 25000: .PICThumbnail = THUMB_4
                        Case Is < 50000: .PICThumbnail = THUMB_16
                        Case Else: .PICThumbnail = THUMB_64
                    End Select
                    .FileName = Trim(rst.Fields("GPATH"))
                Else
                    Select Case FileLen(sFile)
                        Case Is < 10000: .PICThumbnail = THUMB_None
                        Case Is < 25000: .PICThumbnail = THUMB_4
                        Case Is < 50000: .PICThumbnail = THUMB_16
                        Case Else: .PICThumbnail = THUMB_64
                    End Select
                    .FileName = sFile
                End If
                
'''                .FileName = Trim(rst.Fields("GPATH"))
                
            Else
                .FileName = sGPath & Trim(rst.Fields("GFORMAT")) & ".bmp"
            End If
            
            
            
'''            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
'''                .FileName = sGPath & "Acrobatid.bmp"
'''            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "PP" Then
'''                .FileName = sGPath & "PowerPointID.bmp"
'''            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "MP" Then
'''                .FileName = sGPath & "WMPlayerID.bmp"
'''            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "AV" Then
'''                .FileName = sGPath & "WMPlayerID.bmp"
'''            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "MO" Then
'''                .FileName = sGPath & "QuickTimeID.bmp"
'''            Else
'''                .FileName = Trim(rst.Fields("GPATH"))
'''            End If
            
            
'''            If bApprover Then
'''                .ToolTipText = "Right-Click to Reset the Status of this Graphic File"
'''            Else
                .ToolTipText = Trim(rst.Fields("GDESC"))
'''            End If
            .Update = True
            .Buttonize 1, 1, 50
            .Visible = True
            .Refresh
            
            .Tag = CStr(rst.Fields("GID"))
            
        End With
        
        If i >= lbl0.Count Then Load lbl0(i)
        Set lblCon = lbl0(i)
                
        With lblCon
            If Len(Trim(rst.Fields("GDESC"))) > 24 Then
                .Caption = Left(Trim(rst.Fields("GDESC")), 24) & "..."
            Else
                .Caption = Trim(rst.Fields("GDESC"))
            End If
            .Tag = Trim(rst.Fields("GDESC"))
            .Left = imxCon.Left + (imxCon.Width / 2) - (lblCon.Width / 2)
'''            .Left = 120 + (iCol * spaceX) + ((imageX - .Width) / 2)
            .Top = imxCon.Top + imxCon.Height + 30 ''120 + imageY + (iRow * spaceY)
            If rst.Fields("GSTATUS") > 0 Then
                .BackColor = vbWindowBackground
            Else
                .BackColor = vbRed
'                lblInactive(0).Visible = True
            End If
            .Visible = True
        End With
        
        If i >= chkMulti.Count Then Load chkMulti(i)
        Set chkCon = chkMulti(i)
        chkCon.Left = imxCon.Left
        chkCon.Top = imxCon.Top
        chkCon.Visible = False
        chkCon.Value = False
        
        i = i + 1
        iCnt = iCnt + 1
        If iCnt > i2 Then
            Do While Not rst.EOF
                rst.MoveNext
                If Not rst.EOF Then iCnt = iCnt + 1
            Loop
            iCnt = iCnt '' - 1
            GoTo CountDone
        End If
        rst.MoveNext
    Loop
CountDone:
    rst.Close: Set rst = Nothing
    
    iGFXCount = iCnt
    If iGFXCount > 0 Then fraMulti.Visible = True Else fraMulti.Visible = False
    
'''    Call ResetCounts(iCnt, CntIndex)
    Call ResetBatch(iListStart, iGFXCount, CntIndex)
    Call ClearThumbnails0(i)
    
'''    picInner(0).Width = ((iCol + 1) * spaceX) + 240

'''    If picInner(0).Width < picOuter(0).ScaleWidth Then
'''        hsc1(0).Max = picInner(0).Width / 100
'''        hsc1(0).Visible = False
'''    Else
'''        hsc1(0).Max = (picInner(0).Width / 100) - (picOuter(0).ScaleWidth / 100)
'''        hsc1(0).Visible = True
'''    End If
'''    hsc1(0).value = 0 '''picOuter(1).ScaleWidth
'''    hsc1(0).LargeChange = picOuter(0).ScaleWidth / 100

'    picInner(0).Visible = True
    
    bPopped = True
End Sub

'''Public Sub ResetCounts(iCnt As Integer, CntIndex As Integer)
'''    Dim i As Integer, i1 As Integer, i2 As Integer, iInt As Integer, iTotal As Integer
'''
'''    iTotal = iCnt
'''    If iCnt > 200 Then iCnt = 200
'''
'''    For i = 0 To 9
'''        lblCount(i).Visible = False
'''        lblCount(i).Caption = (i * 20) + 1 & "-" & (i + 1) * 20
'''        lblCount(i).ForeColor = vbBlack
'''        lblCount(i).FontBold = False
'''    Next i
'''    If CntIndex = 99 Then
'''        lblViewAll(0).ForeColor = vbRed
'''        lblViewAll(0).Visible = True
'''    Else
'''        lblViewAll(0).ForeColor = vbBlack
'''    End If
'''    iInt = Int((iCnt - 1) / 20)
'''    If iInt > 0 Then
'''        If CntIndex <> 99 Then lblCount(CntIndex).ForeColor = vbRed
'''        For i = 0 To iInt
'''            If i = iInt Then
'''                If iTotal > 200 Then
'''                    lblCount(i).Caption = (i * 20) + 1 & "-" & iTotal
'''                Else
'''                    lblCount(i).Caption = (i * 20) + 1 & "-" & iCnt
'''                End If
'''            End If
'''            lblCount(i).Visible = True
'''        Next i
'''        lblViewAll(0).Left = lblCount(iInt).Left + lblCount(iInt).Width + 300
'''        lblViewAll(0).Top = lblCount(iInt).Top
'''    End If
'''    If iInt > 1 Then lblViewAll(0).Visible = True Else lblViewAll(0).Visible = False
'''
'''End Sub

Public Sub ClearThumbnails0(iStart As Integer)
    Dim i As Integer
    For i = iStart To imx0.Count - 1
        imx0(i).Visible = False
        imx0(i).FileName = ""
        lbl0(i).Visible = False
    Next i
End Sub

Public Sub LoadGraphic(Index As Integer, NodeKey As String, NodeText As String, _
            WelcomeText As String)
    Dim strSelect As String, strInsert As String, strUpdate As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim i As Integer, iLock As Integer, iCol As Integer, iRow As Integer, tIndex As Integer
    Dim sGStatus(0 To 30) As String
    
    cmdRedline.Visible = False
    web1.Visible = False
    
    sGStatus(0) = "DE-ACTIVATED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    Debug.Print NodeKey & " - YOU GOT ONE!"

    strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, " & _
                "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                "FROM " & GFXMas & " GM " & _
                "WHERE GM.GID = " & CLng(NodeKey)
    sTable = "ANNOTATOR.GFX_MASTER"
    
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If Dir(Trim(rst.Fields("GPATH")), vbNormal) <> "" Then
            If Index >= 10 Then tIndex = Index - 10 Else tIndex = Index
            picDirs.Visible = False ''.Visible = False
            bDirsOpen = False
            imgDirs.ToolTipText = "Click to Open File Index..."
'''            Set imgDirs.Picture = imlDirs.ListImages(1).Picture
            
            lGID = rst.Fields("GID")
            sCGDesc = Trim(rst.Fields("GDESC"))
            iCurrGType = rst.Fields("GTYPE")
            
            CurrFile = rst.Fields("GPATH")
            
'''            lblStatus.Caption = "STATUS:  " & sGStatus(rst.Fields("GSTATUS")) & " (Last Status Update " & _
'''                        Trim(rst.Fields("STATDATE")) & ")"
            lblStatus.Caption = "Posting Date:  " & Trim(rst.Fields("STATDATE"))
            lblStatus.Visible = True
            
            
            '///// TIME TO LOAD THE GRAPHIC \\\\\
            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
'                Call LoadThePDF(CurrFile)
            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "PP" Then
                picJPG.Visible = False
                picViewer.Visible = False
                imgMenu.Visible = True
                lblMenu.Visible = True
                mnuResizeGraphic.Visible = False
                mnuMaxGraphic.Visible = False
                lblResize.Enabled = False
                lblFullSize.Enabled = False
                picMenu2.Visible = True
                
                mnuGPrint.Enabled = False
                
                web1.Navigate2 CurrFile
'''                web1.Visible = True
'                cmdRedline.Visible = True
                
                
            ElseIf Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "AV" _
                        Or Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "MP" _
                        Or Left(UCase(Trim(rst.Fields("GFORMAT"))), 2) = "MO" _
                        Then
                
                picJPG.Visible = False
                
                imxViewer.FileName = sGPath & Trim(rst.Fields("GFORMAT")) & ".bmp"
'''                imxViewer.FileName = imx0(iIMXIndex).FileName
                picViewer.Visible = True
                lblName.Caption = sCGDesc & "." & LCase(Trim(rst.Fields("GFORMAT")))
                lblSize.Caption = "File Size: " & Format(FileLen(CurrFile) / 1000, "#,##0") & " KB"
                
                imgMenu.Visible = True
                lblMenu.Visible = True
                mnuResizeGraphic.Visible = False
                mnuMaxGraphic.Visible = False
                lblResize.Enabled = False
                lblFullSize.Enabled = False
                picMenu2.Visible = True
                
                mnuGPrint.Enabled = False
                
''''''                MsgBox "Load PowerPoint File"
'''                frmHTMLViewer.PassFile = CurrFile
'''                frmHTMLViewer.PassFrom = Me.Name
'''                frmHTMLViewer.PassHDR = sCGDesc & "." & LCase(Trim(rst.Fields("GFORMAT")))
'''                frmHTMLViewer.PassDFile = sCGDesc & "." & LCase(Trim(rst.Fields("GFORMAT")))
'''                frmHTMLViewer.PassGID = lGID
'''                frmHTMLViewer.Show 1, Me
            Else
                picViewer.Visible = False
                Call LoadThePicture(CurrFile)
                
                mnuGPrint.Enabled = True
            End If
            
            lblWelcome.Caption = WelcomeText
            lblGraphic.Caption = "Image:  " & NodeText
            
        End If
    Else
        rst.Close
        Set rst = Nothing
        MsgBox "Graphic not found", vbExclamation, "Sorry..."
        GoTo NoGraphic
    End If
    rst.Close: Set rst = Nothing
    
'''        Select Case Index
'''            Case 0
'''                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & cboSHYR(0).Text & " " & cboSHCD.Text
'''            Case 1
'''                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & cboSHYR(1).Text & " " & NodeParText
'''            Case 2
'''                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & NodeParText
'''            Case 3
'''                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & tvwGraphics(3).Nodes(NodeParKey).Parent.Text
'''            Case 10
'''                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text & " " & NodeParText & ":  " & _
'''                            cboSHYR(0).Text & " " & cboSHCD.Text
'''            Case 11
'''                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text & " " & GfxType(iCurrGType) & ":  " & _
'''                            cboSHYR(1).Text '''& " " & tvwGraphics(1).Nodes(NodeParKey).Parent.Text
'''            Case 12
''''''                    lblGraphic.Caption = sTabDesc & ":  " & tvwGraphics(2).Nodes(NodeParKey).Parent.Text & " " & _
''''''                                NodeParText & "  [" & tvwGraphics(Index - 10).SelectedItem.Text & "]"
'''                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text '''& " " & NodeParText & ":  " & _
'''                            tvwGraphics(2).Nodes(NodeParKey).Text
'''            Case 13
'''
''''''                lblGraphic.Caption = Mid(tvwGraphics(3).SelectedItem.Parent.Text, 10) & " " & _
''''''                            tvwGraphics(3).SelectedItem.Text
'''                    lblGraphic.Caption = tvwGraphics(3).SelectedItem.Text
'''            Case 14
'''                lblGraphic.Caption = NodeParText
'''
'''            Case Else
'''                lblGraphic.Caption = sTabDesc & ":  " & NodeParText '''''& "  [" & nodetext & "]"
'''
'''
''''                    lblGraphic.Caption = sTabDesc & ":  " & NodeParText & "  [" & tvwGraphics(Index - 10).SelectedItem.Text & "]"
'''        End Select
    
    
'''    strSelect = "SELECT * " & _
'''                "FROM " & GFXMas & " " & _
'''                "WHERE GID = " & lGID
'''    Call GetGFXData(strSelect, "control")
    
NoGraphic:
    bPassIn = False
    lblGraphic.Visible = True
End Sub

Public Sub LoadThePicture(sPath As String)
    On Error GoTo ErrorOpening
    
    picJPG.Visible = False
    lblByGeorge(0).Visible = False
    lblByGeorge(1).Visible = False
    picJPG.Picture = LoadPicture()
    imgSize.Picture = LoadPicture(sPath)
    rAsp = imgSize.Width / imgSize.Height
    rFAsp = maxX / maxY
    Select Case rAsp
        Case Is = rFAsp
            With picJPG
                If maxX <= imgSize.Width Then
                    .Width = maxX
                    .Height = maxY
                    .Top = dTop: .Left = dLeft
                Else
                    .Width = imgSize.Width
                    .Height = imgSize.Height
                    .Left = dLeft + ((maxX - imgSize.Width) / 2)
                    .Top = dTop + ((maxY - imgSize.Height) / 2)
                End If
                
            End With
        Case Is > rFAsp
            With picJPG
                If maxX <= imgSize.Width Then
                    .Width = maxX
                    .Height = .Width / rAsp
                    .Left = dLeft
                    .Top = dTop + ((maxY - .Height) / 2)
                Else
                    .Width = imgSize.Width
                    .Height = imgSize.Height
                    .Left = dLeft + ((maxX - imgSize.Width) / 2)
                    .Top = dTop + ((maxY - imgSize.Height) / 2)
                End If
            End With
        Case Is < rFAsp
            With picJPG
                If maxY <= imgSize.Height Then
                    .Height = maxY
                    .Width = .Height * rAsp ''''' / rFAsp)
                    .Top = dTop
                    .Left = dLeft + ((maxX - .Width) / 2)
                Else
                    .Height = imgSize.Height
                    .Width = imgSize.Width
                    .Top = dTop + ((maxY - imgSize.Height) / 2)
                    .Left = dLeft + ((maxX - imgSize.Width) / 2)
                End If
            End With
    End Select
    
    picJPG.PaintPicture imgSize.Picture, 0, 0, picJPG.Width, picJPG.Height
    
    Call SetImageState
    
    '///// CHECK IF IMAGE COULD BE STRETCHED \\\\\
    If picJPG.Width < maxX _
                And picJPG.Height < maxY Then
        mnuResizeGraphic.Visible = True
        mnuMaxGraphic.Visible = True
        mnuResizeGraphic.Enabled = False
        mnuMaxGraphic.Enabled = True
        lblResize.Caption = "Maximize"
        lblResize.Enabled = True
        lblFullSize.Enabled = False
    Else
        mnuResizeGraphic.Visible = False
        mnuMaxGraphic.Visible = False
        lblResize.Enabled = False
        lblFullSize.Enabled = True
    End If
    
    picJPG.Visible = True
    imgMenu.Visible = True: lblMenu.Visible = True
    picMenu2.Visible = True
    bPicLoaded = True

Exit Sub

ErrorOpening:
    MsgBox "Error encountered while attempting to open file." & vbNewLine & _
                "Error:  " & Err.Description, vbCritical, "Cannot Open File..."
    picJPG.Visible = False
    imgMenu.Visible = False: lblMenu.Visible = False
    picMenu2.Visible = False
    bPicLoaded = False
'''    cmdMenu.Visible = False
'''    cmdResize.Visible = False
End Sub

Public Sub ResetNodeIcons(sNode As String)
    Dim i As Integer
    tvwGraphics.ImageList = ImageList1
    For i = 1 To tvwGraphics.Nodes.Count
        If tvwGraphics.Nodes(i).Key = sNode Then
            tvwGraphics.Nodes(i).Image = 6
        Else
            tvwGraphics.Nodes(i).Image = 1
        End If
    Next i
            

End Sub

Public Sub GetGFXData(strSelect As String, sDisplay As String)
    Dim sMess As String, sSize As String
    Dim rst As ADODB.Recordset
    Dim sGStatus(0 To 30) As String
    Dim lSize As Long
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(0) = "DE-ACTIVED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Select Case FileLen(Trim(rst.Fields("GPATH")))
            Case Is < 1000: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))), "#,##0") & " bytes"
            Case Is < 2000000: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))) / 1000, "#,##0") & " k"
            Case Else: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))) / 1000000, "#,##0.00") & " mb"
        End Select
            
        sMess = "Graphic Description:" & vbTab & Trim(rst.Fields("GDESC")) & vbNewLine & _
                    "Database I.D.:          " & vbTab & rst.Fields("GID") & vbNewLine & _
                    "File Format:              " & vbTab & Trim(rst.Fields("GFORMAT")) & vbNewLine & _
                    "Graphic Type:           " & vbTab & GfxType(rst.Fields("GTYPE")) & vbNewLine & _
                    "Graphic Status:         " & vbTab & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                    "File Size:                   " & vbTab & sSize & vbNewLine & vbNewLine
        sMess = sMess & "File Added by " & Trim(rst.Fields("ADDUSER")) & " on " & _
                    Format(rst.Fields("ADDDTTM"), "mmmm d, yyyy") & "." & vbNewLine
        sMess = sMess & "File Last Edited by " & Trim(rst.Fields("UPDUSER")) & " on " & _
                    Format(rst.Fields("UPDDTTM"), "mmmm d, yyyy") & "."
        rst.Close
        Select Case sDisplay
            Case "msgbox"
                MsgBox sMess, vbInformation, "Graphic Data..."
'''            Case "control"
'''                txtXData(1).Text = sMess
        End Select
'        picXData.Visible = True
    Else
        rst.Close
'''        picXData.Visible = False
        MsgBox "No Data Available.", vbInformation, "Graphic Data..."
    End If
    Set rst = Nothing

End Sub


Public Sub WhatSupDoc(tGID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT GM.SUPDOC_ID, SD.SUPDOCFORMAT AS FORMAT " & _
                "FROM ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_SUPDOC SD " & _
                "WHERE GM.GID > 0 " & _
                "AND GM.GID = " & tGID & " " & _
                "AND GM.SUPDOC_ID > 0 " & _
                "AND GM.SUPDOC_ID = SD.SUPDOC_ID"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        imgSupDoc.Tag = CStr(rst.Fields("SUPDOC_ID"))
        imgSupDoc.Visible = True
    Else
        imgSupDoc.Tag = ""
        imgSupDoc.Visible = False
    End If
    rst.Close: Set rst = Nothing
    
End Sub

Public Sub ResetBatch(iStart As Integer, iCnt As Integer, CntIndex As Integer)
    Dim sMess As String
    
    If iStart = 1 Then
        lblBatch(0).Enabled = False
        lblBatch(2).Enabled = False
    Else
        lblBatch(0).Enabled = True
        lblBatch(2).Enabled = True
    End If
    
    If iCnt >= iStart + (iCols * iRows) Then
        lblBatch(1).Enabled = True
        lblBatch(3).Enabled = True
    Else
        lblBatch(1).Enabled = False
        lblBatch(3).Enabled = False
    End If
    
    If iCnt > (iCols * iRows) Then
        lblList.Visible = True
        fraBatch.Visible = True
    Else
        lblList.Visible = False
        fraBatch.Visible = False
    End If
    
    If iStart + (iCols * iRows) < iCnt Then
        lblCnt.Caption = "Images: " & iStart & " - " & iStart + (iCols * iRows - 1) & _
                    " of " & iCnt
    Else
        lblCnt.Caption = "Images: " & iStart & " - " & iCnt & _
                    " of " & iCnt
    End If
    
End Sub

Public Sub ClearModes()
    Dim i As Integer
    bDMode = False: mnuDownloadMode.Checked = False: mnuDownloadSels.Enabled = False: lblDownload.ForeColor = vbButtonText: mnuDownloadSels2.Enabled = False
    bEMode = False: mnuEmailMode.Checked = False: mnuEmailSels.Enabled = False: lblEmail.ForeColor = vbButtonText: mnuEmailSels2.Enabled = False
    For i = 0 To imx0.Count - 1
        imx0(i).Enabled = True
        chkMulti(i).Visible = False
        chkMulti(i).Value = 0
    Next i
End Sub

Public Function GetHeader(nodX As Node) As String
    Dim tHDR As String
    
    On Error GoTo Complete
    tHDR = nodX.Text
    tHDR = nodX.Parent.Text & " (" & tHDR & ")"
    tHDR = nodX.Parent.Parent.Text & " (" & tHDR & ")"
    tHDR = nodX.Parent.Parent.Parent.Text & " (" & tHDR & ")"
    tHDR = nodX.Parent.Parent.Parent.Parent.Text & " (" & tHDR & ")"
    tHDR = nodX.Parent.Parent.Parent.Parent.Parent.Text & " (" & tHDR & ")"
    
    GetHeader = tHDR
Exit Function
Complete:
    GetHeader = tHDR
End Function

Private Sub web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'    Debug.Print "Progress: Complete"
'    If web1.LocationName = "http:///" Then '''web1.Visible = False
'        web1.Visible = True
'    Else
'        web1.Visible = False
'    End If
End Sub

Private Sub web1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    Debug.Print "Progress: " & Progress & " -- " & ProgressMax
    If Progress = 0 Then web1.Visible = True
End Sub


Public Sub CheckDownloadCart(pType As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT COUNT(DLID) AS CNT " & _
                "FROM ANNOTATOR.ANO_DOWNLOAD " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND FILE_TYPE = '" & pType & "' " & _
                "AND DLSTATUS = 1"
    Set rst = Conn.Execute(strSelect)
    If rst.Fields("CNT") > 0 Then
        lblDLCnt.Caption = rst.Fields("CNT")
        picDLCart.Visible = True
    Else
        lblDLCnt.Caption = "0"
        picDLCart.Visible = False
    End If
    rst.Close: Set rst = Nothing
    
End Sub


Private Sub movelines(X As Single, Y As Single)

    If Not (xs = 0 And ys = 0) Then

        'delete previous
        '''-frmGrab.Line (xStart, yStart)-(xs - 1, ys - 1), , B
        picInner(0).Line (xStart, yStart)-(xs, ys), , B
    End If

    'draw selection square in invert drawmode
    '''-frmGrab.Line (xStart, yStart)-(x - 1, y - 1), , B
'''    Debug.Print Abs(X - xStart) & ", " & Abs(Y - yStart)
'''
'''    If Abs(X - xStart) < Abs(Y - yStart) Then
'''        Debug.Print "Change y to " & yStart + (X - xStart)
'''        Y = yStart + (X - xStart)
'''    Else
'''        Debug.Print "Change x to " & xStart + (Y - yStart)
'''        X = xStart + (Y - yStart)
'''    End If
        
    picInner(0).Line (xStart, yStart)-(X, Y), , B
    
    xs = X: ys = Y
End Sub

Public Sub GetWindowSels(x0 As Single, y0 As Single, X1 As Single, Y1 As Single)
    Dim i As Integer
    
    For i = 0 To chkMulti.Count - 1
        If chkMulti(i).Left >= x0 And chkMulti(i).Left <= X1 _
                    And chkMulti(i).Top >= y0 And chkMulti(i).Top <= Y1 Then
            chkMulti(i).Value = 1
        End If
    Next i
End Sub

Public Sub ClearDownloadCart(pType As String)
    Dim strDelete As String
    
    strDelete = "DELETE FROM ANNOTATOR.ANO_DOWNLOAD " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND FILE_TYPE = '" & pType & "' " & _
                "AND DLSTATUS = 1"
    Conn.Execute (strDelete)
    
End Sub
