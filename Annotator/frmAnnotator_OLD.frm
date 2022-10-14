VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAnnotator 
   Caption         =   "GPJ Space Plan Viewing & Annotation"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnnotator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   30
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11460
      MouseIcon       =   "frmAnnotator.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   300
      Width           =   795
   End
   Begin TabDlg.SSTab sst1 
      Height          =   6375
      Left            =   0
      TabIndex        =   59
      Top             =   540
      Visible         =   0   'False
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   11245
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAnnotator.frx":074C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvw2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkFPApprove"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CheckBox chkFPApprove 
         Caption         =   "Display only Floorplans requiring my Review..."
         Height          =   435
         Left            =   6240
         TabIndex        =   77
         Top             =   240
         Width           =   2115
      End
      Begin MSComctlLib.TreeView tvw2 
         Height          =   6015
         Left            =   180
         TabIndex        =   62
         Top             =   180
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   10610
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
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
         MouseIcon       =   "frmAnnotator.frx":0768
      End
      Begin VB.Frame fraFP 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   60
         TabIndex        =   63
         Top             =   60
         Visible         =   0   'False
         Width           =   10635
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   195
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   69
            Top             =   5940
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CommandButton cmdFPSI 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9900
            Picture         =   "frmAnnotator.frx":0A82
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   180
            Width           =   615
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmAnnotator.frx":0D8C
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   9900
            Picture         =   "frmAnnotator.frx":1096
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Sort Show List Chronologically (by Show Open Date)"
            Top             =   840
            Width           =   615
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmAnnotator.frx":1960
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   9300
            Picture         =   "frmAnnotator.frx":1C6A
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Sort Show List Alphabetically"
            Top             =   840
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.CommandButton cmdFPS 
            Caption         =   "============= Floorplan Status ============="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Click for Definitions of Floorplan Status Values"
            Top             =   1605
            Width           =   3855
         End
         Begin MSFlexGridLib.MSFlexGrid flx1 
            Height          =   3855
            Left            =   120
            TabIndex        =   67
            Top             =   1980
            Visible         =   0   'False
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   20
            Cols            =   12
            FixedCols       =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            GridColorFixed  =   12632256
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmAnnotator.frx":2534
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   61
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Year:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.PictureBox picFPApprove 
      Height          =   2595
      Left            =   540
      ScaleHeight     =   2535
      ScaleWidth      =   3915
      TabIndex        =   70
      Top             =   6780
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdFPApprove 
         Caption         =   "Approve"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtFPApprove 
         Height          =   1215
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   75
         Top             =   1080
         Width           =   3795
      End
      Begin VB.OptionButton optFPApprove 
         Caption         =   "Approved w/Attached Redlines"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   73
         Top             =   600
         Width           =   2595
      End
      Begin VB.OptionButton optFPApprove 
         Caption         =   "Layout Approved as Drawn"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   72
         Top             =   360
         Width           =   2355
      End
      Begin VB.OptionButton optFPApprove 
         Caption         =   "Approved w/Following Comments"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   74
         Top             =   840
         Width           =   2775
      End
      Begin VB.Image imgMinMax 
         Height          =   315
         Left            =   3180
         Picture         =   "frmAnnotator.frx":284E
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Minimize"
         Top             =   15
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Entered Comments will be Saved with all Approvals"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   60
         TabIndex        =   78
         Top             =   2340
         Width           =   3360
      End
      Begin VB.Image imgFPApprove 
         Height          =   315
         Left            =   3570
         Picture         =   "frmAnnotator.frx":2B58
         Stretch         =   -1  'True
         ToolTipText     =   "Click to Close"
         Top             =   15
         Width           =   315
      End
      Begin VB.Label lblFPApprove 
         BackColor       =   &H0000FFFF&
         Caption         =   " Floorplan Approval..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   4635
      End
   End
   Begin VB.CommandButton cmdFPApproveHide 
      Height          =   495
      Left            =   11340
      Picture         =   "frmAnnotator.frx":2E62
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Click to Re-Open Approval Interface"
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picDirs 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   5640
      ScaleHeight     =   7455
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdNav 
         BackColor       =   &H00000000&
         DisabledPicture =   "frmAnnotator.frx":306C
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   600
         MaskColor       =   &H80000002&
         MouseIcon       =   "frmAnnotator.frx":3376
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":3680
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Step Forward through List"
         Top             =   6900
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         BackColor       =   &H00000000&
         DisabledPicture =   "frmAnnotator.frx":398A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   60
         MaskColor       =   &H80000006&
         MouseIcon       =   "frmAnnotator.frx":3C94
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":3F9E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Step Back through List"
         Top             =   6900
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox lstDwgSorter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   4080
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox lstClientSorter 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComctlLib.TreeView tvw1 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   2990
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   4
         Left            =   4680
         MouseIcon       =   "frmAnnotator.frx":42A8
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":45B2
         Stretch         =   -1  'True
         ToolTipText     =   "Vignettes"
         Top             =   3420
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   4680
         MouseIcon       =   "frmAnnotator.frx":48BC
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":4BC6
         Stretch         =   -1  'True
         ToolTipText     =   "Show Plans"
         Top             =   2820
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgFPS 
         Height          =   495
         Left            =   4680
         MouseIcon       =   "frmAnnotator.frx":4ED0
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":51DA
         Stretch         =   -1  'True
         ToolTipText     =   "Click to View Floorplan Status"
         Top             =   6900
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   4680
         MouseIcon       =   "frmAnnotator.frx":54E4
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":57EE
         Stretch         =   -1  'True
         ToolTipText     =   "Floorplans"
         Top             =   2220
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   47
         Top             =   2100
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblClient 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select File Type to view from Icons below..."
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
         Height          =   360
         Left            =   0
         TabIndex        =   46
         Top             =   30
         Width           =   5235
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   19
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":5AF8
         TabIndex        =   45
         Top             =   6660
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   18
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":5E02
         TabIndex        =   44
         Top             =   6420
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   17
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":610C
         TabIndex        =   43
         Top             =   6180
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   16
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":6416
         TabIndex        =   42
         Top             =   5940
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   15
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":6720
         TabIndex        =   41
         Top             =   5700
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   14
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":6A2A
         TabIndex        =   40
         Top             =   5460
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   13
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":6D34
         TabIndex        =   39
         Top             =   5220
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   12
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":703E
         TabIndex        =   38
         Top             =   4980
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   11
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":7348
         TabIndex        =   37
         Top             =   4740
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   10
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":7652
         TabIndex        =   36
         Top             =   4500
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   9
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":795C
         TabIndex        =   35
         Top             =   4260
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   8
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":7C66
         TabIndex        =   34
         Top             =   4020
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   7
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":7F70
         TabIndex        =   33
         Top             =   3780
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   6
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":827A
         TabIndex        =   32
         Top             =   3540
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   5
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":8584
         TabIndex        =   31
         Top             =   3300
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   4
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":888E
         TabIndex        =   30
         Top             =   3060
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   3
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":8B98
         TabIndex        =   29
         Top             =   2820
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   2
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":8EA2
         TabIndex        =   28
         Top             =   2580
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":91AC
         TabIndex        =   27
         Top             =   2340
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFile 
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
         ForeColor       =   &H80000013&
         Height          =   195
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmAnnotator.frx":94B6
         TabIndex        =   26
         Top             =   2100
         UseMnemonic     =   0   'False
         Width           =   45
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   25
         Top             =   2340
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   24
         Top             =   2580
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   23
         Top             =   2820
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   3060
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   21
         Top             =   3300
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   20
         Top             =   3540
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   19
         Top             =   3780
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   18
         Top             =   4020
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   17
         Top             =   4260
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   16
         Top             =   4500
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   15
         Top             =   4740
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   14
         Top             =   4980
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   13
         Left            =   480
         TabIndex        =   13
         Top             =   5220
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   14
         Left            =   480
         TabIndex        =   12
         Top             =   5460
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   11
         Top             =   5700
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   10
         Top             =   5940
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   17
         Left            =   480
         TabIndex        =   9
         Top             =   6180
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   18
         Left            =   480
         TabIndex        =   8
         Top             =   6420
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblExt 
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
         Height          =   255
         Index           =   19
         Left            =   480
         TabIndex        =   7
         Top             =   6660
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On-Site Photos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1200
         TabIndex        =   6
         Top             =   6900
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Image imgTeam 
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmAnnotator.frx":97C0
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":9ACA
         Stretch         =   -1  'True
         ToolTipText     =   "Select to setup Email Notification Team"
         Top             =   6900
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Viewer Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Click to Access Viewer Options/Right-Click to Hide"
      Top             =   480
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdDirs 
      Caption         =   "Open File Index..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MouseIcon       =   "frmAnnotator.frx":9DD4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   60
      Width           =   1875
   End
   Begin VB.PictureBox picRelatives 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   795
      TabIndex        =   49
      ToolTipText     =   "Relatives Menu (Dbl-click to Re-Dock)"
      Top             =   780
      Visible         =   0   'False
      Width           =   855
      Begin VB.CommandButton cmdOther 
         Caption         =   "Others..."
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4380
         Width           =   795
      End
      Begin VB.Image imgConst 
         Height          =   435
         Left            =   960
         MouseIcon       =   "frmAnnotator.frx":A0DE
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":A3E8
         ToolTipText     =   "Click to view Property Drawings"
         Top             =   4440
         Width           =   480
      End
      Begin VB.Image imgGraphics 
         Height          =   555
         Left            =   900
         MouseIcon       =   "frmAnnotator.frx":A510
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":A81A
         Stretch         =   -1  'True
         ToolTipText     =   "Click to view Assigned Graphics"
         Top             =   3840
         Width           =   555
      End
      Begin VB.Image imgInfo 
         Height          =   555
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":A964
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":AC6E
         Stretch         =   -1  'True
         ToolTipText     =   "Click to view Show Information"
         Top             =   3720
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":AF78
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":B282
         Stretch         =   -1  'True
         ToolTipText     =   "Return to Floor Plan"
         Top             =   120
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":B58C
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":B896
         Stretch         =   -1  'True
         ToolTipText     =   "View Isometric"
         Top             =   720
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":BBA0
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":BEAA
         Stretch         =   -1  'True
         ToolTipText     =   "View Elevations"
         Top             =   1320
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":C1B4
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":C4BE
         Stretch         =   -1  'True
         ToolTipText     =   "View Overall Show Plan"
         Top             =   1920
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   4
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":C7C8
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":CAD2
         Stretch         =   -1  'True
         ToolTipText     =   "View Composite Show Plan"
         Top             =   2520
         Width           =   555
      End
      Begin VB.Image imgDWF 
         Height          =   555
         Index           =   5
         Left            =   120
         MouseIcon       =   "frmAnnotator.frx":CDDC
         MousePointer    =   99  'Custom
         Picture         =   "frmAnnotator.frx":D0E6
         Stretch         =   -1  'True
         ToolTipText     =   "View Press Plan"
         Top             =   3120
         Width           =   555
      End
   End
   Begin VOLOVIEWXLibCtl.AvViewX volFrame 
      Height          =   2295
      Left            =   0
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      _cx             =   4200045
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
      SupportPath     =   $"frmAnnotator.frx":D3F0
      FontPath        =   $"frmAnnotator.frx":79E6B
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
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   0
      ScaleHeight     =   8055
      ScaleWidth      =   12000
      TabIndex        =   51
      Top             =   540
      Width           =   12000
      Begin VB.Label lblByGeorge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GPJ Space Plans"
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
         TabIndex        =   57
         Top             =   6240
         Width           =   4815
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
         Left            =   2340
         TabIndex        =   56
         Top             =   5400
         Width           =   8985
      End
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   3
      Left            =   11220
      Picture         =   "frmAnnotator.frx":B2C18
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblLock 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11400
      TabIndex        =   55
      Top             =   450
      Width           =   45
   End
   Begin VB.Image imgLock 
      Height          =   405
      Left            =   11100
      MouseIcon       =   "frmAnnotator.frx":B2F22
      MousePointer    =   99  'Custom
      Picture         =   "frmAnnotator.frx":B322C
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   360
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2760
      TabIndex        =   54
      Top             =   420
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Annotator is loading..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   2760
      TabIndex        =   53
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   2715
   End
   Begin VB.Image imgComm 
      Height          =   480
      Left            =   1980
      MouseIcon       =   "frmAnnotator.frx":B3536
      MousePointer    =   99  'Custom
      Picture         =   "frmAnnotator.frx":B3840
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   0
      Left            =   11100
      Picture         =   "frmAnnotator.frx":B3C82
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   1
      Left            =   11100
      Picture         =   "frmAnnotator.frx":B40C4
      Top             =   2700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   2
      Left            =   11100
      Picture         =   "frmAnnotator.frx":B4506
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuPanMode 
         Caption         =   " Set Pan as Current Mode"
      End
      Begin VB.Menu mnuZoomDMode 
         Caption         =   "Set Dynamic Zoom as Current Mode"
      End
      Begin VB.Menu mnuZoomWMode 
         Caption         =   "Set Zoom Window as Current Mode"
      End
      Begin VB.Menu mnuZoomFull 
         Caption         =   "Zoom to Full View"
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRedlining 
         Caption         =   "Annotation"
         Begin VB.Menu mnuRedlines 
            Caption         =   "Redlines"
            Begin VB.Menu mnuRedLoad 
               Caption         =   "Load Saved Redline File"
            End
            Begin VB.Menu mnuRedReturn 
               Caption         =   "Return to Original Drawing"
            End
            Begin VB.Menu mnuRedClear 
               Caption         =   "Clear Redlines"
            End
            Begin VB.Menu mnuRedDelete 
               Caption         =   "Delete Redlines"
            End
            Begin VB.Menu mnuRedSave 
               Caption         =   "Save Redlines"
            End
         End
         Begin VB.Menu mnuRedMode 
            Caption         =   "Set 'Sketch' as Current Mode"
         End
         Begin VB.Menu mnuTextMode 
            Caption         =   "Set 'Text' as Current Mode"
         End
         Begin VB.Menu mnuRedModeStop 
            Caption         =   "Discontinue Redline Mode"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComments1 
         Caption         =   "Comments Interface..."
      End
      Begin VB.Menu mnuDash10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display with Default Colors"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display Black on White"
         Index           =   1
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display ClearScale"
         Index           =   2
      End
      Begin VB.Menu mnuMax 
         Caption         =   "Maximize Display"
      End
      Begin VB.Menu mnuDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuDash05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmailPDF 
         Caption         =   "Email PDF of Current Plan..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownloadPDF 
         Caption         =   "Download PDF of Current Plan..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download Copy of DWF File..."
      End
      Begin VB.Menu mnuSendALink 
         Caption         =   "Send-A-Link..."
      End
      Begin VB.Menu mnuDash06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuMenuButton1 
         Caption         =   "Make 'Viewer Menu' Button Available"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCancel1 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuHideMenuButton 
      Caption         =   "mnuHideMenuButton"
      Visible         =   0   'False
      Begin VB.Menu mnuHide 
         Caption         =   "Hide 'Viewer Menu' Button (Use Right-Click Menu)"
      End
      Begin VB.Menu mnuDash08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel3 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuLockPop 
      Caption         =   "mnuLockPop"
      Visible         =   0   'False
      Begin VB.Menu mnuLockStatus 
         Caption         =   "Check Lock Status"
      End
      Begin VB.Menu mnuLockRemove 
         Caption         =   "Remove Lock"
      End
      Begin VB.Menu mnuDash09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel4 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuRC 
      Caption         =   "mnuRC"
      Visible         =   0   'False
      Begin VB.Menu mnuEng 
         Caption         =   "View Engineering Drawings..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGFX 
         Caption         =   "View Associated Graphics..."
      End
      Begin VB.Menu mnuPhoto 
         Caption         =   "View Digital Photo..."
      End
      Begin VB.Menu mnuDash11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogistical 
         Caption         =   "View Logistical Data..."
      End
      Begin VB.Menu mnuUsage 
         Caption         =   "View Element Usage..."
      End
      Begin VB.Menu mnuDash12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel5 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmAnnotator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''Private strClientList As String
Private Const LOCALE_SYSTEM_DEFAULT& = &H800
Private Const LOCALE_USER_DEFAULT& = &H400
Dim TimeOff As Long
Dim Redlining As Boolean, bLoading As Boolean, Selection As Boolean, Zooming As Boolean, _
            SaveRed As Boolean, bRemoving As Boolean, bNotes As Boolean, _
            bReded As Boolean, bGPJNotes As Boolean, _
            bPicLoaded As Boolean, bIBM As Boolean, bTeam As Boolean
Dim CurrFile As String, RedFile As String, EmailHelp As String, sMap As String
Dim iType As Integer, iFileStart As Integer, iLW As Integer, Ind As Integer, _
            iDir As Integer, iSHYR As Integer, iD As Integer
Dim bClient As Boolean, PickPic As Boolean, CommentSaved As Boolean, bComm As Boolean
Dim sCurrPath As String, sCurrShow As String, sCurrClient As String, sClientList As String, _
            sSHYR As String, sPreSHYR As String, sCliPath As String, sPicPath As String, sCurrPic As String, _
            sPrePath As String, sPrePic As String, sPrePicPath As String, sCommPath As String, sRedClient As String, _
            sRedShow As String, sShowPlans As String, sVignettes As String, _
            sPreShow As String, sIBMGroup As String, sMessPath As String, sPreClient As String
Dim sPath As String, RTXFile As String, PatternFiles As String, EMailResp As String, _
            GPJFile As String, AnnoLog As String, _
            FPPath As String, PRPath As String, GRPath As String, FPFile As String, _
            PRFile As String, GRFile As String
Dim xComm As Single, yComm As Single
Dim CurrX As Long, CurrY As Long
Dim dLeft As Double, dRight As Double, dTop As Double, dBottom As Double
Dim CommOpen As Boolean, RelOpen As Boolean, bMapFound As Boolean, bMenuButton As Boolean
Dim NewUse As Boolean, bResetting As Boolean, FirstTime As Boolean, bViewSet As Boolean
Dim TPos As Integer
Dim xStr As Long, yStr As Long
Dim sDWF(5) As String
Dim ServerPath As String
Dim LCKFile As String, LCKFullName As String, LCKLotusName As String
Dim LCKTime As Date, LCKSaveTime As Date
Dim bLock As Boolean
Dim RelativePath(0 To 5) As String
Dim nLockRefID As Long, lLockRefID As Long, lLockID As Long, lDWGID As Long, lSHTID As Long, _
            lRedID As Long, lNewLockId As Long
Dim sDWFPath As String
Dim redBCC As String
Dim redSHCD As Long
Dim redSHYR As Integer
Dim SelOrder(0 To 1) As String
Dim sZMode As String
Dim iFlxCol As Integer
Dim sShowDates As String, sApprClientList As String, sApprClientList2 As String
Dim bApprover As Boolean, bApprovalList As Boolean

Dim fSHYR As Integer
Dim fBCC As String, fFBCN As String
Dim fSHCD As Long

'Const definitions
Const BestFit = 0

'**********************************
'**  Type Definitions:

#If Win32 Then
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

#End If 'WIN32 Types

'**********************************
'**  Function Declarations:

#If Win32 Then
Private Declare Function GetLocaleInfo& Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetSystemTimeAdjustment Lib "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Long)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeZoneInformation& Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION)
Private Declare Function GetTimeFormat& Lib "kernel32" Alias "GetTimeFormatA" _
        (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, _
        ByVal lpFormat As Long, ByVal lpTimeStr As String, ByVal cchTime As Long)
#End If 'WIN32



Private Sub chkFPApprove_Click()
    Dim sNode As String
    
    On Error Resume Next
    sNode = tvw2.SelectedItem.key
    If Err Then
        Err.Clear
        sNode = ""
    End If
    
    Select Case chkFPApprove.value
        Case 1
            bApprovalList = True
            Me.MousePointer = 11
            Call PopApprovalClients(sApprClientList2)
        Case 0
            bApprovalList = False
            Me.MousePointer = 11
            Call PopClients
    End Select
    
    If sNode <> "" Then
        tvw2.Nodes(sNode).Selected = True
    End If
    
    Me.MousePointer = 0
End Sub

'Private Sub cboSHYR_Click()
'    Dim sClient As String
'
'    If cboCUNO.Text <> "" Then sClient = cboCUNO.Text
'     Call PopClients2
'     On Error Resume Next
'     cboCUNO.Text = sClient
'End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFPApprove_Click()
    ''---WILL NEED TO RESET ORACLE RIGHTS FOR 'ANNOTATOR' TO UPDATE IGL_KIT_USE---''
    
    ''Reset IGL_KIT_USE, FPSTATUS=5''
    ''Reset DWG_MASTER, DSTATUS=5''
    ''Write Aprroval record to ANO_COMMENT (w/User Comment, if any)''
    ''Send Email''
        ''Loop thru Mand Recips & create User-based Link''
        ''Should I attach Redline (if part of Approval) in email?''
    ''Step thru flxgrid to current and reset status to 5 (w/date)''
    ''picFPApprove.visible = False''
    
'    Dim strSelect As String, strInsert As String, strUpdate As String
'    Dim rstL As ADODB.Recordset
'    Dim lCOMMID As Long
'    Dim sNewComm As String, sGetDate As String, sDate As String
'
'    On Error Resume Next
'    Conn.BeginTrans
'
'    ''Reset IGL_KIT_USE, FPSTATUS=5''
'    strUpdate = "UPDATE " & IGLKitU & " " & _
'                "SET FPSTATUS = 5, " & _
'                "FPSTATBY = '" & Left(DeGlitch(LogName), 16) & "', " & _
'                "FPSTATDT = SYSDATE, " & _
'                "UPDUSER = '" & Left(DeGlitch(LogName), 16) & "', " & _
'                "UPDDTTM = SYSDATE, " & _
'                "UPDCNT = UPDCNT + 1 " & _
'                "WHERE AN8_CUNO = " & CLng(BCC) & " " & _
'                "AND AN8_SHCD = " & SHCD & " " & _
'                "AND SHYR = " & SHYR
'    Conn.Execute (strUpdate)
'
'    ''Reset DWG_MASTER, DSTATUS=5''
'    strUpdate = "UPDATE " & DWGMas & " " & _
'                "SET DSTATUS = 5, " & _
'                "UPDUSER = '" & Left(DeGlitch(LogName), 24) & "', " & _
'                "UPDDTTM = SYSDATE, " & _
'                "UPDCNT = UPDCNT + 1 " & _
'                "WHERE DWGID = " & lDWGID
'    Conn.Execute (strUpdate)
'
'    ''Write Aprroval record to ANO_COMMENT (w/User Comment, if any)''
'    sGetDate = "SELECT TO_CHAR(SYSDATE, 'MON DD, YYYY HH:MM PM') AS SDATE FROM DUAL"
'    Set rstL = Conn.Execute(sGetDate)
'    sDate = rstL.Fields("SDATE")
'    rstL.Close: Set rstL = Nothing
'    sNewComm = "Layout Approved by " & LogName & " (" & sDate & ")"
'    If txtFPApprove.Text <> "" Then
'        sNewComm = sNewComm & vbCr & _
'                    "Approver Comment:  " & Trim(txtFPApprove.Text)
'    End If
'
'    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
'    lCOMMID = rstL.Fields("nextval")
'    rstL.Close: Set rstL = Nothing
'
'    strInsert = "INSERT INTO " & ANOComment & " " & _
'                "(COMMID, REFID, REFSOURCE, ANO_COMMENT, " & _
'                "COMMUSER, COMMDATE, COMMSTATUS, " & _
'                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
'                "VALUES " & _
'                "(" & lCOMMID & ", " & lDWGID & ", 'DWG_MASTER', '" & sNewComm & "', " & _
'                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
'                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
'    Conn.Execute (strInsert)
'
'    If Err = 0 Then
'        Conn.CommitTrans
'    Else
'        Conn.RollbackTrans
'        MsgBox "For the following reason you are unable to edit one of the database records.  " & _
'                    "You will need to manually alert your Floorplan Specialist " & _
'                    "of Approval and he/she can reset the Floorplan Status." & vbNewLine & _
'                    vbNewLine & "Error: " & Err.Description, vbExclamation, "Sorry..."
'        Exit Sub
'    End If
    
    
End Sub

Private Sub cmdFPApproveHide_Click()
    picFPApprove.Top = volFrame.Top + 120
    picFPApprove.Left = volFrame.Left + volFrame.Width - 90 - picFPApprove.Width
    picFPApprove.Visible = True
    cmdFPApproveHide.Visible = False
End Sub

Private Sub cmdFPS_Click()
    Dim sMess As String
    
    sMess = "Floorplan Status is advanced through the Floorplan Application, " & _
                "and is displayed in a simpler form on the Master Show Schedule.  " & _
                "The Dates shown on the status bars of this form reflect the last Status Date change." & _
                vbNewLine & vbNewLine
    sMess = sMess & "REQ" & vbTab & "STATUS LEVEL 1 - A Floorplan has been requested, but not started." & vbNewLine
    sMess = sMess & "FSU" & vbTab & "STATUS LEVEL 2 - The Floorplan has been setup in the Floorplan Application." & vbNewLine
    sMess = sMess & "BGD" & vbTab & "STATUS LEVEL 3 - The Background Space has been drawn, awaiting Properties." & vbNewLine
    sMess = sMess & "PRE" & vbTab & "STATUS LEVEL 4 - A Preliminary Layout has been completed, awaiting A/E approval." & vbNewLine
    sMess = sMess & vbTab & "NOTE:  It is at this level that the Floorplan is made available on the Annotator." & vbNewLine
    sMess = sMess & "AEA" & vbTab & "STATUS LEVEL 5 - The Account Executive has approved the layout, and it is ready to complete." & vbNewLine
    sMess = sMess & "CMP" & vbTab & "STATUS LEVEL 6 - The Floorplan has been completed and is awaiting review, prior to Release." & vbNewLine
    sMess = sMess & "REL" & vbTab & "STATUS LEVEL 7 - The Floorplan has been Released." & vbNewLine
    sMess = sMess & "REV" & vbTab & "STATUS LEVEL 8 - The Floorplan has been Revised and re-Released."
    sMess = sMess & vbNewLine & vbNewLine
    sMess = sMess & "For a more complete view, click on the 'Floorplan Status' icon above."
    
    MsgBox sMess, vbInformation, "Floorplan Status Definitions..."
End Sub

Private Sub cmdFPSI_Click()
    frmFPStatus.PassBCC = fBCC
    frmFPStatus.PassFBCN = fFBCN
    frmFPStatus.Show 1
End Sub

Private Sub cmdMenu_Click()
    If volFrame.Visible = True Then
        Me.PopupMenu mnuRightClick, 0, cmdMenu.Left, cmdMenu.Top + cmdMenu.Height
    End If
End Sub

Private Sub cmdMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuHideMenuButton
    End If
End Sub


Private Sub cmdOther_Click()
''''' /// ADD LATER \\\   frmOtherViews.PassSHTID = SHTID
    frmOtherViews.Show 1
End Sub

Private Sub cmdSettings_Click()
    frmSettings.PassFrom = "FP"
    frmSettings.PassBCC = CLng(BCC)
    frmSettings.PassFBCN = FBCN
    frmSettings.PassBCC_DEF = defCUNO ''' lBCC_Def
    frmSettings.PassFBCN_DEF = defFBCN ''' sFBCN_Def
    frmSettings.Show 1
End Sub

Private Sub flx1_Click()
    Dim iRow As Integer, iPointer As Integer
    Dim Resp As VbMsgBoxResult
    
    If iFlxCol = 0 Then
'''        bLoading = True
        iRow = flx1.RowSel ''' Int(Y / flx1.RowHeight(0)) + flx1.TopRow - 1
        flx1.Row = iRow: flx1.Col = 0
        If flx1.CellFontBold = True Then
            
            '///// CHECK IF RED SHOULD BE SAVED \\\\\
            If SaveRed = True Then
                Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNoCancel, "Redline Changes...")
                If Resp = vbYes Then
                    mnuRedSave_Click
                ElseIf Resp = vbNo Then
                    volFrame.ClearMarkup
        '''            SaveRed = False
                ElseIf Resp = vbCancel Then
                    Exit Sub
                End If
            End If
            
            If bReded = True And bTeam = True And bPerm(15) Then
                With frmRedAlert
                    .PassSHYR = redSHYR
                    .PassBCC = CLng(redBCC)
                    .PassSHCD = redSHCD
                    .PassHDR = lblWelcome
                    .PassType = 0
                    .Show 1
                End With
        '''        Call RedAlert(0, lblWelcome, redBCC, redSHCD) 'AlertOfRed
            End If
            bReded = False
            redBCC = "": redSHCD = 0: redSHYR = 0
            
            SHYR = fSHYR
            BCC = fBCC
            FBCN = fFBCN
            SHNM = flx1.Text
            SHCD = flx1.TextMatrix(iRow, 10)
            iPointer = Me.MousePointer
            Me.MousePointer = 11
            Call LoadFloorplan(SHYR, CLng(BCC), SHCD, SHNM)
            If volFrame.Visible Then
                volFrame.SetFocus
                sShowDates = flx1.TextMatrix(iRow, 1)
            End If
            Me.MousePointer = iPointer
        End If
    End If
End Sub

Private Sub flx1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRow As Integer, iPointer As Integer
    
    If flx1.Visible = False Then Exit Sub
    
    If X < flx1.ColWidth(0) Then iFlxCol = 0 Else iFlxCol = 1
    
    If Y < flx1.RowHeight(0) And X < flx1.ColWidth(0) Then
        optSort(0).value = True
        Exit Sub
    ElseIf Y < flx1.RowHeight(0) And X < (flx1.ColWidth(0) + flx1.ColWidth(1)) Then
        optSort(1).value = True
        Exit Sub
    End If
    
End Sub

Private Sub flx1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRow As Integer
    iRow = Int(Y / flx1.RowHeight(0)) + flx1.TopRow - 1
    If iRow >= flx1.Rows Then iRow = flx1.Rows - 1
    flx1.Row = iRow: flx1.Col = 0
    If flx1.CellFontBold = True Then flx1.MousePointer = 99 Else flx1.MousePointer = 0
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Debug.Print "Dragging " & Source.Name
    If Source.Name = "picFPApprove" Then
        With picFPApprove
            .Drag 2
            picFPApprove.Visible = False
            cmdFPApproveHide.Visible = True
        End With
    End If
    
End Sub

Private Sub Form_Resize()
    Debug.Print "WindowState = " & Me.WindowState
    Debug.Print Me.Width & "," & Me.Height
    
    If Me.WindowState <> 1 Then
        If Me.Width > 2000 And Me.Height > 2000 Then
            Debug.Print "X = " & Me.Width & ", Y = " & Me.Height
            volFrame.Width = Me.Width - 360
            volFrame.Height = Me.Height - 1230 ''' volFrame.Width * 2 / 3
            picFrame.Width = volFrame.Width
            picFrame.Height = volFrame.Height
            picFPApprove.Left = volFrame.Left + volFrame.Width - 90 - picFPApprove.Width
            
            '///// POSITION LBLBYGEORGEs \\\\\
            lblByGeorge(0).Left = picFrame.Width - 240 - lblByGeorge(0).Width
            lblByGeorge(0).Top = picFrame.Height - 240 - lblByGeorge(0).Height
            lblByGeorge(1).Left = 240
            lblByGeorge(1).Top = lblByGeorge(0).Top + 840
            
            cmdClose.Left = Me.ScaleWidth - 120 - cmdClose.Width
            cmdClose.Top = 300 ''120
            cmdSettings.Left = cmdClose.Left
            
            CurrX = picFrame.Width
            CurrY = picFrame.Height
'''            cmdClose.Left = Me.ScaleWidth - 120 - cmdClose.Width
'''            cmdClose.Top = 120
            cmdFPApproveHide.Top = 120
            cmdFPApproveHide.Left = cmdClose.Left - cmdFPApproveHide.Width - 60
            
            imgLock.Left = cmdClose.Left - imgLock.Width - 60
            lblLock.Left = imgLock.Left + imgLock.Width - lblLock.Width
            Select Case Me.WindowState
                Case 0: mnuMax.Visible = True
                Case 2: mnuMax.Visible = False
            End Select
            
            cmdFPS.Width = flx1.ColWidth(3) * 8
            cmdFPS.Left = flx1.Left + flx1.ColWidth(0) + flx1.ColWidth(1)
'''            cboCUNO.Width = (flx1.Left + flx1.ColWidth(0) + flx1.ColWidth(1)) - cboCUNO.Left
            tvw2.Width = cmdFPS.Left - tvw2.Left - 120
            
            AppWindowState = Me.WindowState
        End If
    End If
End Sub

Private Sub imgComm_Click()
    With frmComments
        .PassREFID = lDWGID
        .PassTable = "DWG_MASTER"
        .PassIType = 0
        .PassBCC = BCC
        .PassFBCN = FBCN
        .PassSHCD = SHCD
        .PassSHYR = SHYR
        .PassMessPath = lblWelcome.Caption
        .PassForm = "frmAnnotator"
        .PassDPath = RelativePath(0)
        .Show 1
    End With
End Sub

Private Sub cmdDirs_Click()
    If sst1.Visible = False Then
        If picRelatives.Visible = True Then
            RelOpen = True
            picRelatives.Visible = False
        Else
            RelOpen = False
        End If
        sst1.Visible = True
        imgComm.Visible = False
        lblWelcome.Visible = False
        lblReds.Visible = False
        cmdDirs.Caption = "Close File Index"
    Else
        sst1.Visible = False
        cmdDirs.Caption = "Open File Index..."
        If bPerm(17) Then imgComm.Visible = True
        lblWelcome.Visible = True
        lblReds.Visible = True
        If RelOpen = True Then
            If bPerm(2) Then picRelatives.Visible = True
        End If
    End If
    If bMenuButton And volFrame.Visible = True Then cmdMenu.Visible = True
End Sub

'''''Private Sub imgConst_Click()
'''''    With frmConst
'''''        .PassBCC = BCC
'''''        .PassFBCN = FBCN
'''''        .PassSHYR = SHYR
'''''        .PassSHCD = SHCD
'''''        .PassSHNM = SHNM
'''''        .Show 1
'''''    End With
'''''End Sub

Private Sub imgDWF_Click(Index As Integer)
    Dim GotIt As Boolean, ViewOK As Boolean
    Screen.MousePointer = 11
    Select Case Index
        Case 0: LoadIt
        Case Else: LoadChild (Index)
    End Select
    ViewOK = InitialView
    Screen.MousePointer = 0
End Sub

Private Sub imgFPApprove_Click()
    picFPApprove.Visible = False
End Sub

Private Sub imgFPS_Click()
    With frmFPStatus
        .PassBCC = BCC
        .PassFBCN = FBCN
        .Show 1
    End With
End Sub

'''''Private Sub imgGraphics_Click()
'''''    With frmGraphics
'''''        .PassBCC = BCC
'''''        .PassFBCN = FBCN
'''''        .PassSHYR = SHYR
'''''        .PassSHCD = SHCD
'''''        .PassSHNM = SHNM
'''''        .Show 1
'''''    End With
'''''End Sub

Private Sub imgIcon_Click(Index As Integer)
    Dim sNode As String, sParent As String, sSuff As String
    Dim nodX As Node
    Dim iFile As Integer, i As Integer, iNode As Integer, iNo As Integer
    If Index < 3 Then
        lblReds.Caption = "": lblReds.Visible = False
        lblWelcome.Caption = "...Ready for your selection..."
        
        PopClients
        
        chkClose.Visible = True
        
        '///// SET VARIABLE CUNOLIST \\\\\\
        CunoList = ""
        For i = 0 To lstClientSorter.ListCount - 1
            Select Case i
                Case 0
                    CunoList = CStr(lstClientSorter.ItemData(i))
                Case Else
                    CunoList = CunoList & ", " & CStr(lstClientSorter.ItemData(i))
            End Select
        Next i
        
        Screen.MousePointer = 0
    End If
Exit Sub
PathNotFound:
    Screen.MousePointer = 0
    MsgBox "Error: " & Err.Description, vbCritical, "Error connecting to File Server"
    Err.Clear
End Sub

Private Sub imgInfo_Click()
    With frmIGLAssignment
        .PassBCC = BCC
        .PassFBCN = FBCN
        .PassSHYR = SHYR
        .PassSHCD = SHCD
        .PassSHNM = SHNM
        .Show 1
    End With
End Sub

Private Sub imgLock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu mnuLockPop, 8, imgLock.Left + imgLock.Width, imgLock.Top + imgLock.Height
End Sub

Private Sub imgMinMax_Click()
    picFPApprove.Visible = False
    cmdFPApproveHide.Visible = True
'    Select Case picFPApprove.Height
'        Case 2595
'            picFPApprove.Height = 390
'            imgMinMax.Picture = imgMax.Picture
'            imgMinMax.ToolTipText = "Click to Maximize"
'        Case 390
'            picFPApprove.Height = 2595
'            imgMinMax.Picture = imgMin.Picture
'            imgMinMax.ToolTipText = "Click to Minimize"
'    End Select
End Sub

Private Sub imgTeam_Click()
    With frmEmailTeam
        .PassBCC = BCC
        .PassFBCN = FBCN
        .Show 1
    End With
End Sub

Private Sub lblClient_Click()
    Dim iFile As Integer, i As Integer
    
    bClient = False
    lblClient.Caption = "Select Client"
    lblClient.ToolTipText = ""
    cmdNav(0).Enabled = False
    cmdNav(1).Enabled = False
    For i = 0 To 19
        lblFile(i).Caption = "": lblExt(i).Caption = ""
        lblFile(i).ForeColor = RGB(150, 150, 102)
    Next
End Sub

Private Sub lblFile_Click(Index As Integer)
    Dim sNode As String, sCommChk As String
    Dim iYear As Integer, i As Integer, iDwg As Integer, iFile As Integer, iDash As Integer, iSearch As Integer
    Dim Resp As VbMsgBoxResult
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iDWFType As Integer
    
    Screen.MousePointer = 11
    
    '///// CHECK IF RED SHOULD BE SAVED \\\\\
    If SaveRed = True Then
        Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNoCancel, "Redline Changes...")
        If Resp = vbYes Then
            mnuRedSave_Click
        ElseIf Resp = vbNo Then
            volFrame.ClearMarkup
'''            SaveRed = False
        ElseIf Resp = vbCancel Then
            GoTo CancelNewFile
        End If
    End If
    
    '///// FIRST SHUT OFF ALL RELATIVES \\\\\
    For i = 0 To 5
        imgDWF(i).Visible = False
    Next i
    cmdOther.Visible = False
    
    RedFile = "": lRedID = 0
    SHCD = CLng(lblFile(Index).Tag)
    SHNM = lblFile(Index).Caption
    strSelect = "SELECT M.DWGID, DWF.SHTID, DWF.DWFID, DWF.DWFTYPE, DWF.DWFPATH " & _
                "From " & DWGShow & " SHO, " & DWGMas & " M, " & _
                "" & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "WHERE SHO.DWGID = M.DWGID  " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "AND SHO.SHYR = " & SHYR & " " & _
                "AND SHO.AN8_SHCD = " & SHCD & " " & _
                "AND SHO.AN8_CUNO = " & CLng(BCC) & " " & _
                "AND DWF.DWFTYPE >= 0 " & _
                "AND DWF.DWFTYPE < 20 " & _
                "AND DWF.DWFSTATUS > 0 " & _
                "ORDER BY DWF.DWFTYPE, DWF.DWFDESC"
                
    Set rst = Conn.Execute(strSelect)
    
    If Not rst.EOF Then
        nLockRefID = rst.Fields("DWGID")
        lDWGID = nLockRefID
        lSHTID = rst.Fields("SHTID")
        Do While Not rst.EOF
            Select Case rst.Fields("DWFTYPE")
                Case 0
                    If Dir(Trim(rst.Fields("DWFPATH")), vbNormal) = "" Then
                        rst.Close
                        Set rst = Nothing
                        MsgBox "File not Found", vbExclamation, "Error Encountered..."
                        picFrame.Visible = True
                        volFrame.Visible = False
                        volFrame.src = ""
                        GoTo CancelNewFile
                    End If
                    imgDWF(rst.Fields("DWFTYPE")).Visible = True
                    RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                    
                        
                Case 1, 2
                    If bPerm(3) Then
                        imgDWF(rst.Fields("DWFTYPE")).Visible = True
                        RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                    End If
                Case 5
                    If bPerm(6) Then
                        imgDWF(rst.Fields("DWFTYPE")).Visible = True
                        RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                    End If
                Case 8
                    If bPerm(3) Then cmdOther.Visible = True
                Case 9
                    RedFile = Trim(rst.Fields("DWFPATH"))
                    lRedID = rst.Fields("DWFID")
                    '*** REDLINE NOTE DEALT WITH @ "LOADIT" ***
            End Select
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    strSelect = "SELECT M.DWGID, M.DWGTYPE, DWF.DWFPATH " & _
                "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "Where SHO.SHYR = " & SHYR & " " & _
                "AND SHO.AN8_SHCD = " & SHCD & " " & _
                "AND SHO.DWGID = M.DWGID " & _
                "AND M.DWGTYPE IN (3, 4, 5) " & _
                "AND M.DSTATUS > 0 " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND M.DWGID = DWF.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "ORDER BY M.DWGTYPE"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case rst.Fields("DWGTYPE")
            Case 3: If bPerm(4) Then imgDWF(3).Visible = True
            Case 4: If bPerm(5) Then imgDWF(4).Visible = True
            Case 5: If bPerm(6) Then imgDWF(5).Visible = True
        End Select
        RelativePath(rst.Fields("DWGTYPE")) = Trim(rst.Fields("DWFPATH"))
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    If bPerm(7) Then
        '///// CHECK FOR CLIENT-SHOW RECORD \\\\\
        strSelect = "SELECT CSY56SHCD FROM " & F5611 & " " & _
                    "WHERE CSY56CUNO = " & CLng(BCC) & " " & _
                    "AND CSY56SHYR = " & SHYR & " " & _
                    "AND CSY56SHCD = " & SHCD
        Set rst = Conn.Execute(strSelect)
        If rst.EOF Then
            imgInfo.Visible = False
        Else
            imgInfo.Visible = True
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    '///// SET NON-VISIBLE TO "" \\\\\
    For i = 0 To 5
        If imgDWF(i).Visible = False Then RelativePath(i) = ""
    Next i
    
    sCurrShow = lblFile(Index)
    ClearChecks
    bLoading = True
    mnuZoomDMode.Checked = True
    sZMode = "Zoom"
    volFrame.UserMode = sZMode
    LoadIt
    lblWelcome.Caption = lblClient & " - " & SHYR & " " & sCurrShow
    If InitialView Then
        volFrame.Visible = True
        picFrame.Visible = False
        bComm = False
        bLoading = False
        lblReds.Visible = True
        
        If bPerm(2) Then picRelatives.Visible = True
        RelOpen = True
        
    
        If bMenuButton And chkClose.value = 1 Then cmdMenu.Visible = True
        
        '///// LET'S CHECK FOR EXISTING COMMENTS \\\\\
        If bTeam Then
            strSelect = "SELECT COMMID " & _
                        "FROM " & ANOComment & " " & _
                        "WHERE REFID = " & lDWGID & " " & _
                        "AND COMMSTATUS > 0"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then '/// COMMENTS EXIST \\\
                imgComm.Picture = imgMail(1).Picture
                imgComm.ToolTipText = "There are saved Comments! Click to access."
            Else '/// NO COMMENTS \\\
                imgComm.Picture = imgMail(0).Picture
                imgComm.ToolTipText = "There are no saved Comments."
            End If
            imgComm.Enabled = True
        Else '/// NO EMAIL TEAM \\\
            imgComm.Picture = imgMail(2).Picture
            imgComm.Enabled = False
        End If
        
        If bPerm(17) Then imgComm.Visible = True
'''        bPicLoaded = True
'''        picDirs.Refresh
    End If
CancelNewFile:
    Screen.MousePointer = 0
End Sub

Private Sub lblFile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To 19
        If i = Index Then lblFile(i).ForeColor = vbWhite Else lblFile(i).ForeColor = RGB(150, 150, 102)
    Next i
End Sub

Private Sub lblFPApprove_DblClick()
    picFPApprove.Top = volFrame.Top + 120
    picFPApprove.Left = volFrame.Left + volFrame.Width - 90 - picFPApprove.Width
End Sub

Private Sub lblFPApprove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Debug.Print "XStr=" & picFPApprove.Left & "    YStr=" & picFPApprove.Top
    Debug.Print "xStr=" & X & "    yStr=" & Y
    xStr = X: yStr = Y
    picFPApprove.Drag 1
End Sub

Private Sub mnuComments1_Click()
    imgComm_Click
End Sub

Private Sub mnuDisplay_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        If i = Index Then mnuDisplay(i).Checked = True Else mnuDisplay(i).Checked = False
    Next i
    Select Case Index
    Case 0
        volFrame.GeometryColor = "DefaultColors"
        volFrame.BACKGROUNDCOLOR = "DefaultColors"
    Case 1
        volFrame.GeometryColor = vbBlack
        volFrame.BACKGROUNDCOLOR = vbWhite
    Case 2
        volFrame.GeometryColor = "ClearScale"
        volFrame.BACKGROUNDCOLOR = "ClearScale"
    End Select
End Sub

Private Sub mnuDownload_Click()
    With frmDownload
        .PassBCC = BCC
        .PassFBCN = FBCN
        .PassSHYR = SHYR
        .PassSHCD = SHCD
        .PassSHNM = SHNM
        .PassDWGID = lDWGID
        .PassFILETYPE = "DWF"
        .Show 1
    End With
End Sub

Private Sub mnuDownloadPDF_Click()
    With frmDownload
        .PassBCC = BCC
        .PassFBCN = FBCN
        .PassSHYR = SHYR
        .PassSHCD = SHCD
        .PassSHNM = SHNM
        .PassDWGID = lDWGID
        .PassFILETYPE = "PDF"
        .Show 1
    End With
End Sub


Private Sub mnuEmailPDF_Click()
'''    frmEmailFile.Show 1
    With frmEmailFile
        .PassFrom = Me.Name
        .PassBCC = BCC
        .PassFBCN = FBCN
        .PassSHYR = SHYR
        .PassSHCD = SHCD
        .PassSHNM = SHNM
        .PassDWGID = lDWGID
        .PassFILETYPE = "PDF"
        .PassSHDT = sShowDates
        .Show 1
    End With
End Sub

Private Sub mnuGFX_Click()
    Screen.MousePointer = 11
    frmPhoto.PassLink = sLinkID
    frmPhoto.PassIn = "2, 3"
    frmPhoto.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub mnuHide_Click()
    cmdMenu.Visible = False
    bMenuButton = False
    If picRelatives.Left = cmdMenu.Left And _
                picRelatives.Top = cmdMenu.Top + cmdMenu.Height + 120 Then
        picRelatives.Top = cmdMenu.Top
    End If
    mnuMenuButton1.Visible = True
End Sub

Private Sub mnuLayers_Click()
    On Error Resume Next
    volFrame.ShowLayersDialog
End Sub

Private Sub mnuLockRemove_Click()
    Dim Resp As VbMsgBoxResult
    If Dir(LCKFile, vbNormal) <> "" Then
        Resp = MsgBox("An email alert will be sent to " & LCKFullName & "," & vbCr & _
                    "notifying that the Lock is being removed." & vbCr & vbCr & _
                    "Are you certain you want to reset the lock?", _
                    vbExclamation + vbYesNo, "Notification...")
        If Resp = vbYes Then
            '///// SEND EMAIL HERE \\\\\
            LockAlert
            KillLocks
            CheckForLock
            MsgBox "The Lock has been removed.", vbExclamation, "Lock Status..."
        End If
    Else
        mnuLockStatus_Click
    End If
End Sub

Private Sub mnuLockStatus_Click()
    Dim sDuration As String
    Dim CHKTime As Date
    Dim rstT As ADODB.Recordset
    Dim strSelect As String
    
    CheckForLock
    If imgLock.Visible = True Then
        strSelect = "SELECT SYSDATE FROM DUAL"
        
        Set rstT = Conn.Execute(strSelect)
        
        CHKTime = rstT.Fields("SYSDATE")
        rstT.Close: Set rstT = Nothing
        sDuration = Int((CHKTime - LCKTime) * 24) & " hrs & " & _
                    CInt((((CHKTime - LCKTime) * 24) - Int((CHKTime - LCKTime) * 24)) * 60) & " minutes"
        MsgBox "This file was opened by " & LCKFullName & ", on " & format(LCKTime, "DDDD") & _
                    ", " & format(LCKTime, "DD-MMM-YYYY") & " at " & _
                    format(LCKTime, "H:NN AM/PM") & " (Det).  (Duration: " & _
                    sDuration & ")" & vbCr & vbCr & _
                    "To remove the Lock, right-click the Lock Icon, but be aware: " & _
                    "Removing the Lock during a current session could result in lost input.", _
                    vbInformation, "Current File Status: IN USE"
    Else
        MsgBox "The Lock has been removed, because the file is no longer in use." & vbCr & vbCr & _
                    "Be aware that a Redline File may have been created by the previous user.  " & _
                    "Please, verify prior to saving any Redline Annotations.", _
                    vbExclamation, "Lock no longer in place..."
    End If
End Sub

Private Sub mnuLogistical_Click()
    On Error Resume Next
    frmLogistics.PassFrom = "FP"
    frmLogistics.Show 1, Me
End Sub

Private Sub mnuMax_Click()
    mnuMax.Visible = False
    Me.WindowState = 2
End Sub

Private Sub mnuMenuButton1_Click()
    cmdMenu.Visible = True
    bMenuButton = True
    If picRelatives.Visible = True And _
                picRelatives.Left < cmdMenu.Left + cmdMenu.Width + 120 And _
                picRelatives.Top < cmdMenu.Top + cmdMenu.Height + 120 Then
        picRelatives.Top = cmdMenu.Top + cmdMenu.Height + 120
        picRelatives.Left = cmdMenu.Left
    End If
    mnuMenuButton1.Visible = False

End Sub

Private Sub mnuPanMode_Click()
    ClearChecks
    mnuPanMode.Checked = True
    volFrame.UserMode = "Pan"
End Sub

Private Sub mnuPhoto_Click()
    Screen.MousePointer = 11
    frmPhoto.PassLink = sLinkID
    frmPhoto.PassIn = "1"
    frmPhoto.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub mnuPrint_Click()
    Dim sMess As String
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(True)
    If Not bPRINTER_ALERT Then Exit Sub
    If bENABLE_PRINTERS Then
        volFrame.ShowPrintDialog
    Else
        sMess = "The Annotator has not found a compatible Printer Driver.  " & _
                    "Printing is temporarily disabled.  " & vbNewLine & vbNewLine & _
                    "NOTE: At times, depending on your printer setup, it may take " & _
                    "a few moments for your printers to be recognized by the server.  " & _
                    "Please, wait a few moments and reselect the 'Print...' option " & _
                    "to check if the Printer mapping has completed yet." & vbNewLine & vbNewLine & _
                    "If this persists, go to the 'Printer Drivers...' " & _
                    "interface on the 'Options' menu ('Key' icon) to submit the issue."
        MsgBox sMess, vbExclamation, "Print Issue..."
    End If
End Sub

Private Sub mnuPrintScreen_Click()
    Dim pScaleX As Long, pScaleY As Long, lXStart As Long, lYStart As Long
    Dim pAspect As Single
    Dim format As Integer
    Dim Annotations As Boolean
    
    On Error Resume Next
    pAspect = picFrame.Width / picFrame.Height
    If pAspect >= 1 Then
        Printer.Orientation = 2
        If pAspect > 1.33 Then
            pScaleX = 10 * 1440
            pScaleY = picFrame.Height / picFrame.Width * pScaleX
        Else
            pScaleY = 7.5 * 1440
            pScaleX = picFrame.Width / picFrame.Height * pScaleY
        End If
        lXStart = (Printer.ScaleWidth - pScaleX) / 2
        lYStart = (Printer.ScaleHeight - pScaleY) / 2
    Else
        Printer.Orientation = 1
        If pAspect < 0.75 Then
            pScaleY = 10 * 1440
            pScaleX = picFrame.Width / picFrame.Height * pScaleY
        Else
            pScaleX = 7.5 * 1440
            pScaleY = picFrame.Height / picFrame.Width * pScaleX
        End If
        lXStart = (Printer.ScaleWidth - pScaleX) / 2
        lYStart = (Printer.ScaleHeight - pScaleY) / 2
    End If

    Printer.PaintPicture picFrame.Image, lXStart, lYStart, pScaleX, pScaleY
    If Err Then
        MsgBox Err.Description
        Err = 0
    Else
        Printer.EndDoc
        If Err Then
            MsgBox "Error: " & Err.Description, vbCritical, "Printer Error"
            Err = 0
        End If
        Printer.Orientation = 1
    End If
End Sub

Private Sub mnuRedClear_Click()
    volFrame.ClearMarkup
End Sub

Private Sub mnuRedDelete_Click()
    Dim Resp As VbMsgBoxResult
    Dim i As Integer
    Dim bOnList As Boolean
    Dim strSelect As String, strDelete As String, sCheck As String
    Dim rst As ADODB.Recordset
    
    If bTeamMember Then
        Resp = MsgBox("Are you certain you want to Permanently delete the Redline File?", _
                    vbCritical + vbYesNoCancel, "Hey...")
        If Resp = vbYes Then
            bViewSet = False
            volFrame.src = RelativePath(0)

            ''CHECK FOR PDF VERSION OF FLOORPLAN OPENNED''
            sCheck = Dir(Left(RelativePath(0), Len(RelativePath(0)) - 3) & "pdf")
            If sCheck <> "" Then
                mnuDownloadPDF.Enabled = True
                mnuEmailPDF.Enabled = True
            Else
                mnuDownloadPDF.Enabled = False
                mnuEmailPDF.Enabled = False
            End If
            
            '///// DELETE REDLINE FROM DATABASE \\\\\
            strDelete = "DELETE FROM " & DWGDwf & " " & _
                        "WHERE DWFID = " & lRedID
            
            Conn.Execute (strDelete)
                
            
            '///// DELETE ACTUAL DWF FILE \\\\\
            Kill RedFile
            
            '///// NOW, CLEAN UP \\\\\
            lRedID = 0
            RedFile = ""
            lblReds.Caption = "NO Redline File exists for this Floor Plan."
            mnuRedLoad.Enabled = False
            mnuRedDelete.Enabled = False
            mnuRedReturn.Enabled = False
            If volFrame.UserMode = "Sketch" Or volFrame.UserMode = "Text" Then
                mnuRedClear.Enabled = True
                mnuRedSave.Enabled = True
            Else
                mnuRedClear.Enabled = False
                mnuRedSave.Enabled = False
            End If
        End If
    Else
        MsgBox "You do not have permission to delete this file." & vbNewLine & _
                "To delete Redline Files, you must be a member" & vbNewLine & _
                "of the Email Notification Team for this Client.", vbCritical, "Sorry..."
    End If
End Sub

Private Sub mnuRedLoad_Click()
    Screen.MousePointer = 11
    bViewSet = False
    volFrame.src = RedFile
    volFrame.Update
    bViewSet = False
    volFrame.src = sCurrPath & RedFile
    lblReds.Caption = "Redline File Loaded."
    mnuRedSave.Enabled = True
    mnuRedClear.Enabled = True
    mnuRedDelete.Enabled = True
    mnuRedReturn.Enabled = True
    
    mnuDownloadPDF.Enabled = False
    mnuEmailPDF.Enabled = False
    
    Screen.MousePointer = 0
End Sub

Private Sub mnuRedMode_Click()
    Dim Resp As VbMsgBoxResult
    Dim sCheck As String
    
    If RedFile = "" Or volFrame.src = RedFile Then
        ClearChecks
        mnuRedMode.Checked = True
        mnuTextMode.Checked = False
        mnuRedModeStop.Enabled = True
        volFrame.UserMode = "Sketch"
        mnuRedSave.Enabled = True
        mnuRedClear.Enabled = True
    ElseIf volFrame.src <> RedFile Then
        Resp = MsgBox("A Redline File already exists, but is not currently loaded." & _
                    vbCr & "Select 'YES' to begin a New Redline, Select 'NO' to Abort." & _
                    vbCr & vbCr & "NOTE: Original Redline will not be overwritten until 'Saved'.", _
                    vbYesNo + vbCritical + vbDefaultButton2, "Existing Redline File...")
        If Resp = vbYes Then
            ClearChecks
            mnuRedMode.Checked = True
            mnuTextMode.Checked = False
            volFrame.UserMode = "Sketch"
            mnuRedSave.Enabled = True
            mnuRedClear.Enabled = True
        End If
    End If
End Sub

Private Sub mnuRedModeStop_Click()
    volFrame.UserMode = sZMode
    mnuRedMode.Checked = False
    mnuTextMode.Checked = False
    Select Case sZMode
        Case "Zoom": mnuZoomDMode.Checked = True
        Case "ZoomToRect": mnuZoomWMode.Checked = True
    End Select
    mnuRedModeStop.Enabled = False
End Sub

Private Sub mnuRedReturn_Click()
    Screen.MousePointer = 11
    LoadIt
    Screen.MousePointer = 0
End Sub

Private Sub mnuRedSave_Click()
    Dim Resp As VbMsgBoxResult
    Dim rstL As ADODB.Recordset, rst As ADODB.Recordset
    Dim lDWFID As Long
    Dim strInsert As String, strUpdate As String
    
    Err = 0
    On Error Resume Next
    Conn.BeginTrans
    If RedFile = "" Then
        '///// WRITE NEW ENTRY \\\\\
        Set rstL = Conn.Execute("SELECT " & DWGSeq & ".NEXTVAL FROM DUAL")
        lDWFID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        strInsert = "INSERT INTO " & DWGDwf & " " & _
                    "(DWGID, SHTID, DWFID, DWFTYPE, " & _
                    "DWFDESC, DWFPATH, DWFSTATUS, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lDWGID & ", " & lSHTID & ", " & lDWFID & ", 9, " & _
                    "'REDLINE', '" & sDWFPath & CStr(lDWFID) & ".dwf" & "', 1, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
        RedFile = sDWFPath & CStr(lDWFID) & ".dwf"
        lRedID = lDWFID
    Else
        strUpdate = "UPDATE " & DWGDwf & " " & _
                    "SET UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE DWGID = " & lDWGID & " " & _
                    "AND SHTID = " & lSHTID & " " & _
                    "AND DWFID = " & lRedID
        Conn.Execute (strUpdate)
        
    End If
    volFrame.SaveMarkup (RedFile)
    If Err Then
        Conn.RollbackTrans
        MsgBox "Redline Annotation cannot be saved." & vbNewLine & vbNewLine & _
                    "Error:  " & Err.Description, vbExclamation, "Error Encountered..."
    Else
        Conn.CommitTrans
        lblReds.Caption = "A Redline File exists for this Floor Plan.  " & _
                    "To Load it, use the right-click menu."
        bReded = True
'''        If Not bLoading Then
            redBCC = BCC: redSHCD = SHCD: redSHYR = SHYR
'''        End If
        SaveRed = False
        lblReds.Visible = True
        mnuRedLoad.Enabled = True
    End If
End Sub

Private Sub mnuSendALink_Click()
'''    MsgBox "CUNO = " & CLng(BCC) & vbNewLine & _
'''            "SHYR = " & SHYR & vbNewLine & _
'''            "SHCD = " & SHCD
    frmSendALink.PassBCC = CLng(BCC)
    frmSendALink.PassSHYR = SHYR
    frmSendALink.PassSHCD = SHCD
    frmSendALink.PassFrom = "FP"
    frmSendALink.PassSub = "AnnoLink: FP - " & FBCN & " - " & SHYR & " " & SHNM
    frmSendALink.Show 1, Me
    
End Sub

Private Sub mnuTextMode_Click()
    If mnuTextMode.Checked = False Then
        ClearChecks
        mnuRedMode.Checked = False
        mnuTextMode.Checked = True
        mnuRedModeStop.Enabled = True
        volFrame.UserMode = "Text"
    Else
        volFrame.UserMode = "Sketch"
        mnuTextMode.Checked = False
        mnuRedMode.Checked = True
        mnuRedModeStop.Enabled = True
    End If
End Sub

Private Sub mnuUsage_Click()
    On Error Resume Next
    frmGantt.PassLink = sLinkID
    frmGantt.Show 1, Me
End Sub

Private Sub mnuZoomDMode_Click()
    ClearChecks
    mnuZoomDMode.Checked = True
    sZMode = "Zoom"
    volFrame.UserMode = sZMode
    mnuRedModeStop.Enabled = False
End Sub

Private Sub mnuZoomFull_Click()
    If volFrame.UserMode = "Pan" Then
        ClearChecks
        mnuZoomWMode.Checked = True
        sZMode = "ZoomToRect"
        volFrame.UserMode = sZMode
    End If
    volFrame.SetCurrentView dLeft, dRight, dBottom, dTop
End Sub

Private Sub mnuZoomWMode_Click()
    ClearChecks
    mnuZoomWMode.Checked = True
    sZMode = "ZoomToRect"
    volFrame.UserMode = sZMode
    mnuRedModeStop.Enabled = False
End Sub

Private Sub optFPApprove_Click(Index As Integer)
    Call CheckIfReadyToApprove(Index)
    
End Sub

Private Sub optSort_Click(Index As Integer)
    Call PopFloorplans(SHYR, CLng(BCC))
End Sub

Private Sub picRelatives_DblClick()
    If bMenuButton Then
        picRelatives.Left = volFrame.Left + 120
        picRelatives.Top = cmdMenu.Top + cmdMenu.Height + 120
    Else
        picRelatives.Left = volFrame.Left + 120
        picRelatives.Top = volFrame.Top + 120
    End If
End Sub

Private Sub picRelatives_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "XStr=" & picRelatives.Left & "    YStr=" & picRelatives.Top
    Debug.Print "xStr=" & X & "    yStr=" & Y
    xStr = X: yStr = Y
    picRelatives.Drag 1
End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, iDwg As Integer, iDash As Integer
    Dim sNode As String, sCheckPath As String
    Dim FirstTime As Boolean
    
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    Select Case UCase(Left(Node.key, 1))
        Case "Y"
            SHYR = Node.Text
            lblClient.Caption = Node.Parent.Text
            BCC = Mid(Node.Parent.key, 2, 8)
            FBCN = Node.Parent.Text
            If bPerm(22) Then imgFPS.Visible = True Else imgFPS.Visible = False
            iDwg = 0
            lstDwgSorter.Clear
'''''            strSelect = "SELECT SHO.AN8_SHCD, SM.SHY56NAMA, " & _
'''''                        "DWF.DWFPATH, M.DWGID " & _
'''''                        "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & F5601 & " SM, " & _
'''''                        "" & DWGDwf & " DWF, " & DWGSht & " SHT " & _
'''''                        "WHERE SHO.AN8_CUNO = " & CLng(BCC) & " " & _
'''''                        "AND SHO.SHYR = " & SHYR & " " & _
'''''                        "AND SHO.DWGID = M.DWGID " & _
'''''                        "AND M.DWGTYPE = 0 " & _
'''''                        "AND M.DSTATUS > 0 " & _
'''''                        "AND M.DWGID = SHT.DWGID " & _
'''''                        "AND M.DWGID = DWF.DWGID " & _
'''''                        "AND SHT.SHTID = DWF.SHTID " & _
'''''                        "AND DWF.DWFTYPE = 0 " & _
'''''                        "AND SHO.SHYR = SM.SHY56SHYR " & _
'''''                        "AND SHO.AN8_SHCD = SM.SHY56SHCD " & _
'''''                        "ORDER BY UPPER(SHY56NAMA)"
            strSelect = "SELECT SHO.AN8_SHCD, AB.ABALPH, " & _
                        "DWF.DWFPATH, M.DWGID " & _
                        "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & F0101 & " AB, " & _
                        "" & DWGDwf & " DWF, " & DWGSht & " SHT " & _
                        "WHERE SHO.AN8_CUNO = " & CLng(BCC) & " " & _
                        "AND SHO.SHYR = " & SHYR & " " & _
                        "AND SHO.DWGID = M.DWGID " & _
                        "AND M.DWGTYPE = 0 " & _
                        "AND M.DSTATUS > 0 " & _
                        "AND M.DWGID = SHT.DWGID " & _
                        "AND M.DWGID = DWF.DWGID " & _
                        "AND SHT.SHTID = DWF.SHTID " & _
                        "AND DWF.DWFTYPE = 0 " & _
                        "AND SHO.AN8_SHCD = AB.ABAN8 " & _
                        "ORDER BY UPPER(AB.ABALPH)"
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If iDwg < 20 Then
'''                    lblFile(iDwg) = DblAmp(UCase(Trim(rst.FIELDS("ABALPH"))))
                    lblFile(iDwg) = UCase(Trim(rst.Fields("ABALPH")))
                    lblFile(iDwg).Tag = CStr(rst.Fields("AN8_SHCD"))
                    lblFile(iDwg).ForeColor = RGB(150, 150, 102)
                End If
'''                lstDwgSorter.AddItem DblAmp(UCase(Trim(rst.FIELDS("ABALPH"))))
                lstDwgSorter.AddItem UCase(Trim(rst.Fields("ABALPH")))
                lstDwgSorter.ItemData(lstDwgSorter.NewIndex) = rst.Fields("AN8_SHCD")
                iDwg = iDwg + 1
                rst.MoveNext
            Loop
            
            cmdNav(0).Enabled = False
            If iDwg < 19 Then
                For iDwg = iDwg To 19
                    lblFile(iDwg) = ""
                    lblFile(iDwg).Tag = ""
                    lblFile(iDwg).ForeColor = RGB(150, 150, 102)
                Next iDwg
            Else: cmdNav(1).Enabled = True
            End If
    End Select
    
    iFileStart = 0
    
    PickPic = True
'''    picDirs.Refresh
'''    imgComm.Visible = False
'''    bTeam = False
    
    '///// TURN OFF NAV CONTROLS IF NO FILES \\\\\
    If lblFile(0).Caption <> "" Then
        cmdNav(0).Visible = True: cmdNav(1).Visible = True: chkClose.Visible = True
    Else
        cmdNav(0).Visible = False: cmdNav(1).Visible = False: chkClose.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim dExeDate As Date, dVerDate As Date
    Dim myTZ As TIME_ZONE_INFORMATION
    Dim dl&
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    
    On Error Resume Next
    bAnnoOpen = True
    CommentSaved = True
    FirstTime = True
    Me.Height = 9000
    Me.Width = 12000
    
    '///// SET TIME OFFSET \\\\\
    dl& = GetTimeZoneInformation(myTZ)
    
    TimeOff = CLng(CInt(myTZ.Bias / 30) / 2)
    
    For i = 0 To 19
        lblFile(i).ForeColor = RGB(150, 150, 102)
    Next i
    
    sDWFPath = "\\Detnovfs2\data1\GPJAnnotator\Floorplans\"
    
    NewUse = True
    
    lblWelcome.Caption = "Welcome " & LogFirstName & "..." '''''  to the GPJ Annotator."
    Err.Clear
    
    sZMode = "Zoom"
    volFrame.UserMode = sZMode
    volFrame.Top = 675 '''120
    volFrame.Left = 120
    volFrame.Width = Me.Width - 360
    volFrame.Height = Me.Height - 1230 ''' volFrame.Width * 2 / 3
    volFrame.src = ""
    volFrame.Visible = False
    
    picFrame.Top = volFrame.Top
    picFrame.Left = volFrame.Left
    picFrame.Width = volFrame.Width
    picFrame.Height = volFrame.Height
    picFrame.Visible = True
    
    CurrX = picFrame.Width
    CurrY = picFrame.Height
    
    cmdDirs.Left = 120 ''' Screen.Width - cmdDirs.Width - 240
    cmdDirs.Top = 120
'''    picDirs.Left = 120 ''' Screen.Width - picDirs.Width - 240
'''    picDirs.Top = cmdDirs.Top + cmdDirs.Height
    sst1.Left = 120
    sst1.Top = cmdDirs.Top + cmdDirs.Height
    imgComm.Left = 2160
    imgComm.Top = cmdDirs.Top
    imgComm.ToolTipText = "Comment Interface"
    
    cmdMenu.Left = volFrame.Left + 120
    cmdMenu.Top = volFrame.Top + 120
    bMenuButton = True
    picRelatives.Top = cmdMenu.Top + cmdMenu.Height + 120
    picRelatives.Left = cmdMenu.Left
    
    picFPApprove.Top = cmdMenu.Top
    
    lblReds.Top = lblWelcome.Top + 300 ''' volFrame.Top - 60 - lblReds.Height
    lblReds.Left = lblWelcome.Left ''' 120
    
    iSHYR = CInt(format(Now, "YYYY"))
    
    lblByGeorge(0).ForeColor = lGeo_Back '' RGB(30, 30, 21)
    lblByGeorge(1).ForeColor = lGeo_Fore '' RGB(100, 100, 68)
    
    
    CommentSaved = True
    
'''''    '///// SET IBM CLIENT GROUP [I208, I600-I670, I700, I800] \\\\\
'''''    sIBMGroup = "'  I208'"
'''''    For i = 0 To 70
'''''        If i < 10 Then
'''''            sIBMGroup = sIBMGroup & ", '  I60" & CStr(i) & "'"
'''''        Else
'''''            sIBMGroup = sIBMGroup & ", '  I6" & CStr(i) & " '"
'''''        End If
'''''    Next i
'''''    sIBMGroup = sIBMGroup & ", '  I700', '  I800'"
    
'    Call GetYears
'    cboSHYR.Text = iSHYR
    
    Call SizeGrid
    
    lNewLockId = 0
    imgIcon_Click (1)
    
    sst1.TabCaption(0) = ""
    
    
    
    '///// CHECK A FEW PERMISSIONS, AND RESET VIEW \\\\\
'''''    If bPerm(4) Then imgIcon(3).Visible = True Else imgIcon(3).Visible = False '/// SHOW PLANS ICON \\\
'''''    If bPerm(5) Then imgIcon(4).Visible = True Else imgIcon(4).Visible = False '/// VIGNETTES ICON \\\
'''''    If bPerm(22) Then imgFPS.Visible = True Else imgFPS.Visible = False '/// FLOORPLAN STATUS \\\
    If bPerm(20) Then mnuLayers.Visible = True Else mnuLayers.Visible = False '/// LAYER MODE \\\
    
    '///// EDITDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(False)
'''    If bPerm(21) And bENABLE_PRINTERS Then mnuPrint.Visible = True Else mnuPrint.Visible = False '/// PRINT OPTION \\\
    If bPerm(21) Then mnuPrint.Visible = True Else mnuPrint.Visible = False '/// PRINT OPTION \\\
    '\\\\\ -------------------------------------------------------- /////
    
    If bPerm(17) Then mnuComments1.Visible = True Else mnuComments1.Visible = False '/// RIGHTCLICK COMMENTS \\\
    If bPerm(7) Then imgInfo.Visible = True Else imgInfo.Visible = False '/// SHOW INFO \\\
'''''    If bPerm(23) Then imgGraphics.Visible = True Else imgGraphics.Visible = False '/// GRAPHICS \\\
'''''    If bPerm(32) Then imgConst.Visible = True Else imgConst.Visible = False '/// CONSTRUCTION \\\
'''''    If bPerm(13) Then mnuRedlining.Visible = True Else mnuRedlining.Visible = False '/// REDLINING INTERFACE \\\
'''''    If bPerm(14) Then mnuRedlines.Visible = True Else mnuRedlines.Visible = False '/// REDLINE CAPABILTY \\\
    If bPerm(15) Then '/// REDLINE SAVE \\\
        mnuRedSave.Visible = True
        mnuRedClear.Visible = True
        mnuRedMode.Visible = True
        mnuTextMode.Visible = True
    Else
        mnuRedSave.Visible = False
        mnuRedClear.Visible = False
        mnuRedMode.Visible = False
        mnuTextMode.Visible = False
    End If
    If bPerm(16) Then mnuRedDelete.Visible = True Else mnuRedDelete.Visible = False '/// REDLINE DELETE \\\
'''    If bPerm(0) Then imgTeam.Visible = True Else imgTeam.Visible = False '/// TEAM ACCESS \\\
    If bPerm(47) Then mnuDownload.Visible = True Else mnuDownload.Visible = False '/// DOWNLOAD DWFS \\\
    
    ''///// GET sApprClientList \\\\\''
    sApprClientList2 = ""
    strSelect = "SELECT DISTINCT ET.AN8_CUNO " & _
                "FROM " & ANOETeamUR & " ETU, " & ANOETeam & " ET " & _
                "WHERE ETU.USER_SEQ_ID = " & UserID & " " & _
                "AND ETU.TEAM_ID = ET.TEAM_ID"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sApprClientList = "|"
        Do While Not rst.EOF
            sApprClientList = sApprClientList & CStr(rst.Fields("AN8_CUNO")) & "|"
            If sApprClientList2 = "" Then
                sApprClientList2 = CStr(rst.Fields("AN8_CUNO"))
            Else
                sApprClientList2 = sApprClientList2 & ", " & CStr(rst.Fields("AN8_CUNO"))
            End If
            rst.MoveNext
        Loop
    Else
        sApprClientList = ""
    End If
    rst.Close: Set rst = Nothing
    Debug.Print sApprClientList
    If sApprClientList = "" Then
        chkFPApprove.Visible = False
    Else
        chkFPApprove.Visible = True
    End If
    
    ''///// THIS IS A TEST... \\\\\''
    If bPassIn Then
        Dim tSHYR As Integer, iDash As Integer
        Dim tCUNO As String
        tSHYR = CInt(Left(sPassInValue, 4))
        iDash = InStr(6, sPassInValue, "|")
        tCUNO = Right("00000000" & Mid(sPassInValue, 6, iDash - 6), 8)
        Call tvw2_NodeClick(frmAnnotator.tvw1.Nodes("y" & CStr(tSHYR) & "-" & tCUNO))
'''        flx1.RowSel = 3: iFlxCol = 0
'''        Call flx1_Click
    ElseIf defCUNO > 0 Then
        tSHYR = CInt(format(Now, "yyyy"))
        tCUNO = Right("00000000" & CStr(defCUNO), 8)
'''        Call tvw2_NodeClick(frmAnnotator.tvw1.Nodes("y" & CStr(tSHYR) & "-" & tCUNO))
        
        Err.Clear
        Call tvw2_NodeClick(tvw2.Nodes("y" & CStr(tSHYR) & "-" & tCUNO))
        If Err = 0 Then
            If bPerm(22) Then cmdFPSI.Visible = True Else cmdFPSI.Visible = False
            If tvw2.Height <> 1695 Then tvw2.Height = 1695
    '''        tvw2.Nodes("c" & tCUNO).EnsureVisible
            tvw2.Nodes("y" & CStr(tSHYR) & "-" & tCUNO).Selected = True
            tvw2.Nodes("c" & tCUNO).EnsureVisible
        End If
'        imgComm.Visible = False
        sst1.Visible = True
        
        cmdDirs.Caption = "Close File Index"
    End If


    Me.MousePointer = 0
    Me.WindowState = AppWindowState
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sCheck As String, strUpdate As String
    Dim RetVal As Variant
    Dim i As Integer
    Dim Resp As VbMsgBoxResult

     '///// CHECK IF RED SHOULD BE SAVED \\\\\
    If SaveRed = True Then
        Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNo, "Redline Changes...")
        If Resp = vbYes Then
            mnuRedSave_Click
        ElseIf Resp = vbNo Then
            volFrame.ClearMarkup
        End If
    End If
    
    If bReded = True And bTeam = True And bPerm(15) Then
        With frmRedAlert
            .PassSHYR = redSHYR
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 0
            .Show 1
        End With
'''        Call RedAlert(0, lblWelcome, redBCC, redSHCD)
    End If
    bReded = False
    redBCC = "": redSHCD = 0: redSHYR = 0
    
    '///// KILL LOCK FILE, IF ACTIVE \\\\\
    If lNewLockId <> 0 Then
        strUpdate = "UPDATE " & ANOLockLog & " " & _
                    "SET LOCKCLOSEDTTM = SYSDATE, " & _
                    "LOCKSTATUS = LOCKSTATUS * -1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, " & _
                    "UPDCNT = UPDCNT + 1 " & _
                    "WHERE LOCKID = " & lNewLockId & " " & _
                    "AND LOCKREFID = " & lLockRefID
        Conn.Execute (strUpdate)
    End If
    
    ''ADD TEST FOR PASSING IN VARS''
'    BCC = "": SHCD = 0: SHYR = 0: FBCN = "": SHNM = ""

    bAnnoOpen = False
        Unload frmComments
        bCommentsOpen = False
End Sub

Private Sub tvw2_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim SelClause As String
    Dim iDash As Integer, iRow As Integer
    Dim tSHCD As Long
    
    Select Case UCase(Left(Node.key, 1))
        Case "Y"
            fSHYR = Node.Text
            fBCC = Mid(Node.Parent.key, 2, 8)
            fFBCN = Node.Parent.Text
            cmdSettings.Enabled = True
            If bPerm(22) Then cmdFPSI.Visible = True Else cmdFPSI.Visible = False
'            If bApprovalList Then
            Call PopFloorplans(fSHYR, CLng(fBCC))
            If tvw2.Height <> 1695 Then tvw2.Height = 1695
            tvw2.Nodes(Node.key).EnsureVisible
'''            imgComm.Visible = False
'''            bTeam = False
            If bPassIn Then
                bTeam = False
                tSHCD = CLng(Mid(sPassInValue, InStr(6, sPassInValue, "|") + 1))
                For iRow = 1 To flx1.Rows - 1
                    If flx1.TextMatrix(iRow, 10) = tSHCD Then
                        flx1.RowSel = iRow: iFlxCol = 0
                        Call flx1_Click
                        bPassIn = False
                        GoTo JumpOut
                    End If
                Next iRow
JumpOut:
            End If
            fraFP.Visible = True
    End Select
End Sub


Private Sub txtFPApprove_Change()
    If optFPApprove(2).value = True Then
        If txtFPApprove.Text <> "" Then
            cmdFPApprove.Enabled = True
        Else
            cmdFPApprove.Enabled = False
        End If
    End If
End Sub

Private Sub volFrame_DoNavigateToURL(ByVal URL As String, ByVal window_name As String, enable_default As Boolean)
    enable_default = False
    sLinkID = URL
    Call CheckForImages(sLinkID)
    Me.PopupMenu mnuRC
End Sub

Private Sub volFrame_DragDrop(Source As Control, X As Single, Y As Single)
    Debug.Print "Dragging " & Source.Name
    If Source.Name = "picRelatives" Then
        With picRelatives
            .Drag 2
            .Move CSng(volFrame.Left + (X - xStr)), CSng(volFrame.Top + (Y - yStr))
        End With
    ElseIf Source.Name = "picFPApprove" Then
        With picFPApprove
            .Drag 2
            .Move CSng(volFrame.Left + (X - xStr)), CSng(volFrame.Top + (Y - yStr))
        End With
    End If
End Sub

Private Sub volFrame_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuRightClick
    Else
        If volFrame.UserMode = "Sketch" Or volFrame.UserMode = "Text" Then
            SaveRed = True
        End If
    End If
    
End Sub

Public Sub LoadIt()
    Dim sCheck As String
    Dim Resp As VbMsgBoxResult
    Dim strUpdate As String, strSelect As String
    Dim rst As ADODB.Recordset
    
'    If bPerm(14) Then mnuRedlines.Visible = True Else mnuRedlines.Visible = False '/// REDLINE CAPABILTY \\\
    If bPerm(13) Then mnuRedlining.Visible = True Else mnuRedlining.Visible = False '/// REDLINING INTERFACE \\\
    
    
'''    If bReded = True And bTeam = True And bPerm(15) Then
'''        With frmRedAlert
'''            .PassSHYR = redSHYR
'''            .PassBCC = CLng(redBCC)
'''            .PassSHCD = redSHCD
'''            .PassHDR = lblWelcome
'''            .PassType = 0
'''            .Show 1
'''        End With
''''''        Call RedAlert(0, lblWelcome, redBCC, redSHCD) 'AlertOfRed
'''    End If
'''    bReded = False
'''    redBCC = "": redSHCD = 0: redSHYR = 0
    
    CommentSaved = True
    
    '????? WHAT AM I WRITING TO LOG ?????
    If lNewLockId <> 0 Then
        strUpdate = "UPDATE " & ANOLockLog & " " & _
                    "SET LOCKCLOSEDTTM = SYSDATE, " & _
                    "LOCKSTATUS = LOCKSTATUS * -1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, " & _
                    "UPDCNT = UPDCNT + 1 " & _
                    "WHERE LOCKID = " & lNewLockId & " " & _
                    "AND LOCKREFID = " & lLockRefID
        Conn.Execute (strUpdate)
    End If
    
    If chkClose.value = 1 Then
'''        picDirs.Visible = False
        sst1.Visible = False
        cmdDirs.Caption = "Open File Index..."
    End If
    
    bViewSet = False
    
    Err = 0
    On Error Resume Next
    volFrame.src = RelativePath(0)
    If Err Then
        picFrame.Visible = True
        volFrame.Visible = False
        GoTo ChickenOut
    Else
        ''CHECK FOR PDF VERSION OF FLOORPLAN OPENED''
        sCheck = Dir(Left(RelativePath(0), Len(RelativePath(0)) - 3) & "pdf")
        If sCheck <> "" Then
            mnuDownloadPDF.Enabled = True
            mnuEmailPDF.Enabled = True
        Else
            mnuDownloadPDF.Enabled = False
            mnuEmailPDF.Enabled = False
        End If
    End If
    On Error GoTo 0
    
    '????? HOW DO I DEAL WITH LCKFILE ?????
    CheckForLock
    bTeam = CheckForTeam(BCC, SHCD, frmAnnotator)
    
    If bPerm(13) Then
        If RedFile = "" Then
            lblReds.Caption = "NO Redline File exists for this Floor Plan."
            mnuRedLoad.Enabled = False
            mnuRedDelete.Enabled = False
            If volFrame.UserMode = "Sketch" Or volFrame.UserMode = "Text" Then
                mnuRedClear.Enabled = True
                mnuRedSave.Enabled = True
            Else
                mnuRedClear.Enabled = False
                mnuRedSave.Enabled = False
            End If
        Else
            lblReds.Caption = "A Redline File exists for this Floor Plan.  " & _
                        "To Load it, use the right-click menu."
            mnuRedLoad.Enabled = True
            mnuRedDelete.Enabled = False
            mnuRedClear.Enabled = False
            mnuRedSave.Enabled = False
        End If
        mnuRedReturn.Enabled = False
    Else
        lblReds.Caption = "Floorplan View"
    End If
ChickenOut:
End Sub

Public Sub ClearChecks()
    mnuPanMode.Checked = False
    mnuZoomDMode.Checked = False
    mnuZoomWMode.Checked = False
    mnuRedMode.Checked = False
    mnuTextMode.Checked = False
End Sub

Public Function DeQuotate(txt As String) As String
    Dim i As Integer
    Dim strCheck As String
    DeQuotate = ""
    i = Len(txt)
    strCheck = Trim(txt)
    Do While i > 0
        If Mid(strCheck, i, 1) = """" Then
            strCheck = Left(strCheck, i - 1) & "''" & Mid(strCheck, i + 1)
        End If
        i = i - 1
    Loop
    DeQuotate = strCheck
End Function

Public Function DeDblApost(txt As String) As String
    Dim Pos As Integer
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(1, txt, "''")
        If Pos <> 0 Then txt = Left(txt, Pos - 1) & Chr(34) & Mid(txt, Pos + 2)
    Loop
    DeDblApost = txt
End Function

'''Public Function DblAmp(txt As String) As String
'''    Dim Pos As Integer
'''    Pos = 1
'''    Do While Pos <> 0
'''        Pos = InStr(Pos, txt, "&")
'''        If Pos <> 0 Then
'''            txt = Left(txt, Pos - 1) & Chr(38) & Mid(txt, Pos)
'''            Pos = Pos + 2
'''        End If
'''    Loop
'''    DblAmp = txt
'''End Function

Public Function InitialView() As Boolean
    On Error Resume Next
    volFrame.GetCurrentView dLeft, dRight, dBottom, dTop
    If Err Then
        MsgBox "Error Loading Floorplan", vbExclamation, "Error Encountered..."
        picFrame.Visible = True
        volFrame.Visible = False
        volFrame.src = ""
        InitialView = False
    Else
        InitialView = True
    End If
End Function

'''Private Sub volFrame_MouseUp(Button As Integer, Shift As Integer, x As Double, y As Double)
'''    If volFrame.UserMode = "Sketch" Then
'''        volFrame.UserMode = "Text"
'''        volFrame.UserMode = "Sketch"
'''    End If
'''End Sub

Private Sub volFrame_OnClearMarkup(enable_default As Boolean)
    SaveRed = False
End Sub

Private Sub volFrame_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    Dim ViewOK As Boolean
    If bViewSet = False Then
        If StatusCode = 42 Then
            ViewOK = InitialView
            bViewSet = True
        End If
    End If
End Sub

Private Sub volFrame_OnSaveMarkup(enable_default As Boolean)
    enable_default = False
    
    mnuRedSave_Click
End Sub

Public Function DeSlash(sName As String) As String
    Dim Pos As Integer
    '///// DeSLASH \\\\\
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(Pos, sName, "/")
        If Pos > 0 Then
            sName = Left(sName, Pos - 1) & "_" & Mid(sName, Pos + 1)
        End If
    Loop
    DeSlash = sName
End Function

Public Function ReSlash(sName As String) As String
    Dim Pos As Integer
    '///// ReSLASH \\\\\
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(Pos, sName, "_")
        If Pos > 0 Then
            sName = Left(sName, Pos - 1) & "/" & Mid(sName, Pos + 1)
        End If
    Loop
    ReSlash = sName
End Function

'''Public Sub GetPres()
'''    sPrePic = sCurrPic
'''    sPrePath = sCurrPath
'''    sPreShow = sCurrShow
'''
'''    Call RedAlert(0, lblWelcome, redBCC, redSHCD)
'''    bReded = False
'''End Sub

Public Sub LoadChild(Index As Integer)
    Dim sViewSet As String, sCheck As String

    If bPerm(13) Then mnuRedlining.Visible = True Else mnuRedlining.Visible = False '/// REDLINING INTERFACE \\\
'''    If bPerm(14) Then mnuRedlines.Visible = True Else mnuRedlines.Visible = False '/// REDLINE CAPABILTY \\\
    bViewSet = False
    On Error GoTo DrawingNotFound
    
    volFrame.src = RelativePath(Index) ''' "file:" & sDWF(Index)
    
    ''CHECK FOR PDF VERSION OF DRAWING''
    sCheck = Dir(Left(RelativePath(Index), Len(RelativePath(Index)) - 3) & "pdf")
    If sCheck <> "" Then
        mnuDownloadPDF.Enabled = True
        mnuEmailPDF.Enabled = True
    Else
        mnuDownloadPDF.Enabled = False
        mnuEmailPDF.Enabled = False
    End If
    
    Select Case Index
        Case 1: sViewSet = "Isometric View"
        Case 2: sViewSet = "Elevations Drawing"
        Case 3
            sViewSet = "Overall Show Plan"
            mnuRedlining.Visible = False
'''            mnuRedlines.Visible = False
        Case 4
            sViewSet = "Composite Show Plan"
            mnuRedlining.Visible = False
'''            mnuRedlines.Visible = False
        Case 5: sViewSet = "Press Plan Drawing"
        Case Else: sViewSet = ""
    End Select
    lblReds.Caption = sViewSet
Exit Sub
DrawingNotFound:
    MsgBox "Error:  " & Err.Description, vbExclamation, "File Not Found..."
    Err.Clear
End Sub

'''''Public Sub CheckForIBM(sBCC As String)
'''''    bIBM = False
'''''End Sub

Public Sub CheckForLock()
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim LStatus As Integer
    
    
    '///// CHECK TO SEE IF IN USE \\\\\
    strSelect = "SELECT U.NAME_FIRST, U.NAME_LAST, U.EMAIL_ADDRESS, " & _
                "L.LOCKID, L.LOCKOPENDTTM " & _
                "FROM " & ANOLockLog & " L, " & IGLUser & " U " & _
                "WHERE L.LOCKREFID = " & nLockRefID & " " & _
                "AND L.LOCKSTATUS = 2 " & _
                "AND L.USER_SEQ_ID = U.USER_SEQ_ID"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        LCKFullName = Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST"))
        LCKLotusName = Trim(rst.Fields("EMAIL_ADDRESS"))
        LCKTime = rst.Fields("LOCKOPENDTTM")
        lLockID = rst.Fields("LOCKID")
        imgLock.ToolTipText = "Redline Lock: " & LCKFullName & " (" & _
                    format(rst.Fields("LOCKOPENDTTM"), "dd mmm yyyy - h:nn am/pm") & _
                    " Detroit Time)"
        lblLock.Caption = "File in use by " & LCKFullName
        bLock = True
        imgLock.Visible = True
        If bPerm(44) Or UCase(LogName) = UCase(LCKFullName) Then _
            imgLock.Enabled = True Else imgLock.Enabled = False
        LStatus = 1
    Else
        lblLock.Caption = ""
        bLock = False
        lLockID = 0
        imgLock.Visible = False
        LStatus = 2
    End If
    rst.Close
    Set rst = Nothing
    
    '///// ADD ENTRY TO LOCKLOG \\\\\
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    lNewLockId = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing
        
    strInsert = "INSERT INTO " & ANOLockLog & " " & _
                "(LOCKID, LOCKREFID, LOCKREFSOURCE, " & _
                "USER_SEQ_ID, LOCKOPENDTTM, LOCKSTATUS, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                "VALUES " & _
                "(" & lNewLockId & ", " & nLockRefID & ", 'DWG_MASTER', " & _
                UserID & ", SYSDATE, " & LStatus & ", " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
        
    lLockRefID = nLockRefID
    
    If bLock Then mnuRedlining.Enabled = False Else mnuRedlining.Enabled = True
End Sub

Public Sub LockAlert()
    Dim MessBody As String, MessHdr As String
    
    MessHdr = "Annotator Lock being removed!"
    MessBody = "You are documented as currently accessing the following " & _
                "Floor Plan in the GPJ Annotator." & _
                vbCr & vbCr & _
                vbTab & "Client:" & vbTab & FBCN & vbCr & _
                vbTab & "Show:" & vbTab & SHYR & " - " & sCurrShow & _
                vbCr & vbCr & _
                "I am removing the Lock on the file.  " & _
                "Please, contact me immediately if you will be saving any Annotations " & _
                "before leaving the file.  If you and I each save, the Annotations could be overwritten."
        
    '///// EXECUTE E-MAIL \\\\\
    Dim myNotes As New Domino.NOTESSESSION
    Dim myDB As New Domino.NOTESDATABASE
    Dim myItem  As Object ''' NOTESITEM
    Dim myDoc As Object ''' NOTESDOCUMENT
    Dim myRichText As Object ' NOTESRICHTEXTITEM
    Dim myReply  As Object ''' NOTESITEM
    Dim Address As String

    Address = LCKLotusName

    myNotes.Initialize
    
'/// ACTIVATE FOR CITRIX \\\
    Set myDB = myNotes.GETDATABASE("detsrv1/det/GPJNotes", "mail\gannotat.nsf")
'''    Set myDB = myNotes.GETDATABASE("detsrv1/det/GPJNotes", "mail\swesterh.nsf")
    Set myDoc = myDB.CREATEDOCUMENT
    Call myDoc.REPLACEITEMVALUE("Principal", LogName)
    Set myItem = myDoc.APPENDITEMVALUE("Subject", MessHdr)
    Set myReply = myDoc.APPENDITEMVALUE("ReplyTo", LogAddress)
    Set myRichText = myDoc.CREATERICHTEXTITEM("Body")
    With myRichText
        .APPENDTEXT MessBody
        .ADDNEWLINE 2
        .APPENDTEXT LogName
    End With
    myDoc.APPENDITEMVALUE "SENDTO", Address
    myDoc.SAVEMESSAGEONSEND = True
    
    On Error Resume Next
    Call myDoc.SEND(False, Address)
    If Err Then
        MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
                    vbExclamation, "Error Encountered"
        Err = 0
        GoTo GetOut
    End If
GetOut:
    Set myReply = Nothing
    Set myRichText = Nothing
    Set myItem = Nothing
    Set myDoc = Nothing
    Set myDB = Nothing
    Set myNotes = Nothing
End Sub

Public Sub PopClients()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sCNode As String, sYNode As String, sDesc As String
    
    lstClientSorter.Clear
    tvw1.Nodes.Clear
    tvw2.Visible = False
    tvw2.Nodes.Clear
    tvw2.Visible = True
    sCNode = "": sYNode = ""
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT SHO.AN8_CUNO, C.ABALPH, SHO.SHYR " & _
                    "FROM " & DWGShow & " SHO, " & F0101 & " C " & _
                    "WHERE SHO.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY C.ABALPH, SHO.SHYR"
'''        strSelect = "SELECT DISTINCT SHO.AN8_CUNO, C.ABALPH, SHO.SHYR " & _
'''                    "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & _
'''                    "" & DWGDwf & " DWF, " & F0101 & " C, " & IGLKitU & " KU " & _
'''                    "WHERE SHO.DWGID = M.DWGID " & _
'''                    "AND M.DWGTYPE = 0 " & _
'''                    "AND M.DSTATUS > 0 " & _
'''                    "AND M.DWGID = SHT.DWGID " & _
'''                    "AND M.DWGID = DWF.DWGID " & _
'''                    "AND SHT.SHTID = DWF.SHTID " & _
'''                    "AND DWF.DWFTYPE = 0 " & _
'''                    "AND SHO.AN8_CUNO = KU.AN8_CUNO " & _
'''                    "AND SHO.SHYR = KU.SHYR " & _
'''                    "AND SHO.AN8_SHCD = KU.AN8_SHCD " & _
'''                    "AND KU.FPSTATUS > 0 " & _
'''                    "AND SHO.AN8_CUNO = C.ABAN8 " & _
'''                    "ORDER BY C.ABALPH, SHO.SHYR"
    Else
        strSelect = "SELECT DISTINCT SHO.AN8_CUNO, C.ABALPH, SHO.SHYR " & _
                    "FROM " & DWGShow & " SHO, " & F0101 & " C " & _
                    "WHERE SHO.DWGID > 0 " & _
                    "AND SHO.SHOW_ID > 0 " & _
                    "AND SHO.AN8_CUNO IN (" & strCunoList & ") " & _
                    "AND SHO.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAN8 > 0 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY C.ABALPH, SHO.SHYR"
'''        strSelect = "SELECT DISTINCT SHO.an8_CUNO, C.ABALPH, SHO.SHYR " & _
'''                    "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & _
'''                    "" & DWGDwf & " DWF, " & F0101 & " C, " & IGLKitU & " KU " & _
'''                    "WHERE SHO.AN8_CUNO IN (" & strCunoList & ") " & _
'''                    "AND SHO.DWGID = M.DWGID " & _
'''                    "AND M.DWGTYPE = 0 " & _
'''                    "AND M.DSTATUS > 0 " & _
'''                    "AND M.DWGID = SHT.DWGID " & _
'''                    "AND M.DWGID = DWF.DWGID " & _
'''                    "AND SHT.SHTID = DWF.SHTID " & _
'''                    "AND DWF.DWFTYPE = 0 " & _
'''                    "AND SHO.AN8_CUNO = KU.AN8_CUNO " & _
'''                    "AND SHO.SHYR = KU.SHYR " & _
'''                    "AND SHO.AN8_SHCD = KU.AN8_SHCD " & _
'''                    "AND KU.FPSTATUS > 0 " & _
'''                    "AND SHO.AN8_CUNO = C.ABAN8 " & _
'''                    "ORDER BY C.ABALPH, SHO.SHYR"
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sCNode <> "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8) Then
            sCNode = "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
            sDesc = UCase(Trim(rst.Fields("ABALPH")))
            Set nodX = tvw1.Nodes.Add(, , sCNode, sDesc)
            Set nodX = tvw2.Nodes.Add(, , sCNode, sDesc)
            lstClientSorter.AddItem UCase(Trim(rst.Fields("ABALPH")))
            lstClientSorter.ItemData(lstClientSorter.NewIndex) = rst.Fields("AN8_CUNO")
        End If
        
        sYNode = "y" & rst.Fields("SHYR") & "-" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
        sDesc = rst.Fields("SHYR")
        Set nodX = tvw1.Nodes.Add(sCNode, tvwChild, sYNode, sDesc)
        Set nodX = tvw2.Nodes.Add(sCNode, tvwChild, sYNode, sDesc)
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

'Public Sub PopClients2()
'    Dim strSelect As String
'    Dim rst As ADODB.Recordset
'
'    cboCUNO.Clear
'    lstClientSorter.Clear
'    If bClientAll_Enabled Then
'        strSelect = "SELECT DISTINCT KU.AN8_CUNO, AB.ABALPH " & _
'                    "FROM " & IGLKitU & " KU, " & F0101 & " AB " & _
'                    "WHERE KU.FPSTATUS > 0 " & _
'                    "AND KU.SHYR = " & cboSHYR.Text & " " & _
'                    "AND KU.AN8_CUNO = AB.ABAN8 " & _
'                    "ORDER BY UPPER(AB.ABALPH)"
'    Else
'        strSelect = "SELECT DISTINCT KU.AN8_CUNO, AB.ABALPH " & _
'                    "FROM " & IGLKitU & " KU, " & F0101 & " AB " & _
'                    "WHERE KU.AN8_CUNO IN (" & strCunoList & ") " & _
'                    "AND KU.FPSTATUS > 0 " & _
'                    "AND KU.SHYR = " & cboSHYR.Text & " " & _
'                    "AND KU.AN8_CUNO = AB.ABAN8 " & _
'                    "ORDER BY UPPER(AB.ABALPH)"
'    End If
'    Set rst = Conn.Execute(strSelect)
'    Do While Not rst.EOF
'        cboCUNO.AddItem UCase(Trim(rst.Fields("ABALPH")))
'        cboCUNO.ItemData(cboCUNO.NewIndex) = rst.Fields("AN8_CUNO")
'
'
''''        If sCNode <> "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8) Then
''''            sCNode = "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
''''            sDesc = UCase(Trim(rst.Fields("ABALPH")))
''''            Set nodX = tvw1.Nodes.Add(, , sCNode, sDesc)
'            lstClientSorter.AddItem UCase(Trim(rst.Fields("ABALPH")))
'            lstClientSorter.ItemData(lstClientSorter.NewIndex) = rst.Fields("AN8_CUNO")
''''        End If
''''
''''        sYNode = "y" & rst.Fields("SHYR") & "-" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
''''        sDesc = rst.Fields("SHYR")
''''        Set nodX = tvw1.Nodes.Add(sCNode, tvwChild, sYNode, sDesc)
'
'        rst.MoveNext
'    Loop
'    rst.Close: Set rst = Nothing
'End Sub

Public Sub KillLocks()
    Dim strUpdate As String
    
    '///// KILL BOTH EXISTING LOCKS, CHECKFORLOCK WILL CREATE NEW ONE \\\\\
    strUpdate = "UPDATE " & ANOLockLog & " " & _
                "SET LOCKSTATUS = LOCKSTATUS * -1, " & _
                "LOCKCLOSEDTTM = SYSDATE, " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, " & _
                "UPDCNT = UPDCNT + 1 " & _
                "WHERE LOCKID = " & lLockID & " " & _
                "OR LOCKID = " & lNewLockId & " " & _
                "AND LOCKREFID = " & lLockRefID
    Conn.Execute (strUpdate)
    lLockID = 0
End Sub


'Public Sub GetYears()
'    Dim strSelect As String
'    Dim rst As ADODB.Recordset
'
'    cboSHYR.Clear
'    strSelect = "SELECT DISTINCT SHYR " & _
'                "FROM " & IGLKitU & " " & _
'                "WHERE FPSTATUS > 0 " & _
'                "ORDER BY SHYR"
'    Set rst = Conn.Execute(strSelect)
'    Do While Not rst.EOF
'        cboSHYR.AddItem rst.Fields("SHYR")
'        rst.MoveNext
'    Loop
'    rst.Close: Set rst = Nothing
'End Sub

Public Function FillGrid(strSelect As String)
    'Dim Conn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim iRow As Integer, nrow As Integer, ictr As Integer
    Dim i As Integer
    
    flx1.Visible = False
    flx1.Rows = 1
'''    strSelect = SelClause & WhereClause & AndClause
'''    Debug.Print strSelect
    Set rst = Conn.Execute(strSelect)
    iRow = 1
    nrow = 1
    ictr = 0
    Do While Not rst.EOF
        flx1.Rows = flx1.Rows + 1
        flx1.Row = iRow + ictr
        flx1.TextMatrix(flx1.Row, 0) = UCase(Trim(rst.Fields("SHY56NAMA")))
        If rst.Fields("FPSTATUS") < 4 Then
            flx1.Col = 0: flx1.CellForeColor = QBColor(8)
        Else
            flx1.Col = 0: flx1.CellFontBold = True
        End If
        flx1.TextMatrix(flx1.Row, 1) = format(rst.Fields("BEG_DATE"), "mmm d") & " - " & _
                    format(rst.Fields("END_DATE"), "mmm d, yyyy")
        For i = 1 To rst.Fields("FPSTATUS")
            flx1.Col = i + 1
'''            flx1.CellFontName = "Wingdings"
'''            flx1.CellFontSize = 14
'''            flx1.Text = "n"
            flx1.CellBackColor = lColor2 ' vbActiveTitleBar
        Next i
        flx1.CellForeColor = vbTitleBarText
        flx1.Text = format(rst.Fields("FPSTATDT"), "m/d")
        flx1.TextMatrix(flx1.Row, 10) = rst.Fields("SHY56SHCD")
        ictr = ictr + 1
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
            
'''''            If .Fields("SHY56SHCD") = ShoCode Then
'''''                clnt = UCase(Trim(.Fields("ABALPH")))
'''''                cuno = Right("00000000" & CStr(.Fields("CSY56CUNO")), 8)
'''''                fpst = .Fields("FPSTATUS") + 1
'''''                fpsd = format(.Fields("FPSTATDT"), "d-mmm-yy")
'''''                flex1.Col = 1
'''''                flex1.Text = CLng(cuno) & " - " & clnt
'''''                For iCol = 2 To fpst
'''''                    flex1.Col = iCol
'''''                    flex1.CellBackColor = vbActiveTitleBar
'''''                Next
''''''''                Flex1.CellAlignment = 4
'''''                flex1.CellForeColor = vbTitleBarText
'''''                flex1.Text = fpsd
'''''                .MoveNext
'''''            Else
'''''                ShoCode = .Fields("SHY56SHCD")
''''''''                t_shcd = Right("0000" & CStr(ShoCode), 4)
'''''                t_nama = UCase(Trim(.Fields("SHY56NAMA")))
'''''                d_begd = format(DateValue(.Fields("BEG_DATE")), "d-mmm-yy")
''''''                d_endd = Format(.Fields("END_DATE"), "m/d/yy")
'''''                With flex1
'''''                    .Col = 0
''''''''                    .CellAlignment = 4
'''''                    .CellFontBold = True
'''''                    .Text = CStr(ShoCode)
'''''                    .Col = 1
'''''                    .CellFontBold = True
'''''                    .Text = t_nama
'''''                    .Col = 2
''''''''                    .CellAlignment = 4
'''''                    .Text = "Start date"
'''''                    .Col = 3
''''''''                    .CellAlignment = 4
'''''                    .Text = d_begd
''''''                    .Col = 4
''''''                    .CellAlignment = 4
''''''                    .Text = d_endd
'''''                End With
'''''            End If
'''''            ictr = ictr + 1
'''        Loop
'''        .Close
'''    End With
'''    If ictr = 0 Then flex1.Rows = 1 Else flex1.Rows = ictr + 1
    flx1.Visible = True
End Function


Public Sub SizeGrid()
    Dim i As Integer
    
    flx1.Row = 0
    flx1.Col = 0: flx1.ColWidth(0) = (flx1.Width - 240) * 0.42: flx1.ColAlignment(0) = 1
                flx1.CellAlignment = 4: flx1.Text = "Show <Click to view Floorplan>"
    flx1.Col = 1: flx1.ColWidth(1) = (flx1.Width - 240) * 0.18: flx1.ColAlignment(1) = 4: flx1.CellAlignment = 4: flx1.Text = "Show Dates"
    For i = 1 To 8
        flx1.Col = i + 1: flx1.ColWidth(i + 1) = (flx1.Width - 240) * 0.05: flx1.ColAlignment(i + 1) = 4
        Select Case i
            Case 1: flx1.Text = "REQ"
            Case 2: flx1.Text = "FSU"
            Case 3: flx1.Text = "BGD"
            Case 4: flx1.Text = "PRE"
            Case 5: flx1.Text = "AEA"
            Case 6: flx1.Text = "CMP"
            Case 7: flx1.Text = "REL"
            Case 8: flx1.Text = "REV"
        End Select
    Next i
    flx1.Col = 10: flx1.ColWidth(10) = 0 ''' (flx1.Width - 240) * 0.06: flx1.ColAlignment(10) = 4 ''': Set flx1.CellPicture = imgMail(3).Picture
    flx1.Col = 11: flx1.ColWidth(11) = 0
End Sub

Public Sub PopFloorplans(tSHYR As Integer, tCUNO As Long)
    Dim SelClause(0 To 1) As String
    
    If Not bApprovalList Then
        '///// SELCLAUSE(0) IS USED IF CONTROL IS TO DISPLAY POSTED FLOORPLANS ONLY \\\\\
        SelClause(0) = "SELECT SM.SHY56SHCD, SM.SHY56NAMA, CS.CSY56CUNO, CU.ABALPH, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
                    "KU.FPSTATUS, KU.FPSTATDT " & _
                    "FROM " & F5601 & " SM, " & F5611 & " CS, " & _
                    "" & F0101 & " CU, " & IGLKitU & " KU, " & DWGShow & " DS " & _
                    "WHERE CS.CSY56CUNO = " & tCUNO & " " & _
                    "AND SM.SHY56SHYR = " & tSHYR & " " & _
                    "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = KU.AN8_SHCD " & _
                    "AND SM.SHY56SHYR = KU.SHYR " & _
                    "AND CS.CSY56CUNO = KU.AN8_CUNO " & _
                    "AND SM.SHY56SHYR = DS.SHYR " & _
                    "AND SM.SHY56SHCD = DS.AN8_SHCD " & _
                    "AND CS.CSY56CUNO = DS.AN8_CUNO " & _
                    "AND CS.CSY56CUNO = CU.ABAN8 " & _
                    "AND KU.FPSTATUS > 3 "
                
        '///// SELCLAUSE(1) IS USED IF ALL FLOORPLANS ARE TO BE DISPLAYED \\\\\
        SelClause(1) = "SELECT SM.SHY56SHCD, SM.SHY56NAMA, CS.CSY56CUNO, CU.ABALPH, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
                    "KU.FPSTATUS, KU.FPSTATDT " & _
                    "FROM " & F5601 & " SM, " & F5611 & " CS, " & _
                    "" & F0101 & " CU, " & IGLKitU & " KU " & _
                    "WHERE CS.CSY56CUNO = " & tCUNO & " " & _
                    "AND SM.SHY56SHYR = " & tSHYR & " " & _
                    "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = KU.AN8_SHCD " & _
                    "AND SM.SHY56SHYR = KU.SHYR " & _
                    "AND CS.CSY56CUNO = KU.AN8_CUNO " & _
                    "AND CS.CSY56CUNO = CU.ABAN8 " & _
                    "AND KU.FPSTATUS > 0 "
    
    Else
        '///// SELCLAUSE(0) IS USED IF CONTROL IS TO DISPLAY PRELIM FLOORPLANS ONLY \\\\\
        SelClause(0) = "SELECT SM.SHY56SHCD, SM.SHY56NAMA, CS.CSY56CUNO, CU.ABALPH, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
                    "KU.FPSTATUS, KU.FPSTATDT " & _
                    "FROM " & F5601 & " SM, " & F5611 & " CS, " & _
                    "" & F0101 & " CU, " & IGLKitU & " KU, " & DWGShow & " DS " & _
                    "WHERE CS.CSY56CUNO = " & tCUNO & " " & _
                    "AND SM.SHY56SHYR = " & tSHYR & " " & _
                    "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = KU.AN8_SHCD " & _
                    "AND SM.SHY56SHYR = KU.SHYR " & _
                    "AND CS.CSY56CUNO = KU.AN8_CUNO " & _
                    "AND SM.SHY56SHYR = DS.SHYR " & _
                    "AND SM.SHY56SHCD = DS.AN8_SHCD " & _
                    "AND CS.CSY56CUNO = DS.AN8_CUNO " & _
                    "AND CS.CSY56CUNO = CU.ABAN8 " & _
                    "AND KU.FPSTATUS = 4 "
                    
        SelClause(1) = SelClause(0)
    
    End If
    
    SelOrder(0) = "ORDER BY UPPER(SM.SHY56NAMA)"
    SelOrder(1) = "ORDER BY SM.SHY56BEGDT, UPPER(SM.SHY56NAMA)"
        
    If optSort(0).value = True Then
        Select Case bPerm(55)
            Case True
                Call FillGrid(SelClause(1) & SelOrder(0))
            Case False
                Call FillGrid(SelClause(0) & SelOrder(0))
        End Select
    Else
        Select Case bPerm(55)
            Case True
                Call FillGrid(SelClause(1) & SelOrder(1))
            Case False
                Call FillGrid(SelClause(0) & SelOrder(1))
        End Select
    End If
End Sub

Public Sub LoadFloorplan(tSHYR As Integer, tCUNO As Long, tSHCD As Long, tSHNM As String)
    Dim Resp As VbMsgBoxResult
    Dim strSelect As String, sChk As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
'''    Screen.MousePointer = 11
    
'''    '///// CHECK IF RED SHOULD BE SAVED \\\\\
'''    If SaveRed = True Then
'''        Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNoCancel, "Redline Changes...")
'''        If Resp = vbYes Then
'''            mnuRedSave_Click
'''        ElseIf Resp = vbNo Then
'''            volFrame.ClearMarkup
''''''            SaveRed = False
'''        ElseIf Resp = vbCancel Then
'''            GoTo CancelNewFile
'''        End If
'''    End If
    
    '///// FIRST SHUT OFF ALL RELATIVES \\\\\
    For i = 0 To 5
        imgDWF(i).Visible = False
    Next i
    cmdOther.Visible = False
    
    RedFile = "": lRedID = 0
    
    strSelect = "SELECT M.DWGID, DWF.SHTID, DWF.DWFID, " & _
                "DWF.DWFTYPE, DWF.DWFPATH, M.DSTATUS " & _
                "From " & DWGShow & " SHO, " & DWGMas & " M, " & _
                "" & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "WHERE SHO.DWGID = M.DWGID  " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "AND SHO.SHYR = " & tSHYR & " " & _
                "AND SHO.AN8_SHCD = " & tSHCD & " " & _
                "AND SHO.AN8_CUNO = " & tCUNO & " " & _
                "AND DWF.DWFTYPE >= 0 " & _
                "AND DWF.DWFTYPE < 20 " & _
                "AND DWF.DWFSTATUS > 0 " & _
                "ORDER BY DWF.DWFTYPE, DWF.DWFDESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        nLockRefID = rst.Fields("DWGID")
        lDWGID = nLockRefID
        lSHTID = rst.Fields("SHTID")
        Do While Not rst.EOF
            Select Case rst.Fields("DWFTYPE")
                Case 0
                    On Error Resume Next
                    Err.Clear
                    sChk = Dir(Trim(rst.Fields("DWFPATH")), vbNormal)
                    If Err Or sChk = "" Then
                        rst.Close
                        Set rst = Nothing
                        If Err Then
                            MsgBox "File not Found.  File server may be momentarily down." & _
                                        vbNewLine & "Error:  " & Err.Description, _
                                        vbExclamation, "Error Encountered..."
                            Err.Clear
                        Else
                            MsgBox "File not Found", vbExclamation, "Error Encountered..."
                        End If
                        picFrame.Visible = True
                        volFrame.Visible = False
                        volFrame.src = ""
                        GoTo CancelNewFile
                    End If
                    
                    cmdFPApproveHide.Visible = False
                    
                    ''ADD BACK IN LATER''
'                    If rst.Fields("DSTATUS") = 4 Then
'                        bApprover = CheckForApprover(CLng(BCC))
'                    Else
'                        bApprover = False
'                    End If
'                    If bApprover Then
'                        optFPApprove(0).value = False
'                        optFPApprove(1).value = False
'                        optFPApprove(2).value = False
'                        txtFPApprove.Text = ""
'                        picFPApprove.Visible = True
'                    Else
'                        picFPApprove.Visible = False
'                    End If
                
                    imgDWF(rst.Fields("DWFTYPE")).Visible = True
                    RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                Case 1, 2
                    If bPerm(3) Then
                        imgDWF(rst.Fields("DWFTYPE")).Visible = True
                        RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                    End If
                Case 5
                    If bPerm(6) Then
                        imgDWF(rst.Fields("DWFTYPE")).Visible = True
                        RelativePath(rst.Fields("DWFTYPE")) = Trim(rst.Fields("DWFPATH"))
                    End If
                Case 8
                    If bPerm(3) Then cmdOther.Visible = True
                Case 9
                    RedFile = Trim(rst.Fields("DWFPATH"))
                    lRedID = rst.Fields("DWFID")
                    '*** REDLINE NOTE DEALT WITH @ "LOADIT" ***
            End Select
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    strSelect = "SELECT M.DWGID, M.DWGTYPE, DWF.DWFPATH " & _
                "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "Where SHO.SHYR = " & tSHYR & " " & _
                "AND SHO.AN8_SHCD = " & tSHCD & " " & _
                "AND SHO.DWGID = M.DWGID " & _
                "AND M.DWGTYPE IN (3, 4, 5) " & _
                "AND M.DSTATUS > 0 " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND M.DWGID = DWF.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "ORDER BY M.DWGTYPE"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case rst.Fields("DWGTYPE")
            Case 3: If bPerm(4) Then imgDWF(3).Visible = True
            Case 4: If bPerm(5) Then imgDWF(4).Visible = True
            Case 5: If bPerm(6) Then imgDWF(5).Visible = True
        End Select
        RelativePath(rst.Fields("DWGTYPE")) = Trim(rst.Fields("DWFPATH"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    If bPerm(7) Then
        '///// CHECK FOR CLIENT-SHOW RECORD \\\\\
        strSelect = "SELECT CSY56SHCD FROM " & F5611 & " " & _
                    "WHERE CSY56CUNO = " & tCUNO & " " & _
                    "AND CSY56SHYR = " & tSHYR & " " & _
                    "AND CSY56SHCD = " & tSHCD
        Set rst = Conn.Execute(strSelect)
        If rst.EOF Then
            imgInfo.Visible = False
        Else
            imgInfo.Visible = True
        End If
        rst.Close: Set rst = Nothing
    End If
    
    '///// SET NON-VISIBLE TO "" \\\\\
    For i = 0 To 5
        If imgDWF(i).Visible = False Then RelativePath(i) = ""
    Next i
    
    sCurrShow = tSHNM
    ClearChecks
    bLoading = True
    mnuZoomDMode.Checked = True
    sZMode = "Zoom"
    volFrame.UserMode = sZMode
    LoadIt
    lblWelcome.Caption = FBCN & " - " & SHYR & " " & sCurrShow
    lblWelcome.Visible = True
    
    If InitialView Then
        volFrame.Visible = True
        picFrame.Visible = False
        bComm = False
        bLoading = False
        lblReds.Visible = True
        
        If bPerm(2) Then picRelatives.Visible = True
        RelOpen = True
        
    
        If bMenuButton And chkClose.value = 1 Then cmdMenu.Visible = True
        
        '///// LET'S CHECK FOR EXISTING COMMENTS \\\\\
        If bTeam Then
            strSelect = "SELECT COMMID " & _
                        "FROM " & ANOComment & " " & _
                        "WHERE REFID = " & lDWGID & " " & _
                        "AND COMMSTATUS > 0"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then '/// COMMENTS EXIST \\\
                imgComm.Picture = imgMail(1).Picture
                imgComm.ToolTipText = "There is an Active Comment! Click to access."
            Else '/// NO COMMENTS \\\
                imgComm.Picture = imgMail(0).Picture
                imgComm.ToolTipText = "There are no Stored Comments."
            End If
            imgComm.Enabled = True
        Else '/// NO EMAIL TEAM \\\
            imgComm.Picture = imgMail(2).Picture
            imgComm.Enabled = False
        End If
        
        If bPerm(17) Then imgComm.Visible = True
'''        bPicLoaded = True
'''        picDirs.Refresh
    End If
    
    
    
    
    
CancelNewFile:
'''    Screen.MousePointer = 0
End Sub

Public Function CheckForApprover(lCUNO As Long) As Boolean
    Dim iChk As Integer
    If sApprClientList <> "" Then
        iChk = InStr(1, sApprClientList, "|" & CStr(lCUNO) & "|")
        If iChk > 0 Then
            CheckForApprover = True
        Else
            CheckForApprover = False
        End If
    Else
        CheckForApprover = False
    End If
End Function

Public Sub PopApprovalClients(sList As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sCNode As String, sYNode As String, sDesc As String
    
    lstClientSorter.Clear
    tvw1.Nodes.Clear
    tvw2.Visible = False
    tvw2.Nodes.Clear
    tvw2.Visible = True
    sCNode = "": sYNode = ""
'''    strSelect = "SELECT DISTINCT SHO.AN8_CUNO, C.ABALPH, SHO.SHYR " & _
'''                "FROM " & DWGShow & " SHO, " & F0101 & " C, " & DWGMas & " DM " & _
'''                "WHERE SHO.AN8_CUNO IN (" & sList & ") " & _
'''                "AND SHO.DWGID = DM.DWGID " & _
'''                "AND DM.DSTATUS = 4 " & _
'''                "AND SHO.AN8_CUNO = C.ABAN8 " & _
'''                "AND C.ABAT1 = 'C' " & _
'''                "ORDER BY C.ABALPH, SHO.SHYR"
    strSelect = "SELECT DISTINCT KU.AN8_CUNO, C.ABALPH, KU.SHYR  " & _
                "FROM " & IGLKitU & " KU, " & F0101 & " C " & _
                "WHERE KU.AN8_CUNO IN (" & sList & ") " & _
                "AND KU.FPSTATUS = 4 " & _
                "AND KU.AN8_CUNO = C.ABAN8 " & _
                "AND C.ABAT1 = 'C' " & _
                "ORDER BY C.ABALPH, KU.SHYR"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sCNode <> "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8) Then
            sCNode = "c" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
            sDesc = UCase(Trim(rst.Fields("ABALPH")))
            Set nodX = tvw1.Nodes.Add(, , sCNode, sDesc)
            Set nodX = tvw2.Nodes.Add(, , sCNode, sDesc)
            lstClientSorter.AddItem UCase(Trim(rst.Fields("ABALPH")))
            lstClientSorter.ItemData(lstClientSorter.NewIndex) = rst.Fields("AN8_CUNO")
        End If
        
        sYNode = "y" & rst.Fields("SHYR") & "-" & Right("00000000" & rst.Fields("AN8_CUNO"), 8)
        sDesc = rst.Fields("SHYR")
        Set nodX = tvw1.Nodes.Add(sCNode, tvwChild, sYNode, sDesc)
        Set nodX = tvw2.Nodes.Add(sCNode, tvwChild, sYNode, sDesc)
        rst.MoveNext
    Loop

End Sub

Public Sub CheckIfReadyToApprove(Index As Integer)
    Dim sChk As String
    Dim sDate As Date
    Dim Resp As VbMsgBoxResult
    
    Select Case Index
        Case 0: cmdFPApprove.Enabled = True
        Case 1
            ''CHECK FOR REDFILE BEFORE ENABLING''
            ''CHECK DATE FOR CURRENT''
            On Error Resume Next
            If RedFile <> "" Then
                sChk = Dir(RedFile, vbNormal)
                If sChk = "" Or Err > 0 Then
                    MsgBox "No Redline File exists for this Floorplan", _
                                vbExclamation, "Sorry..."
                    cmdFPApprove.Enabled = False
                    Exit Sub
                Else
                    sDate = FileDateTime(RedFile)
                    If sDate < DateAdd("d", -1, Now) Then
                        Resp = MsgBox("The existing Redline file is dated:" & vbCr & vbCr & _
                                    vbTab & sDate & vbCr & vbCr & _
                                    "Is this the file you want to Approve with?", _
                                    vbQuestion + vbYesNoCancel, "Confirming Redline File...")
                        If Resp = vbYes Then
                            cmdFPApprove.Enabled = True
                        Else
                            cmdFPApprove.Enabled = False
                        End If
                    Else
                        cmdFPApprove.Enabled = True
                    End If
                End If
            Else
                MsgBox "No Redline File exists for this Floorplan", _
                            vbExclamation, "Sorry..."
                cmdFPApprove.Enabled = False
            End If
                
        Case 2
            If txtFPApprove.Text <> "" Then
                cmdFPApprove.Enabled = True
            Else
                cmdFPApprove.Enabled = False
            End If
    End Select
End Sub

Public Sub CheckForImages(tLink As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim tBCC As Long
    Dim tElt As String
    
    tBCC = CLng(Left(tLink, 8))
    tElt = Mid(tLink, 10)
    
    strSelect = "SELECT GM.GID " & _
                "FROM GFX_ELEMENT GE, GFX_MASTER GM " & _
                "WHERE GE.ELTID IN " & _
                "(SELECT E.ELTID " & _
                "FROM IGL_ELEMENT E, IGL_KIT K " & _
                "WHERE K.AN8_CUNO = " & tBCC & " " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTFNAME = '" & tElt & "') " & _
                "AND GE.GID = GM.GID " & _
                "AND GM.GTYPE = 1"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then mnuPhoto.Enabled = False Else mnuPhoto.Enabled = True
    rst.Close
    
    strSelect = "SELECT GM.GID " & _
                "FROM GFX_ELEMENT GE, GFX_MASTER GM " & _
                "WHERE GE.ELTID IN " & _
                "(SELECT E.ELTID " & _
                "FROM IGL_ELEMENT E, IGL_KIT K " & _
                "WHERE K.AN8_CUNO = " & tBCC & " " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTFNAME = '" & tElt & "') " & _
                "AND GE.GID = GM.GID " & _
                "AND GM.GTYPE IN (2, 3)"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then mnuGFX.Enabled = False Else mnuGFX.Enabled = True
    rst.Close
    
    Set rst = Nothing
    
End Sub
