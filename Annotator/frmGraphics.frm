VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{8C445A83-9D0A-11D3-A8FB-444553540000}#1.0#0"; "ImagXpr5.dll"
Object = "{23319180-2253-11D7-BD2E-08004608C318}#3.0#0"; "XpdfViewerCtrl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGraphics 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "GPJ Graphics Handler & Annotator"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraphics.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "FindNext"
      Height          =   315
      Left            =   960
      TabIndex        =   219
      Top             =   7800
      Width           =   975
   End
   Begin MSComctlLib.ImageList imlRedMode 
      Left            =   2340
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":12C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":26B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":30B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":3AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":44A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":4AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":511A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":56B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":5C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":6508
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":6DC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":767C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":7F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":87F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstZoom 
      Height          =   645
      ItemData        =   "frmGraphics.frx":90AA
      Left            =   1200
      List            =   "frmGraphics.frx":90F6
      TabIndex        =   210
      Top             =   6600
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSComctlLib.ImageList imlNav 
      Left            =   7740
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":9190
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":97CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":9E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":A43E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Index           =   4
      ItemData        =   "frmGraphics.frx":AA78
      Left            =   120
      List            =   "frmGraphics.frx":AA7A
      TabIndex        =   209
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Index           =   3
      ItemData        =   "frmGraphics.frx":AA7C
      Left            =   120
      List            =   "frmGraphics.frx":AA7E
      TabIndex        =   208
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Index           =   2
      ItemData        =   "frmGraphics.frx":AA80
      Left            =   120
      List            =   "frmGraphics.frx":AA82
      TabIndex        =   207
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Index           =   1
      ItemData        =   "frmGraphics.frx":AA84
      Left            =   120
      List            =   "frmGraphics.frx":AA86
      TabIndex        =   206
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Index           =   0
      ItemData        =   "frmGraphics.frx":AA88
      Left            =   120
      List            =   "frmGraphics.frx":AA8A
      TabIndex        =   205
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList imlJPGTools 
      Left            =   1620
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":AA8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":B1B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":B8E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":C00A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":C734
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":CE5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D588
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DCB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E3DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPageMode 
      Left            =   1020
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":EB06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":F230
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":10084
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":107AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":10D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":112E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":1187C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSkins 
      Left            =   10740
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1600
      ImageHeight     =   75
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":11E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":69CAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlDirs 
      Left            =   4980
      Top             =   0
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
            Picture         =   "frmGraphics.frx":C1B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":C2698
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picWait 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   6600
      ScaleHeight     =   555
      ScaleWidth      =   2415
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label4 
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
         TabIndex        =   13
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.PictureBox picMess 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   5460
      ScaleHeight     =   915
      ScaleWidth      =   4155
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   4155
      Begin VB.Label lblMess 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1740
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox picTabs 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      Height          =   7155
      Left            =   1980
      ScaleHeight     =   7155
      ScaleWidth      =   11550
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   11550
      Begin VB.PictureBox picType 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   9540
         ScaleHeight     =   435
         ScaleWidth      =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   2940
         Begin VB.Label lblDisplayAll 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Display All"
            Height          =   195
            Left            =   120
            MouseIcon       =   "frmGraphics.frx":C31F2
            MousePointer    =   99  'Custom
            TabIndex        =   11
            ToolTipText     =   "Display All Graphic Types"
            Top             =   120
            Width           =   735
         End
         Begin VB.Image imgType 
            Height          =   315
            Index           =   1
            Left            =   1020
            MouseIcon       =   "frmGraphics.frx":C34FC
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":C3806
            Stretch         =   -1  'True
            ToolTipText     =   "Display Digital Photos Only"
            Top             =   60
            Width           =   375
         End
         Begin VB.Image imgType 
            Height          =   315
            Index           =   2
            Left            =   1500
            MouseIcon       =   "frmGraphics.frx":C3904
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":C3C0E
            Stretch         =   -1  'True
            ToolTipText     =   "Display Graphic Files Only"
            Top             =   60
            Width           =   380
         End
         Begin VB.Image imgType 
            Height          =   375
            Index           =   3
            Left            =   1970
            MouseIcon       =   "frmGraphics.frx":C3DBD
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":C40C7
            Stretch         =   -1  'True
            ToolTipText     =   "Display Graphic Layouts Only"
            Top             =   0
            Width           =   380
         End
         Begin VB.Image imgType 
            Height          =   330
            Index           =   4
            Left            =   2480
            MouseIcon       =   "frmGraphics.frx":C43D1
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":C46DB
            Stretch         =   -1  'True
            ToolTipText     =   "Display Presentation Files Only"
            Top             =   60
            Width           =   320
         End
         Begin VB.Shape shpType 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000002&
            BorderWidth     =   3
            Height          =   435
            Left            =   2400
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7200
         TabIndex        =   167
         Top             =   660
         Visible         =   0   'False
         Width           =   1995
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            Caption         =   "Email Copy..."
            Height          =   195
            Left            =   960
            MouseIcon       =   "frmGraphics.frx":C4CB2
            MousePointer    =   99  'Custom
            TabIndex        =   169
            Top             =   0
            Width           =   960
         End
         Begin VB.Label lblDownload 
            AutoSize        =   -1  'True
            Caption         =   "Download..."
            Height          =   195
            Left            =   0
            MouseIcon       =   "frmGraphics.frx":C4FBC
            MousePointer    =   99  'Custom
            TabIndex        =   168
            Top             =   0
            Width           =   885
         End
      End
      Begin TabDlg.SSTab sst1 
         Height          =   7155
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   12621
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   882
         ShowFocusRect   =   0   'False
         BackColor       =   0
         TabCaption(0)   =   "Active Show Graphics"
         TabPicture(0)   =   "frmGraphics.frx":C52C6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblInactive(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblLast(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblList(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblPipe(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblPipe(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblPipe(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblFirst(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblNext(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblPrevious(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblCnt(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "fra0"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "tvwGraphics(0)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "picOuter(1)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "picOuter(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "chkClose(0)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "picOuter(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "hsc1(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "chkApproved(0)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Show Season Graphics"
         TabPicture(1)   =   "frmGraphics.frx":C52E2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkApproved(1)"
         Tab(1).Control(1)=   "hsc1(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "optSort(0)"
         Tab(1).Control(3)=   "optSort(1)"
         Tab(1).Control(4)=   "cboCUNO(1)"
         Tab(1).Control(5)=   "cboSHYR(1)"
         Tab(1).Control(6)=   "chkClose(1)"
         Tab(1).Control(7)=   "tvwGraphics(1)"
         Tab(1).Control(8)=   "lblFirst(1)"
         Tab(1).Control(9)=   "lblCnt(1)"
         Tab(1).Control(10)=   "lblPrevious(1)"
         Tab(1).Control(11)=   "lblNext(1)"
         Tab(1).Control(12)=   "lblPipe(5)"
         Tab(1).Control(13)=   "lblPipe(4)"
         Tab(1).Control(14)=   "lblPipe(3)"
         Tab(1).Control(15)=   "lblList(1)"
         Tab(1).Control(16)=   "lblLast(1)"
         Tab(1).Control(17)=   "lblInactive(1)"
         Tab(1).Control(18)=   "Label5"
         Tab(1).Control(19)=   "Label1"
         Tab(1).ControlCount=   20
         TabCaption(2)   =   "Kit-Based Graphics"
         TabPicture(2)   =   "frmGraphics.frx":C52FE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkApproved(2)"
         Tab(2).Control(1)=   "cboCUNO(2)"
         Tab(2).Control(2)=   "chkClose(2)"
         Tab(2).Control(3)=   "hsc1(2)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "tvwGraphics(2)"
         Tab(2).Control(5)=   "lblCnt(2)"
         Tab(2).Control(6)=   "lblPrevious(2)"
         Tab(2).Control(7)=   "lblNext(2)"
         Tab(2).Control(8)=   "lblFirst(2)"
         Tab(2).Control(9)=   "lblPipe(8)"
         Tab(2).Control(10)=   "lblPipe(7)"
         Tab(2).Control(11)=   "lblPipe(6)"
         Tab(2).Control(12)=   "lblList(2)"
         Tab(2).Control(13)=   "lblLast(2)"
         Tab(2).Control(14)=   "lblInactive(2)"
         Tab(2).Control(15)=   "Label6(0)"
         Tab(2).ControlCount=   16
         TabCaption(3)   =   "Client Graphics"
         TabPicture(3)   =   "frmGraphics.frx":C531A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label7"
         Tab(3).Control(1)=   "lblInactive(3)"
         Tab(3).Control(2)=   "lblLast(3)"
         Tab(3).Control(3)=   "lblList(3)"
         Tab(3).Control(4)=   "lblPipe(9)"
         Tab(3).Control(5)=   "lblPipe(10)"
         Tab(3).Control(6)=   "lblPipe(11)"
         Tab(3).Control(7)=   "lblFirst(3)"
         Tab(3).Control(8)=   "lblNext(3)"
         Tab(3).Control(9)=   "lblPrevious(3)"
         Tab(3).Control(10)=   "lblCnt(3)"
         Tab(3).Control(11)=   "tvwGraphics(3)"
         Tab(3).Control(12)=   "chkClose(3)"
         Tab(3).Control(13)=   "cboCUNO(3)"
         Tab(3).Control(14)=   "picOuter(3)"
         Tab(3).Control(15)=   "hsc1(3)"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).Control(16)=   "chkApproved(3)"
         Tab(3).ControlCount=   17
         TabCaption(4)   =   "Approval Interface"
         TabPicture(4)   =   "frmGraphics.frx":C5336
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraRefresh"
         Tab(4).Control(1)=   "picReview"
         Tab(4).Control(2)=   "cboFolder"
         Tab(4).Control(3)=   "txtNoShows"
         Tab(4).Control(4)=   "chkClose(4)"
         Tab(4).Control(5)=   "picHelp"
         Tab(4).Control(6)=   "picOuter(4)"
         Tab(4).Control(7)=   "cboCUNO(4)"
         Tab(4).Control(8)=   "cboSHYR(4)"
         Tab(4).Control(9)=   "cboASHCD"
         Tab(4).Control(10)=   "flxApprove"
         Tab(4).Control(11)=   "Label9"
         Tab(4).Control(12)=   "lblFileCount"
         Tab(4).Control(13)=   "lblClient"
         Tab(4).Control(14)=   "Label10"
         Tab(4).ControlCount=   15
         Begin VB.Frame fraRefresh 
            Caption         =   "Refresh"
            Height          =   1095
            Left            =   -70260
            TabIndex        =   188
            Top             =   600
            Width           =   2055
            Begin VB.CommandButton cmdRefresh 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Get Files"
               Enabled         =   0   'False
               Height          =   795
               Left            =   1020
               Picture         =   "frmGraphics.frx":C5352
               Style           =   1  'Graphical
               TabIndex        =   193
               ToolTipText     =   "Click to refresh list based on changed settings"
               Top             =   180
               Width           =   915
            End
            Begin VB.PictureBox Picture4 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   675
               Left            =   60
               ScaleHeight     =   675
               ScaleWidth      =   1935
               TabIndex        =   189
               Top             =   360
               Width           =   1935
               Begin VB.OptionButton optApproverView 
                  Caption         =   "Search Result"
                  Enabled         =   0   'False
                  Height          =   195
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   192
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   1455
               End
               Begin VB.OptionButton optApproverView 
                  Caption         =   "My Files"
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  TabIndex        =   191
                  Top             =   60
                  Width           =   915
               End
               Begin VB.OptionButton optApproverView 
                  Caption         =   "All Files"
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   190
                  Top             =   360
                  Width           =   915
               End
            End
         End
         Begin VB.PictureBox picReview 
            AutoRedraw      =   -1  'True
            Height          =   495
            Left            =   -67920
            ScaleHeight     =   435
            ScaleWidth      =   3795
            TabIndex        =   18
            Top             =   1320
            Visible         =   0   'False
            Width           =   3855
            Begin VB.OptionButton optFilter 
               Height          =   435
               Index           =   2
               Left            =   4080
               MaskColor       =   &H8000000F&
               Picture         =   "frmGraphics.frx":C5CDC
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Click to View Relased Files Only"
               Top             =   840
               UseMaskColor    =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton optFilter 
               Height          =   435
               Index           =   1
               Left            =   2040
               MaskColor       =   &H8000000F&
               Picture         =   "frmGraphics.frx":C6366
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Click to View Draft Files Only"
               Top             =   840
               UseMaskColor    =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton optFilter 
               Caption         =   "Display All"
               Height          =   435
               Index           =   0
               Left            =   300
               MaskColor       =   &H8000000F&
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Click to View Both Draft and Released Files"
               Top             =   780
               UseMaskColor    =   -1  'True
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.CommandButton cmdStatusEdit_View 
               Caption         =   "Edit All..."
               Enabled         =   0   'False
               Height          =   435
               Left            =   3000
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Edit All Files in Current View..."
               Top             =   0
               Width           =   780
            End
            Begin VB.CommandButton cmdStatusEdit_Notify 
               Caption         =   "Notify..."
               Height          =   435
               Left            =   5460
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   780
               Width           =   840
            End
            Begin VB.CommandButton cmdHelp 
               Height          =   435
               Left            =   6600
               Picture         =   "frmGraphics.frx":C69F0
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   660
               Width           =   460
            End
            Begin VB.Image imgStatus 
               Height          =   585
               Index           =   1
               Left            =   900
               MouseIcon       =   "frmGraphics.frx":C6B3A
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":C6E44
               ToolTipText     =   "Display Internal Draft Files Only"
               Top             =   120
               Width           =   585
            End
            Begin VB.Label lblFilterAll 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Display All"
               Height          =   195
               Left            =   60
               MouseIcon       =   "frmGraphics.frx":C7312
               MousePointer    =   99  'Custom
               TabIndex        =   25
               ToolTipText     =   "Display both Internal & Client Draft Files"
               Top             =   120
               Width           =   735
            End
            Begin VB.Image imgStatus 
               Height          =   585
               Index           =   2
               Left            =   1620
               MouseIcon       =   "frmGraphics.frx":C761C
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":C7926
               ToolTipText     =   "Display Client Draft Files Only"
               Top             =   120
               Width           =   585
            End
            Begin VB.Image imgStatus 
               Height          =   585
               Index           =   3
               Left            =   2340
               MouseIcon       =   "frmGraphics.frx":C7DF4
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":C80FE
               ToolTipText     =   "Display Files Returned for Changes Only"
               Top             =   120
               Width           =   585
            End
            Begin VB.Image imgStatus 
               Height          =   585
               Index           =   4
               Left            =   3180
               MouseIcon       =   "frmGraphics.frx":C8C98
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":C8FA2
               ToolTipText     =   "Display Approved Files Only"
               Top             =   120
               Width           =   585
            End
            Begin VB.Shape shpStatus 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000002&
               BorderWidth     =   3
               Height          =   435
               Left            =   0
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.ComboBox cboFolder 
            Height          =   315
            Left            =   -74340
            Style           =   2  'Dropdown List
            TabIndex        =   186
            Top             =   1500
            Width           =   3975
         End
         Begin VB.CheckBox chkApproved 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Approved Only"
            Height          =   195
            Index           =   1
            Left            =   -72480
            MaskColor       =   &H8000000F&
            TabIndex        =   184
            Top             =   6825
            Width           =   1875
         End
         Begin VB.CheckBox chkApproved 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Approved Only"
            Height          =   195
            Index           =   3
            Left            =   -72480
            MaskColor       =   &H8000000F&
            TabIndex        =   183
            Top             =   6825
            Width           =   1875
         End
         Begin VB.CheckBox chkApproved 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Approved Only"
            Height          =   195
            Index           =   2
            Left            =   -72480
            MaskColor       =   &H8000000F&
            TabIndex        =   182
            Top             =   6825
            Width           =   1875
         End
         Begin VB.CheckBox chkApproved 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Approved Only"
            Height          =   195
            Index           =   0
            Left            =   2520
            MaskColor       =   &H8000000F&
            TabIndex        =   181
            Top             =   6825
            Width           =   1875
         End
         Begin VB.TextBox txtNoShows 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "No Graphics awaiting review have been assigned to a Show"
            Top             =   1125
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.HScrollBar hsc1 
            Height          =   195
            Index           =   3
            Left            =   -70500
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   6495
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.HScrollBar hsc1 
            Height          =   195
            Index           =   1
            Left            =   -70500
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   6495
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.HScrollBar hsc1 
            Height          =   195
            Index           =   0
            Left            =   4500
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   6495
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   195
            Index           =   4
            Left            =   -74820
            MaskColor       =   &H8000000F&
            TabIndex        =   76
            Top             =   6810
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmGraphics.frx":C9470
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
            Left            =   -70380
            Picture         =   "frmGraphics.frx":C977A
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Sort Show List Alphabetically"
            Top             =   615
            Width           =   675
         End
         Begin VB.OptionButton optSort 
            DownPicture     =   "frmGraphics.frx":CA044
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
            Left            =   -69720
            Picture         =   "frmGraphics.frx":CA34E
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Sort Show List Chronologically (by Show Open Date)"
            Top             =   615
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.PictureBox picHelp 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   1935
            Left            =   -67200
            ScaleHeight     =   1875
            ScaleWidth      =   2955
            TabIndex        =   59
            Top             =   4215
            Visible         =   0   'False
            Width           =   3015
            Begin VB.CommandButton cmdHelpClose 
               Caption         =   "Close"
               Height          =   435
               Left            =   1500
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   0
               Width           =   1275
            End
            Begin SHDocVwCtl.WebBrowser web1 
               Height          =   1455
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   2895
               ExtentX         =   5106
               ExtentY         =   2566
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
         Begin VB.PictureBox picOuter 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5295
            Index           =   4
            Left            =   -74820
            ScaleHeight     =   5295
            ScaleWidth      =   3030
            TabIndex        =   55
            Top             =   2220
            Width           =   3030
            Begin VB.PictureBox picInner 
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Height          =   5895
               Index           =   4
               Left            =   0
               ScaleHeight     =   5895
               ScaleWidth      =   3330
               TabIndex        =   56
               Top             =   600
               Width           =   3330
               Begin IMAGXPR5LibCtl.ImagXpress imx4 
                  Height          =   960
                  Index           =   0
                  Left            =   60
                  TabIndex        =   57
                  ToolTipText     =   "Click to Open - Rght-Click to reset Status"
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1693
                  ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
                  ErrCode         =   1288381336
                  ErrInfo         =   -275179512
                  Persistence     =   -1  'True
                  _cx             =   132055552
                  _cy             =   1
                  FileName        =   ""
                  MouseIcon       =   "frmGraphics.frx":CAC18
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
               Begin VB.CheckBox chk4 
                  BackColor       =   &H80000005&
                  Height          =   195
                  Index           =   0
                  Left            =   60
                  MaskColor       =   &H80000005&
                  MouseIcon       =   "frmGraphics.frx":CAF32
                  MousePointer    =   99  'Custom
                  TabIndex        =   174
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin VB.Image imgV 
                  Height          =   240
                  Index           =   0
                  Left            =   1320
                  MouseIcon       =   "frmGraphics.frx":CB23C
                  MousePointer    =   99  'Custom
                  Picture         =   "frmGraphics.frx":CB546
                  ToolTipText     =   "Click to view Multiple Versions..."
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   240
               End
               Begin VB.Label lblStat 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Height          =   195
                  Index           =   0
                  Left            =   1980
                  MouseIcon       =   "frmGraphics.frx":CBAD0
                  MousePointer    =   99  'Custom
                  TabIndex        =   58
                  ToolTipText     =   "Click to reset Status"
                  Top             =   435
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   45
               End
               Begin VB.Image imgStat 
                  Height          =   360
                  Index           =   0
                  Left            =   1560
                  MouseIcon       =   "frmGraphics.frx":CBDDA
                  MousePointer    =   99  'Custom
                  Stretch         =   -1  'True
                  ToolTipText     =   "Click to reset Status"
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   360
               End
               Begin VB.Shape shp4 
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H80000005&
                  Height          =   1080
                  Index           =   0
                  Left            =   0
                  Top             =   0
                  Width           =   3030
               End
            End
         End
         Begin VB.ComboBox cboCUNO 
            Height          =   315
            Index           =   4
            Left            =   -74340
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   735
            Width           =   3975
         End
         Begin VB.PictureBox picOuter 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   5535
            Index           =   0
            Left            =   4500
            ScaleHeight     =   5475
            ScaleWidth      =   6795
            TabIndex        =   50
            Top             =   1155
            Width           =   6855
            Begin VB.PictureBox picInner 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   5475
               Index           =   0
               Left            =   0
               ScaleHeight     =   5475
               ScaleWidth      =   6375
               TabIndex        =   51
               Top             =   0
               Width           =   6375
               Begin VB.CheckBox chk0 
                  BackColor       =   &H80000005&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  MaskColor       =   &H80000005&
                  MouseIcon       =   "frmGraphics.frx":CC0E4
                  MousePointer    =   99  'Custom
                  TabIndex        =   170
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin IMAGXPR5LibCtl.ImagXpress imx0 
                  Height          =   960
                  Index           =   0
                  Left            =   120
                  TabIndex        =   52
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1693
                  ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
                  ErrCode         =   1288381336
                  ErrInfo         =   -275179512
                  Persistence     =   -1  'True
                  _cx             =   132055184
                  _cy             =   1
                  FileName        =   ""
                  MouseIcon       =   "frmGraphics.frx":CC3EE
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
                     Name            =   "Arial Narrow"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   720
                  TabIndex        =   53
                  Top             =   1080
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   45
               End
            End
         End
         Begin VB.PictureBox picOuter 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   5535
            Index           =   3
            Left            =   -70500
            ScaleHeight     =   5475
            ScaleWidth      =   6795
            TabIndex        =   46
            Top             =   1155
            Width           =   6855
            Begin VB.PictureBox picInner 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   5475
               Index           =   3
               Left            =   0
               ScaleHeight     =   5475
               ScaleWidth      =   6375
               TabIndex        =   47
               Top             =   0
               Width           =   6375
               Begin VB.CheckBox chk3 
                  BackColor       =   &H80000005&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  MaskColor       =   &H80000005&
                  MouseIcon       =   "frmGraphics.frx":CC708
                  MousePointer    =   99  'Custom
                  TabIndex        =   171
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin IMAGXPR5LibCtl.ImagXpress imx3 
                  Height          =   960
                  Index           =   0
                  Left            =   120
                  TabIndex        =   48
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1693
                  ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
                  ErrCode         =   1288381336
                  ErrInfo         =   -275179512
                  Persistence     =   -1  'True
                  _cx             =   132063040
                  _cy             =   1
                  FileName        =   ""
                  MouseIcon       =   "frmGraphics.frx":CCA12
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
               Begin VB.Label lbl3 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   0
                  Left            =   720
                  TabIndex        =   49
                  Top             =   1080
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   75
               End
               Begin VB.Shape shp3 
                  BackColor       =   &H000000FF&
                  BackStyle       =   1  'Opaque
                  BorderStyle     =   0  'Transparent
                  Height          =   195
                  Index           =   0
                  Left            =   660
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   255
               End
            End
         End
         Begin VB.ComboBox cboCUNO 
            Height          =   315
            Index           =   3
            Left            =   -74820
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   765
            Width           =   4335
         End
         Begin VB.ComboBox cboCUNO 
            Height          =   315
            Index           =   2
            Left            =   -74820
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   765
            Width           =   4335
         End
         Begin VB.ComboBox cboCUNO 
            Height          =   315
            Index           =   1
            Left            =   -73860
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   765
            Width           =   3375
         End
         Begin VB.ComboBox cboSHYR 
            Height          =   315
            Index           =   1
            Left            =   -74820
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   765
            Width           =   855
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   195
            Index           =   3
            Left            =   -74820
            MaskColor       =   &H8000000F&
            TabIndex        =   41
            Top             =   6810
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   315
            Index           =   0
            Left            =   180
            MaskColor       =   &H8000000F&
            TabIndex        =   40
            Top             =   6765
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   195
            Index           =   2
            Left            =   -74820
            MaskColor       =   &H8000000F&
            TabIndex        =   39
            Top             =   6810
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox chkClose 
            Caption         =   "Auto-Close with Selection"
            Height          =   195
            Index           =   1
            Left            =   -74820
            MaskColor       =   &H8000000F&
            TabIndex        =   38
            Top             =   6810
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.PictureBox picOuter 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   5535
            Index           =   2
            Left            =   4500
            ScaleHeight     =   5475
            ScaleWidth      =   6795
            TabIndex        =   34
            Top             =   1155
            Width           =   6855
            Begin VB.PictureBox picInner 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   5475
               Index           =   2
               Left            =   0
               ScaleHeight     =   5475
               ScaleWidth      =   6375
               TabIndex        =   35
               Top             =   0
               Width           =   6375
               Begin VB.CheckBox chk2 
                  BackColor       =   &H80000005&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  MaskColor       =   &H80000005&
                  MouseIcon       =   "frmGraphics.frx":CCD2C
                  MousePointer    =   99  'Custom
                  TabIndex        =   173
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin IMAGXPR5LibCtl.ImagXpress imx2 
                  Height          =   960
                  Index           =   0
                  Left            =   120
                  TabIndex        =   36
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1693
                  ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
                  ErrCode         =   1288381336
                  ErrInfo         =   -275179512
                  Persistence     =   -1  'True
                  _cx             =   132062768
                  _cy             =   1
                  FileName        =   ""
                  MouseIcon       =   "frmGraphics.frx":CD036
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
               Begin VB.Label lbl2 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   225
                  Index           =   0
                  Left            =   720
                  TabIndex        =   37
                  Top             =   1080
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   75
               End
               Begin VB.Shape shp2 
                  BackColor       =   &H000000FF&
                  BackStyle       =   1  'Opaque
                  BorderStyle     =   0  'Transparent
                  Height          =   195
                  Index           =   0
                  Left            =   600
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   255
               End
            End
         End
         Begin VB.PictureBox picOuter 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            Height          =   5535
            Index           =   1
            Left            =   4500
            ScaleHeight     =   5475
            ScaleWidth      =   6795
            TabIndex        =   30
            Top             =   1155
            Width           =   6855
            Begin VB.PictureBox picInner 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   5475
               Index           =   1
               Left            =   0
               ScaleHeight     =   5475
               ScaleWidth      =   6375
               TabIndex        =   31
               Top             =   0
               Width           =   6375
               Begin VB.CheckBox chk1 
                  BackColor       =   &H80000005&
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  MaskColor       =   &H80000005&
                  MouseIcon       =   "frmGraphics.frx":CD350
                  MousePointer    =   99  'Custom
                  TabIndex        =   172
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   195
               End
               Begin IMAGXPR5LibCtl.ImagXpress imx1 
                  Height          =   960
                  Index           =   0
                  Left            =   120
                  TabIndex        =   32
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   1693
                  ErrStr          =   "QWZ600P0GEP-YB305TSXEP"
                  ErrCode         =   1288381336
                  ErrInfo         =   -275179512
                  Persistence     =   -1  'True
                  _cx             =   132062496
                  _cy             =   1
                  FileName        =   ""
                  MouseIcon       =   "frmGraphics.frx":CD65A
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
               Begin VB.Label lbl1 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   720
                  TabIndex        =   33
                  Top             =   1080
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   105
               End
               Begin VB.Shape shp1 
                  BackColor       =   &H000000FF&
                  BackStyle       =   1  'Opaque
                  BorderStyle     =   0  'Transparent
                  Height          =   195
                  Index           =   0
                  Left            =   660
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   255
               End
            End
         End
         Begin VB.HScrollBar hsc1 
            Height          =   195
            Index           =   2
            Left            =   -70500
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   6495
            Visible         =   0   'False
            Width           =   6855
         End
         Begin VB.ComboBox cboSHYR 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   -74340
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1140
            Width           =   855
         End
         Begin VB.ComboBox cboASHCD 
            Enabled         =   0   'False
            Height          =   315
            Left            =   -73440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1125
            Width           =   3075
         End
         Begin MSComctlLib.TreeView tvwGraphics 
            Height          =   5535
            Index           =   0
            Left            =   180
            TabIndex        =   64
            Top             =   1155
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
         End
         Begin MSComctlLib.TreeView tvwGraphics 
            Height          =   5535
            Index           =   1
            Left            =   -74820
            TabIndex        =   65
            Top             =   1155
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
         End
         Begin MSComctlLib.TreeView tvwGraphics 
            Height          =   5535
            Index           =   3
            Left            =   -74820
            TabIndex        =   66
            Top             =   1155
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
         End
         Begin MSComctlLib.TreeView tvwGraphics 
            Height          =   5535
            Index           =   2
            Left            =   -74820
            TabIndex        =   67
            Top             =   1155
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   9763
            _Version        =   393217
            Indentation     =   176
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
         End
         Begin MSFlexGridLib.MSFlexGrid flxApprove 
            Height          =   5235
            Left            =   -74820
            TabIndex        =   68
            Top             =   1980
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   9234
            _Version        =   393216
            Rows            =   10
            Cols            =   10
            FixedCols       =   0
            BackColorBkg    =   -2147483643
            GridColorFixed  =   -2147483643
            WordWrap        =   -1  'True
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLines       =   0
            ScrollBars      =   2
            BorderStyle     =   0
            Appearance      =   0
            MouseIcon       =   "frmGraphics.frx":CD974
         End
         Begin VB.Frame fra0 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   60
            TabIndex        =   69
            Top             =   555
            Width           =   11415
            Begin VB.ComboBox cboSHCD 
               Height          =   315
               Left            =   4560
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   210
               Width           =   4335
            End
            Begin VB.ComboBox cboSHYR 
               Height          =   315
               Index           =   0
               Left            =   120
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   210
               Width           =   855
            End
            Begin VB.ComboBox cboCUNO 
               Height          =   315
               Index           =   0
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   180
               Width           =   3375
            End
            Begin VB.Label lblShow 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show:"
               Height          =   195
               Left            =   4560
               TabIndex        =   75
               Top             =   0
               Width           =   450
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Show Year:"
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   0
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Client:"
               Height          =   195
               Left            =   1080
               TabIndex        =   73
               Top             =   0
               Width           =   465
            End
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Folder:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   187
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   510
         End
         Begin VB.Label lblFirst 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   -70440
            MouseIcon       =   "frmGraphics.frx":CDC8E
            MousePointer    =   99  'Custom
            TabIndex        =   185
            Top             =   6825
            Width           =   315
         End
         Begin VB.Label lblCnt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   3
            Left            =   -63720
            TabIndex        =   166
            Top             =   6825
            Width           =   45
         End
         Begin VB.Label lblCnt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   2
            Left            =   -63720
            TabIndex        =   165
            Top             =   6825
            Width           =   45
         End
         Begin VB.Label lblCnt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   1
            Left            =   -63720
            TabIndex        =   164
            Top             =   6825
            Width           =   45
         End
         Begin VB.Label lblCnt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   0
            Left            =   11280
            TabIndex        =   163
            Top             =   6825
            Width           =   45
         End
         Begin VB.Label lblPrevious 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   -69840
            MouseIcon       =   "frmGraphics.frx":CDF98
            MousePointer    =   99  'Custom
            TabIndex        =   162
            Top             =   6825
            Width           =   615
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   -68940
            MouseIcon       =   "frmGraphics.frx":CE2A2
            MousePointer    =   99  'Custom
            TabIndex        =   161
            Top             =   6825
            Width           =   345
         End
         Begin VB.Label lblFirst 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   -70440
            MouseIcon       =   "frmGraphics.frx":CE5AC
            MousePointer    =   99  'Custom
            TabIndex        =   160
            Top             =   6825
            Width           =   315
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   11
            Left            =   -70020
            MouseIcon       =   "frmGraphics.frx":CE8B6
            TabIndex        =   159
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   10
            Left            =   -69120
            MouseIcon       =   "frmGraphics.frx":CEBC0
            TabIndex        =   158
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   9
            Left            =   -68520
            MouseIcon       =   "frmGraphics.frx":CEECA
            TabIndex        =   157
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text List..."
            Height          =   195
            Index           =   3
            Left            =   -67560
            MouseIcon       =   "frmGraphics.frx":CF1D4
            MousePointer    =   99  'Custom
            TabIndex        =   156
            Top             =   6825
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLast 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   -68340
            MouseIcon       =   "frmGraphics.frx":CF4DE
            MousePointer    =   99  'Custom
            TabIndex        =   155
            Top             =   6825
            Width           =   300
         End
         Begin VB.Label lblPrevious 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   -69840
            MouseIcon       =   "frmGraphics.frx":CF7E8
            MousePointer    =   99  'Custom
            TabIndex        =   154
            Top             =   6825
            Width           =   615
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   -68940
            MouseIcon       =   "frmGraphics.frx":CFAF2
            MousePointer    =   99  'Custom
            TabIndex        =   153
            Top             =   6825
            Width           =   345
         End
         Begin VB.Label lblFirst 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   -70440
            MouseIcon       =   "frmGraphics.frx":CFDFC
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   6825
            Width           =   315
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   8
            Left            =   -70020
            MouseIcon       =   "frmGraphics.frx":D0106
            TabIndex        =   151
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   -69120
            MouseIcon       =   "frmGraphics.frx":D0410
            TabIndex        =   150
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   -68520
            MouseIcon       =   "frmGraphics.frx":D071A
            TabIndex        =   149
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text List..."
            Height          =   195
            Index           =   2
            Left            =   -67560
            MouseIcon       =   "frmGraphics.frx":D0A24
            MousePointer    =   99  'Custom
            TabIndex        =   148
            Top             =   6825
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLast 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   -68340
            MouseIcon       =   "frmGraphics.frx":D0D2E
            MousePointer    =   99  'Custom
            TabIndex        =   147
            Top             =   6825
            Width           =   300
         End
         Begin VB.Label lblPrevious 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   -69840
            MouseIcon       =   "frmGraphics.frx":D1038
            MousePointer    =   99  'Custom
            TabIndex        =   146
            Top             =   6825
            Width           =   615
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   -68940
            MouseIcon       =   "frmGraphics.frx":D1342
            MousePointer    =   99  'Custom
            TabIndex        =   145
            Top             =   6825
            Width           =   345
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   -70020
            MouseIcon       =   "frmGraphics.frx":D164C
            TabIndex        =   144
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   -69120
            MouseIcon       =   "frmGraphics.frx":D1956
            TabIndex        =   143
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Left            =   -68520
            MouseIcon       =   "frmGraphics.frx":D1C60
            TabIndex        =   142
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text List..."
            Height          =   195
            Index           =   1
            Left            =   -67560
            MouseIcon       =   "frmGraphics.frx":D1F6A
            MousePointer    =   99  'Custom
            TabIndex        =   141
            Top             =   6825
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLast 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   -68340
            MouseIcon       =   "frmGraphics.frx":D2274
            MousePointer    =   99  'Custom
            TabIndex        =   140
            Top             =   6825
            Width           =   300
         End
         Begin VB.Label lblPrevious 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Previous"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   5160
            MouseIcon       =   "frmGraphics.frx":D257E
            MousePointer    =   99  'Custom
            TabIndex        =   106
            Top             =   6825
            Width           =   615
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   6060
            MouseIcon       =   "frmGraphics.frx":D2888
            MousePointer    =   99  'Custom
            TabIndex        =   105
            Top             =   6825
            Width           =   345
         End
         Begin VB.Label lblFirst 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "First"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   4560
            MouseIcon       =   "frmGraphics.frx":D2B92
            MousePointer    =   99  'Custom
            TabIndex        =   104
            Top             =   6825
            Width           =   315
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Left            =   4980
            MouseIcon       =   "frmGraphics.frx":D2E9C
            TabIndex        =   103
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Left            =   5880
            MouseIcon       =   "frmGraphics.frx":D31A6
            TabIndex        =   102
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblPipe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "|"
            Enabled         =   0   'False
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
            Left            =   6480
            MouseIcon       =   "frmGraphics.frx":D34B0
            TabIndex        =   101
            Top             =   6795
            Width           =   90
         End
         Begin VB.Label lblList 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text List..."
            Height          =   195
            Index           =   0
            Left            =   7440
            MouseIcon       =   "frmGraphics.frx":D37BA
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   6825
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLast 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Last"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   6660
            MouseIcon       =   "frmGraphics.frx":D3AC4
            MousePointer    =   99  'Custom
            TabIndex        =   99
            Top             =   6825
            Width           =   300
         End
         Begin VB.Label lblFileCount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   -63660
            TabIndex        =   87
            Top             =   585
            Width           =   45
         End
         Begin VB.Label lblClient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   86
            Top             =   795
            Width           =   465
         End
         Begin VB.Label lblInactive 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "INACTIVE Graphics are denoted with a Red Background"
            Height          =   195
            Index           =   3
            Left            =   -70440
            TabIndex        =   85
            Top             =   855
            Visible         =   0   'False
            Width           =   3990
         End
         Begin VB.Label lblInactive 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "INACTIVE Graphics"
            Height          =   195
            Index           =   2
            Left            =   -70380
            TabIndex        =   84
            Top             =   855
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblInactive 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "INACTIVE Graphics"
            Height          =   195
            Index           =   0
            Left            =   8640
            TabIndex        =   83
            Top             =   6855
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lblInactive 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "INACTIVE Graphics"
            Height          =   195
            Index           =   1
            Left            =   -65700
            TabIndex        =   82
            Top             =   6795
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   81
            Top             =   555
            Width           =   465
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Index           =   0
            Left            =   -74820
            TabIndex        =   80
            Top             =   555
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client:"
            Height          =   195
            Left            =   -73860
            TabIndex        =   79
            Top             =   555
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show Year:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   78
            Top             =   555
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   77
            Top             =   1185
            Width           =   450
         End
      End
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   1380
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2340
      ScaleHeight     =   1065
      ScaleWidth      =   4665
      TabIndex        =   5
      Top             =   4560
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
         TabIndex        =   6
         Top             =   240
         Width           =   4275
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D3DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D3EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D409B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D43B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D499C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D4F36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D5250
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D556A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D5884
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D5B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D5EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D6452
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D69EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D6F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D7520
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D7ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D8054
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D85EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D8908
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D8EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D943C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D99D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":D9F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DA50A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   -500
      Width           =   1575
   End
   Begin VB.PictureBox picMenu2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   1260
      TabIndex        =   176
      Top             =   1170
      Visible         =   0   'False
      Width           =   1260
      Begin VB.Label lblBackground 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "White Canvas"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmGraphics.frx":DAAA4
         MousePointer    =   99  'Custom
         TabIndex        =   194
         Top             =   780
         Width           =   1005
      End
      Begin VB.Label lblKeyEdit 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmGraphics.frx":DADAE
         MousePointer    =   99  'Custom
         TabIndex        =   179
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label lblSettings 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Settings..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmGraphics.frx":DB0B8
         MousePointer    =   99  'Custom
         TabIndex        =   178
         Top             =   420
         Width           =   765
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   165
         MouseIcon       =   "frmGraphics.frx":DB3C2
         MousePointer    =   99  'Custom
         TabIndex        =   177
         Top             =   60
         Width           =   675
      End
      Begin VB.Image imgKeyEdit 
         Height          =   360
         Left            =   0
         Picture         =   "frmGraphics.frx":DB6CC
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Image imgBackground 
         Height          =   360
         Left            =   0
         Picture         =   "frmGraphics.frx":DB9EE
         Top             =   720
         Width           =   1260
      End
      Begin VB.Image imgFullSize 
         Height          =   360
         Left            =   0
         Picture         =   "frmGraphics.frx":DBD10
         Top             =   360
         Width           =   1260
      End
      Begin VB.Image imgResize 
         Height          =   360
         Left            =   0
         Picture         =   "frmGraphics.frx":DC032
         Top             =   0
         Width           =   1260
      End
   End
   Begin MSComctlLib.ImageList imlZoomMode 
      Left            =   420
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DC354
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DCA7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DD1A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DD8D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DDFFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DE726
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DEE50
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DF57A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":DFCA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E03CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E0AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E1222
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E194C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E1EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E2480
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGraphics.frx":E2A1A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPDF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   1260
      ScaleHeight     =   3795
      ScaleWidth      =   9075
      TabIndex        =   195
      Top             =   1200
      Visible         =   0   'False
      Width           =   9075
      Begin VB.TextBox txtRed 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5160
         MultiLine       =   -1  'True
         TabIndex        =   202
         Top             =   1200
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.VScrollBar vsc1 
         Height          =   1995
         Left            =   8760
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin XpdfViewerCtl.XpdfViewer Xpdf1 
         Height          =   1275
         Left            =   0
         TabIndex        =   196
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
         showScrollbars  =   -1  'True
         showBorder      =   0   'False
         showPasswordDialog=   0   'False
      End
      Begin VB.PictureBox picROuter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   0
         ScaleHeight     =   2595
         ScaleWidth      =   5835
         TabIndex        =   197
         Top             =   0
         Visible         =   0   'False
         Width           =   5835
         Begin VB.PictureBox picRed 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            DrawWidth       =   5
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1875
            Left            =   0
            MouseIcon       =   "frmGraphics.frx":E2FB4
            ScaleHeight     =   1875
            ScaleWidth      =   5325
            TabIndex        =   198
            Top             =   0
            Width           =   5325
            Begin VB.Label lblRed 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   0
               Left            =   60
               MouseIcon       =   "frmGraphics.frx":E32BE
               MousePointer    =   99  'Custom
               TabIndex        =   200
               Top             =   60
               Visible         =   0   'False
               Width           =   5115
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblEsc 
               AutoSize        =   -1  'True
               BackColor       =   &H0000FFFF&
               Caption         =   "Use ESC key to end"
               Height          =   195
               Left            =   600
               TabIndex        =   199
               Top             =   1680
               UseMnemonic     =   0   'False
               Visible         =   0   'False
               Width           =   1410
            End
         End
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
      Height          =   6630
      Left            =   1260
      MouseIcon       =   "frmGraphics.frx":E35C8
      ScaleHeight     =   6630
      ScaleMode       =   0  'User
      ScaleWidth      =   10440
      TabIndex        =   2
      Top             =   1200
      Width           =   10440
      Begin VB.Label lblRedNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   405
         Left            =   7800
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   9540
      ScaleHeight     =   420
      ScaleWidth      =   2100
      TabIndex        =   203
      Top             =   720
      Width           =   2100
      Begin VB.Image imgNav 
         Height          =   300
         Index           =   1
         Left            =   1740
         MouseIcon       =   "frmGraphics.frx":E38D2
         MousePointer    =   99  'Custom
         Picture         =   "frmGraphics.frx":E3BDC
         Top             =   60
         Width           =   300
      End
      Begin VB.Image imgNav 
         Height          =   300
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmGraphics.frx":E4206
         MousePointer    =   99  'Custom
         Picture         =   "frmGraphics.frx":E4510
         Top             =   60
         Width           =   300
      End
      Begin VB.Label lblNavCnt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1 of 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   204
         Top             =   60
         UseMnemonic     =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox picTools 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   12495
      TabIndex        =   211
      Top             =   8220
      Visible         =   0   'False
      Width           =   12555
      Begin VB.PictureBox picRedTools 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6900
         ScaleHeight     =   315
         ScaleWidth      =   7035
         TabIndex        =   213
         Top             =   0
         Visible         =   0   'False
         Width           =   7035
         Begin VB.PictureBox imgColor 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   1920
            Picture         =   "frmGraphics.frx":E4B3A
            ScaleHeight     =   300
            ScaleWidth      =   960
            TabIndex        =   214
            ToolTipText     =   "Click to change Annotation Color"
            Top             =   8
            Width           =   960
         End
         Begin VB.Image imgUtility 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   4740
            MouseIcon       =   "frmGraphics.frx":E5524
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":E582E
            ToolTipText     =   "Click to Delete Redline File"
            Top             =   8
            Width           =   720
         End
         Begin VB.Image imgUtility 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   4020
            MouseIcon       =   "frmGraphics.frx":E60D8
            Picture         =   "frmGraphics.frx":E63E2
            ToolTipText     =   "Click to Clear Redline File back to last Save"
            Top             =   8
            Width           =   720
         End
         Begin VB.Image imgUtility 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   3300
            MouseIcon       =   "frmGraphics.frx":E6C8C
            Picture         =   "frmGraphics.frx":E6F96
            ToolTipText     =   "Click to Save Redline File"
            Top             =   8
            Width           =   720
         End
         Begin VB.Image imgDo 
            Height          =   300
            Index           =   1
            Left            =   3240
            MouseIcon       =   "frmGraphics.frx":E7840
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":E7B4A
            ToolTipText     =   "Click to Redo"
            Top             =   15
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image imgDo 
            Height          =   300
            Index           =   0
            Left            =   2940
            MouseIcon       =   "frmGraphics.frx":E8174
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":E847E
            ToolTipText     =   "Click to Undo"
            Top             =   8
            Width           =   300
         End
         Begin VB.Image imgRedReload 
            Height          =   300
            Left            =   5520
            MouseIcon       =   "frmGraphics.frx":E8AA8
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":E8DB2
            ToolTipText     =   "Return to original file"
            Top             =   8
            Width           =   960
         End
         Begin VB.Image imgRedMode 
            Height          =   300
            Index           =   0
            Left            =   0
            MouseIcon       =   "frmGraphics.frx":E979C
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":E9AA6
            ToolTipText     =   "Click to start Sketch mode"
            Top             =   8
            Width           =   960
         End
         Begin VB.Image imgRedMode 
            Height          =   300
            Index           =   1
            Left            =   960
            MouseIcon       =   "frmGraphics.frx":EA490
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":EA79A
            ToolTipText     =   "Click to start Text mode"
            Top             =   8
            Width           =   960
         End
      End
      Begin VB.PictureBox picJPGTools 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9060
         ScaleHeight     =   315
         ScaleWidth      =   2955
         TabIndex        =   216
         Top             =   0
         Visible         =   0   'False
         Width           =   2955
         Begin VB.Label lblSize 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100%"
            ForeColor       =   &H00996633&
            Height          =   195
            Left            =   1620
            TabIndex        =   218
            Top             =   60
            Width           =   465
         End
         Begin VB.Image Image1 
            Height          =   300
            Left            =   1500
            Picture         =   "frmGraphics.frx":EB184
            Top             =   15
            Width           =   720
         End
         Begin VB.Image imgJPGZoom 
            Height          =   300
            Index           =   2
            Left            =   2220
            MouseIcon       =   "frmGraphics.frx":EBA2E
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":EBD38
            ToolTipText     =   "If larger than screen allows, Click to view Full Size"
            Top             =   8
            Width           =   480
         End
         Begin VB.Image imgJPGZoom 
            Height          =   300
            Index           =   1
            Left            =   1020
            MouseIcon       =   "frmGraphics.frx":EC452
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":EC75C
            ToolTipText     =   "Click to Maximize Image"
            Top             =   8
            Width           =   480
         End
         Begin VB.Image imgJPGZoom 
            Height          =   300
            Index           =   0
            Left            =   540
            MouseIcon       =   "frmGraphics.frx":ECE76
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":ED180
            ToolTipText     =   "Click to view as Posted Size"
            Top             =   8
            Width           =   480
         End
         Begin VB.Image imgRed 
            Height          =   300
            Index           =   0
            Left            =   0
            MouseIcon       =   "frmGraphics.frx":ED89A
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":EDBA4
            ToolTipText     =   "Click to open file in Redline Mode"
            Top             =   8
            Width           =   480
         End
      End
      Begin VB.PictureBox picPDFTools 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   300
         ScaleHeight     =   315
         ScaleWidth      =   8295
         TabIndex        =   220
         Top             =   0
         Width           =   8295
         Begin VB.PictureBox picZoom 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   540
            ScaleHeight     =   315
            ScaleWidth      =   3795
            TabIndex        =   224
            Top             =   0
            Width           =   3795
            Begin VB.ComboBox cboZoom 
               Height          =   315
               ItemData        =   "frmGraphics.frx":EE2BE
               Left            =   2340
               List            =   "frmGraphics.frx":EE2E6
               TabIndex        =   225
               Top             =   0
               Width           =   1080
            End
            Begin VB.Image imgZoom 
               Height          =   240
               Index           =   1
               Left            =   3480
               Picture         =   "frmGraphics.frx":EE338
               ToolTipText     =   "Click to zoom in"
               Top             =   30
               Width           =   240
            End
            Begin VB.Image imgZoom 
               Height          =   240
               Index           =   0
               Left            =   2040
               Picture         =   "frmGraphics.frx":EE8C2
               ToolTipText     =   "Click to zoom out"
               Top             =   30
               Width           =   240
            End
            Begin VB.Image imgPDF 
               Height          =   300
               Index           =   1
               Left            =   540
               MouseIcon       =   "frmGraphics.frx":EEE4C
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":EF156
               ToolTipText     =   "Click to fit image width"
               Top             =   8
               Width           =   480
            End
            Begin VB.Image imgPDF 
               Height          =   300
               Index           =   0
               Left            =   60
               MouseIcon       =   "frmGraphics.frx":EF870
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":EFB7A
               ToolTipText     =   "Click to fit full image"
               Top             =   8
               Width           =   480
            End
            Begin VB.Image imgPDF 
               Height          =   300
               Index           =   2
               Left            =   1020
               MouseIcon       =   "frmGraphics.frx":F0294
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":F059E
               ToolTipText     =   "Click to select a zoom window"
               Top             =   8
               Width           =   480
            End
            Begin VB.Image imgPDF 
               Height          =   300
               Index           =   3
               Left            =   1500
               MouseIcon       =   "frmGraphics.frx":F0CB8
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":F0FC2
               ToolTipText     =   "Pan mode"
               Top             =   8
               Width           =   480
            End
         End
         Begin VB.PictureBox picPDFOpts 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7440
            ScaleHeight     =   315
            ScaleWidth      =   360
            TabIndex        =   223
            Top             =   0
            Width           =   360
            Begin VB.Image imgSearchPDF 
               Height          =   300
               Index           =   0
               Left            =   60
               Picture         =   "frmGraphics.frx":F16DC
               ToolTipText     =   "Search for Text"
               Top             =   8
               Width           =   300
            End
         End
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4380
            ScaleHeight     =   315
            ScaleWidth      =   3075
            TabIndex        =   221
            Top             =   0
            Width           =   3075
            Begin VB.Image imgPage 
               Height          =   240
               Index           =   1
               Left            =   1740
               Picture         =   "frmGraphics.frx":F1D06
               ToolTipText     =   "Click to page forward"
               Top             =   30
               Width           =   240
            End
            Begin VB.Image imgPage 
               Height          =   240
               Index           =   0
               Left            =   0
               Picture         =   "frmGraphics.frx":F2290
               ToolTipText     =   "Click to page back"
               Top             =   30
               Width           =   240
            End
            Begin VB.Label lblPage 
               Alignment       =   2  'Center
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Page 1 of 1"
               ForeColor       =   &H80000007&
               Height          =   315
               Left            =   300
               TabIndex        =   222
               Top             =   0
               Width           =   1380
            End
            Begin VB.Image imgPageMode 
               Height          =   300
               Index           =   0
               Left            =   2040
               MouseIcon       =   "frmGraphics.frx":F281A
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":F2B24
               ToolTipText     =   "Click for single page mode"
               Top             =   8
               Width           =   480
            End
            Begin VB.Image imgPageMode 
               Height          =   300
               Index           =   1
               Left            =   2520
               MouseIcon       =   "frmGraphics.frx":F323E
               MousePointer    =   99  'Custom
               Picture         =   "frmGraphics.frx":F3548
               ToolTipText     =   "Click for continuous page mode"
               Top             =   8
               Width           =   480
            End
         End
         Begin VB.Image imgSearchPDF 
            Height          =   300
            Index           =   1
            Left            =   7800
            Picture         =   "frmGraphics.frx":F3C62
            ToolTipText     =   "Search for Next Text"
            Top             =   8
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Image imgRed 
            Height          =   300
            Index           =   1
            Left            =   0
            MouseIcon       =   "frmGraphics.frx":F428C
            MousePointer    =   99  'Custom
            Picture         =   "frmGraphics.frx":F4596
            ToolTipText     =   "Click to open file in Redline Mode"
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   60
         TabIndex        =   212
         Top             =   60
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin VB.ListBox lstUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      ItemData        =   "frmGraphics.frx":F4CB0
      Left            =   60
      List            =   "frmGraphics.frx":F4CB2
      TabIndex        =   215
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblRN 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   165
      MouseIcon       =   "frmGraphics.frx":F4CB4
      MousePointer    =   99  'Custom
      TabIndex        =   226
      ToolTipText     =   "Click to hide"
      Top             =   3480
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   945
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpRN 
      BorderColor     =   &H00666666&
      BorderWidth     =   2
      Height          =   3555
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblRedline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2220
      MouseIcon       =   "frmGraphics.frx":F4FBE
      MousePointer    =   99  'Custom
      TabIndex        =   217
      ToolTipText     =   "Click to open Redline File"
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Image imgUndo 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   6540
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgCur 
      Height          =   225
      Index           =   2
      Left            =   900
      Picture         =   "frmGraphics.frx":F52C8
      Top             =   5520
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgCur 
      Height          =   300
      Index           =   1
      Left            =   180
      Picture         =   "frmGraphics.frx":F583E
      Top             =   5460
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCur 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "frmGraphics.frx":F5A08
      Top             =   5340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgImporter 
      Height          =   480
      Left            =   9540
      MouseIcon       =   "frmGraphics.frx":F5D12
      MousePointer    =   99  'Custom
      Picture         =   "frmGraphics.frx":F601C
      ToolTipText     =   "Click to access the Graphic Importer"
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSearch 
      Height          =   480
      Left            =   10020
      MouseIcon       =   "frmGraphics.frx":F68E6
      MousePointer    =   99  'Custom
      Picture         =   "frmGraphics.frx":F6BF0
      ToolTipText     =   "Click to access the Search Tool"
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSupDoc 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmGraphics.frx":F74BA
      MousePointer    =   99  'Custom
      Picture         =   "frmGraphics.frx":F77C4
      ToolTipText     =   "Click to view the Supoort Document for the current Image"
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmGraphics.frx":F808E
      MousePointer    =   99  'Custom
      Picture         =   "frmGraphics.frx":F8398
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Width           =   720
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
      MouseIcon       =   "frmGraphics.frx":F8EE2
      MousePointer    =   99  'Custom
      TabIndex        =   180
      Top             =   780
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Close Graphics Handler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   11460
      MouseIcon       =   "frmGraphics.frx":F91EC
      MousePointer    =   99  'Custom
      TabIndex        =   175
      Top             =   60
      Width           =   1410
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblViewAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All"
      Height          =   195
      Index           =   3
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":F94F6
      MousePointer    =   99  'Custom
      TabIndex        =   139
      Top             =   4800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "181-200"
      Height          =   195
      Index           =   39
      Left            =   18180
      MouseIcon       =   "frmGraphics.frx":F9800
      MousePointer    =   99  'Custom
      TabIndex        =   138
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "161-180"
      Height          =   195
      Index           =   38
      Left            =   17400
      MouseIcon       =   "frmGraphics.frx":F9B0A
      MousePointer    =   99  'Custom
      TabIndex        =   137
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "141-160"
      Height          =   195
      Index           =   37
      Left            =   16620
      MouseIcon       =   "frmGraphics.frx":F9E14
      MousePointer    =   99  'Custom
      TabIndex        =   136
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "121-140"
      Height          =   195
      Index           =   36
      Left            =   15840
      MouseIcon       =   "frmGraphics.frx":FA11E
      MousePointer    =   99  'Custom
      TabIndex        =   135
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "101-120"
      Height          =   195
      Index           =   35
      Left            =   15060
      MouseIcon       =   "frmGraphics.frx":FA428
      MousePointer    =   99  'Custom
      TabIndex        =   134
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81-100"
      Height          =   195
      Index           =   34
      Left            =   14400
      MouseIcon       =   "frmGraphics.frx":FA732
      MousePointer    =   99  'Custom
      TabIndex        =   133
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "61-80"
      Height          =   195
      Index           =   33
      Left            =   13800
      MouseIcon       =   "frmGraphics.frx":FAA3C
      MousePointer    =   99  'Custom
      TabIndex        =   132
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41-60"
      Height          =   195
      Index           =   32
      Left            =   13200
      MouseIcon       =   "frmGraphics.frx":FAD46
      MousePointer    =   99  'Custom
      TabIndex        =   131
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-40"
      Height          =   195
      Index           =   31
      Left            =   12600
      MouseIcon       =   "frmGraphics.frx":FB050
      MousePointer    =   99  'Custom
      TabIndex        =   130
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1-20"
      Height          =   195
      Index           =   30
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FB35A
      MousePointer    =   99  'Custom
      TabIndex        =   129
      Top             =   4980
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblViewAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All"
      Height          =   195
      Index           =   2
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FB664
      MousePointer    =   99  'Custom
      TabIndex        =   128
      Top             =   4260
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "181-200"
      Height          =   195
      Index           =   29
      Left            =   18180
      MouseIcon       =   "frmGraphics.frx":FB96E
      MousePointer    =   99  'Custom
      TabIndex        =   127
      Top             =   4440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "161-180"
      Height          =   195
      Index           =   28
      Left            =   17400
      MouseIcon       =   "frmGraphics.frx":FBC78
      MousePointer    =   99  'Custom
      TabIndex        =   126
      Top             =   4440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "141-160"
      Height          =   195
      Index           =   27
      Left            =   16620
      MouseIcon       =   "frmGraphics.frx":FBF82
      MousePointer    =   99  'Custom
      TabIndex        =   125
      Top             =   4440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "121-140"
      Height          =   195
      Index           =   26
      Left            =   15840
      MouseIcon       =   "frmGraphics.frx":FC28C
      MousePointer    =   99  'Custom
      TabIndex        =   124
      Top             =   4440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "101-120"
      Height          =   195
      Index           =   25
      Left            =   15060
      MouseIcon       =   "frmGraphics.frx":FC596
      MousePointer    =   99  'Custom
      TabIndex        =   123
      Top             =   4440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81-100"
      Height          =   195
      Index           =   24
      Left            =   14400
      MouseIcon       =   "frmGraphics.frx":FC8A0
      MousePointer    =   99  'Custom
      TabIndex        =   122
      Top             =   4440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "61-80"
      Height          =   195
      Index           =   23
      Left            =   13800
      MouseIcon       =   "frmGraphics.frx":FCBAA
      MousePointer    =   99  'Custom
      TabIndex        =   121
      Top             =   4440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41-60"
      Height          =   195
      Index           =   22
      Left            =   13200
      MouseIcon       =   "frmGraphics.frx":FCEB4
      MousePointer    =   99  'Custom
      TabIndex        =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-40"
      Height          =   195
      Index           =   21
      Left            =   12600
      MouseIcon       =   "frmGraphics.frx":FD1BE
      MousePointer    =   99  'Custom
      TabIndex        =   119
      Top             =   4440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1-20"
      Height          =   195
      Index           =   20
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FD4C8
      MousePointer    =   99  'Custom
      TabIndex        =   118
      Top             =   4440
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblViewAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All"
      Height          =   195
      Index           =   1
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FD7D2
      MousePointer    =   99  'Custom
      TabIndex        =   117
      Top             =   3780
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "181-200"
      Height          =   195
      Index           =   19
      Left            =   18180
      MouseIcon       =   "frmGraphics.frx":FDADC
      MousePointer    =   99  'Custom
      TabIndex        =   116
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "161-180"
      Height          =   195
      Index           =   18
      Left            =   17400
      MouseIcon       =   "frmGraphics.frx":FDDE6
      MousePointer    =   99  'Custom
      TabIndex        =   115
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "141-160"
      Height          =   195
      Index           =   17
      Left            =   16620
      MouseIcon       =   "frmGraphics.frx":FE0F0
      MousePointer    =   99  'Custom
      TabIndex        =   114
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "121-140"
      Height          =   195
      Index           =   16
      Left            =   15840
      MouseIcon       =   "frmGraphics.frx":FE3FA
      MousePointer    =   99  'Custom
      TabIndex        =   113
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "101-120"
      Height          =   195
      Index           =   15
      Left            =   15060
      MouseIcon       =   "frmGraphics.frx":FE704
      MousePointer    =   99  'Custom
      TabIndex        =   112
      Top             =   3960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81-100"
      Height          =   195
      Index           =   14
      Left            =   14400
      MouseIcon       =   "frmGraphics.frx":FEA0E
      MousePointer    =   99  'Custom
      TabIndex        =   111
      Top             =   3960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "61-80"
      Height          =   195
      Index           =   13
      Left            =   13800
      MouseIcon       =   "frmGraphics.frx":FED18
      MousePointer    =   99  'Custom
      TabIndex        =   110
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41-60"
      Height          =   195
      Index           =   12
      Left            =   13200
      MouseIcon       =   "frmGraphics.frx":FF022
      MousePointer    =   99  'Custom
      TabIndex        =   109
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-40"
      Height          =   195
      Index           =   11
      Left            =   12600
      MouseIcon       =   "frmGraphics.frx":FF32C
      MousePointer    =   99  'Custom
      TabIndex        =   108
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1-20"
      Height          =   195
      Index           =   10
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FF636
      MousePointer    =   99  'Custom
      TabIndex        =   107
      Top             =   3960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblViewAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View All"
      Height          =   195
      Index           =   0
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":FF940
      MousePointer    =   99  'Custom
      TabIndex        =   98
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "181-200"
      Height          =   195
      Index           =   9
      Left            =   18180
      MouseIcon       =   "frmGraphics.frx":FFC4A
      MousePointer    =   99  'Custom
      TabIndex        =   97
      Top             =   3510
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "161-180"
      Height          =   195
      Index           =   8
      Left            =   17400
      MouseIcon       =   "frmGraphics.frx":FFF54
      MousePointer    =   99  'Custom
      TabIndex        =   96
      Top             =   3510
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "141-160"
      Height          =   195
      Index           =   7
      Left            =   16620
      MouseIcon       =   "frmGraphics.frx":10025E
      MousePointer    =   99  'Custom
      TabIndex        =   95
      Top             =   3510
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "121-140"
      Height          =   195
      Index           =   6
      Left            =   15840
      MouseIcon       =   "frmGraphics.frx":100568
      MousePointer    =   99  'Custom
      TabIndex        =   94
      Top             =   3510
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "101-120"
      Height          =   195
      Index           =   5
      Left            =   15060
      MouseIcon       =   "frmGraphics.frx":100872
      MousePointer    =   99  'Custom
      TabIndex        =   93
      Top             =   3510
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81-100"
      Height          =   195
      Index           =   4
      Left            =   14400
      MouseIcon       =   "frmGraphics.frx":100B7C
      MousePointer    =   99  'Custom
      TabIndex        =   92
      Top             =   3510
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "61-80"
      Height          =   195
      Index           =   3
      Left            =   13800
      MouseIcon       =   "frmGraphics.frx":100E86
      MousePointer    =   99  'Custom
      TabIndex        =   91
      Top             =   3510
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41-60"
      Height          =   195
      Index           =   2
      Left            =   13200
      MouseIcon       =   "frmGraphics.frx":101190
      MousePointer    =   99  'Custom
      TabIndex        =   90
      Top             =   3510
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21-40"
      Height          =   195
      Index           =   1
      Left            =   12600
      MouseIcon       =   "frmGraphics.frx":10149A
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   3510
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
      Left            =   12120
      MouseIcon       =   "frmGraphics.frx":1017A4
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   3510
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgApprovers 
      Height          =   360
      Left            =   12480
      Picture         =   "frmGraphics.frx":101AAE
      Top             =   1860
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Index           =   27
      Left            =   1260
      Picture         =   "frmGraphics.frx":101CB8
      Top             =   8700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Index           =   30
      Left            =   1800
      Picture         =   "frmGraphics.frx":102582
      Top             =   8700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Index           =   20
      Left            =   720
      Picture         =   "frmGraphics.frx":10288C
      Top             =   8700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStatus 
      Height          =   480
      Index           =   10
      Left            =   240
      Picture         =   "frmGraphics.frx":102B96
      Top             =   8700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   2
      Left            =   10020
      Picture         =   "frmGraphics.frx":102EA0
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   60
      Picture         =   "frmGraphics.frx":1031AA
      Top             =   4860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "frmGraphics.frx":1032FC
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   1
      Left            =   10080
      Picture         =   "frmGraphics.frx":103606
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   0
      Left            =   3600
      Picture         =   "frmGraphics.frx":103A48
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgComm 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "frmGraphics.frx":103E8A
      MousePointer    =   99  'Custom
      Picture         =   "frmGraphics.frx":104194
      Top             =   660
      Visible         =   0   'False
      Width           =   480
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
      Left            =   1800
      TabIndex        =   1
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Image imgSize 
      Height          =   735
      Left            =   120
      Top             =   7260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Graphics Viewer is loading..."
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
      Left            =   1080
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   3150
   End
   Begin VB.Image imgClose 
      Height          =   945
      Left            =   11820
      Picture         =   "frmGraphics.frx":1045D6
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgMenu 
      Height          =   570
      Left            =   0
      Picture         =   "frmGraphics.frx":104B7F
      Top             =   600
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgBack 
      Height          =   735
      Index           =   1
      Left            =   9420
      Top             =   0
      Width           =   1515
   End
   Begin VB.Image imgBack 
      Height          =   855
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   5955
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "mnuCheck"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckUsage 
         Caption         =   "Check Usage..."
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel01 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuGfxApproval 
         Caption         =   "Graphic Approval Interface..."
      End
      Begin VB.Menu mnuAssign 
         Caption         =   "Assign Current Image to a Show..."
      End
      Begin VB.Menu mnuRedlining 
         Caption         =   "Annotation"
         Begin VB.Menu mnuGRedlines 
            Caption         =   "Redlines"
            Begin VB.Menu mnuGRedLoad 
               Caption         =   "Load"
            End
            Begin VB.Menu mnuGRedSave 
               Caption         =   "Save"
            End
            Begin VB.Menu mnuGRedClear 
               Caption         =   "Clear"
            End
            Begin VB.Menu mnuGRedDelete 
               Caption         =   "Delete"
            End
         End
         Begin VB.Menu mnuGRedSketch 
            Caption         =   "  Redline ""Sketch"" Mode"
         End
         Begin VB.Menu mnuGRedText 
            Caption         =   "  Redline ""Text"" Mode"
         End
         Begin VB.Menu mnuGRedEnd 
            Caption         =   "  End Redline Mode"
         End
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResizeGraphic 
         Caption         =   "Resize Graphic to Actual Size"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaxGraphic 
         Caption         =   "Maximize Graphic"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGPrint 
         Caption         =   "Print"
         Index           =   0
      End
      Begin VB.Menu mnuGPrint 
         Caption         =   "Print w/Annotation"
         Index           =   1
      End
      Begin VB.Menu mnuSendALink 
         Caption         =   "Send-A-Link..."
      End
      Begin VB.Menu mnuEmailSel 
         Caption         =   "Email a Copy of File..."
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download a Copy of File..."
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGFXData2 
         Caption         =   "View Graphic Data..."
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuCancel02 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuShowData 
      Caption         =   "mnuShowData"
      Visible         =   0   'False
      Begin VB.Menu mnuShowName 
         Caption         =   "Show Name:"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowLoc 
         Caption         =   "Show Loc:"
      End
      Begin VB.Menu mnuShowOpen 
         Caption         =   "Show Open:"
      End
      Begin VB.Menu mnuShowClose 
         Caption         =   "Show Close:"
      End
      Begin VB.Menu mnuDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuGfx 
      Caption         =   "mnuGFX"
      Visible         =   0   'False
      Begin VB.Menu mnuGFXData 
         Caption         =   "View Graphic Data..."
      End
      Begin VB.Menu mnuCheckUse 
         Caption         =   "Check Usage..."
      End
      Begin VB.Menu mnuCommThumb 
         Caption         =   "Comments..."
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "mnuSettings"
      Visible         =   0   'False
      Begin VB.Menu mnuSelOptions 
         Caption         =   "Graphic Viewer File Options"
      End
      Begin VB.Menu mnuDash05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelByName 
         Caption         =   "Select Graphics by Filename"
      End
      Begin VB.Menu mnuSelByImage 
         Caption         =   "Select Graphics from High-Res ""Thumbnails"""
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuSelByImage 
         Caption         =   "Select Graphics from Low-Res ""Thumbnails"""
         Index           =   3
      End
      Begin VB.Menu mnuDash06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuResetStatus 
      Caption         =   "mnuResetStatus"
      Visible         =   0   'False
      Begin VB.Menu mnuName 
         Caption         =   "Graphic Name"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuName 
         Caption         =   "Graphic Name"
         Index           =   1
      End
      Begin VB.Menu mnuDash09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusReset 
         Caption         =   "Reset Status to 'DRAFT' and restrict Client Viewing"
         Index           =   0
      End
      Begin VB.Menu mnuStatusReset 
         Caption         =   "Set Status as 'RELEASED', allowing Client Viewing"
         Index           =   1
      End
      Begin VB.Menu mnuStatusReset 
         Caption         =   "Advance Status up to 'APPROVED'"
         Index           =   2
      End
      Begin VB.Menu mnuDash07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusCancel 
         Caption         =   "Reset Status as 'CANCELED'"
      End
      Begin VB.Menu mnuDash08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open in Full Viewer..."
      End
      Begin VB.Menu mnuCancel03 
         Caption         =   "Close Menu without Resetting Status"
      End
   End
   Begin VB.Menu mnuPart 
      Caption         =   "mnuPart"
      Visible         =   0   'False
      Begin VB.Menu mnuPartData 
         Caption         =   "View Part Data..."
      End
   End
   Begin VB.Menu mnuDownloadMulti 
      Caption         =   "mnuDownloadMulti"
      Visible         =   0   'False
      Begin VB.Menu mnuDownloadMode 
         Caption         =   "Activate Download Mode"
      End
      Begin VB.Menu mnuDownloadSels 
         Caption         =   "Download Selections..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEmailMulti 
      Caption         =   "mnuEmailMulti"
      Visible         =   0   'False
      Begin VB.Menu mnuEmailMode 
         Caption         =   "Activate Email Copy Mode"
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
      Begin VB.Menu mnuDash_Opt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmailSels2 
         Caption         =   "Email Copy of Selections..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDownloadSels2 
         Caption         =   "Download Selections..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "mnuText"
      Visible         =   0   'False
      Begin VB.Menu mnuTextEdit 
         Caption         =   "Edit Text..."
      End
      Begin VB.Menu mnuTextClear 
         Caption         =   "Clear Note"
      End
      Begin VB.Menu mnuTextColor 
         Caption         =   "Reset Color to Current"
      End
      Begin VB.Menu mnuTextDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bUndoCleared As Boolean
Dim strSearchText As String
Dim iUndoIndex As Integer, iUndoListIndex As Integer, iCurrUndo As Integer, iUndoMin As Integer
Dim lUndoGID As Long

Dim picCurrentRed As PictureBox

Dim x0 As Long, y0 As Long

Dim dZF As Double
Dim iImage(0 To 4) As Integer

Dim lAnnoColor As Long
Public iAnnoColor As Integer

Dim lBGColor(0 To 1) As Long
Dim iBGColor As Integer
Dim iRed As Integer

Dim iPageMode As Integer ''0=SinglePage, 1=ContinuousPage''
Dim iZoomMode As Integer ''0=FitPage, 1=FitWidth, 2=ZoomWindow, 3=Pan''
Dim iRedMode As Integer ''0=None, 1=Sketch, 2=Text''

Dim bZWindow As Boolean, bPan As Boolean, bRedLine As Boolean, bRedded As Boolean, _
            bRedText As Boolean, bRedMode As Boolean, bSelMode As Boolean, bRedSaved As Boolean

Dim panX As Long, panY As Long

Dim xs As Long, xE As Long, ys As Long, yE As Long
Dim pxS As Double, pyS As Double, pxE As Double, pyE As Double

Dim hhkLowLevelKybd As Long

'''Private pdfGraphic As AcroPDFLibCtl.AcroPDF
Dim bAcro As Boolean, bResizing As Boolean

Dim bPopped(0 To 4) As Boolean

Public bAddMode As Boolean

Dim iApprovedIndex As Integer
Public iApprovalRow As Integer
Dim sNoShows As String
Dim lBColor(0 To 30) As Long
Dim lBCC_Def As Long
Dim sFBCN_Def As String
Dim bApproveDown As Boolean
Dim maxX As Double, maxY As Double, dTop As Double, dLeft As Double, _
            dGTop As Double, dGLeft As Double, _
            dMTop As Double, dMLeft As Double, dSTop As Double, dSLeft As Double
Dim iAccess As Integer
Dim rAsp As Double, rFAsp As Double, rX As Double, rY As Double, rXO As Double, rYO As Double, _
            rMX As Double, rMY As Double, rSX As Double, rSY As Double
Dim bMenuButton As Boolean
Dim tSHYR As Integer, CurrSHYR As Integer, fSHYR(0 To 4) As Integer
Dim tBCC As String, fBCC(0 To 4) As String
Dim tSHCD As Long, fSHCD(0 To 4) As Long
Dim tFBCN As String, fFBCN(0 To 4) As String
Dim tSHNM As String, fSHNM(0 To 4) As String
Dim lFID As Long
Dim sTable As String
Dim CurrNode As String, CurrNodeText As String, sTabDesc As String, sMessPath As String, TNode As String
Dim RTXFile As String, RedFile As String, CurrFile As String, sGPath As String, RedName As String, _
            RedMess As String, SaveMess As String
Dim xStr As Single, yStr As Single
Dim bRedding As Boolean, RedMode As Boolean, bTexting As Boolean, TextMode As Boolean
Dim lRedID As Long, lRefID As Long
Dim bTeam As Boolean, bPicLoaded As Boolean, bApprover As Boolean
Dim redBCC As String
Dim redSHCD As Long
Dim lNewLockId As Long
Public CurrIndex As Integer
Dim CurrParNode(0 To 3) As String, CurrParText(0 To 3) As String
'Public CurrSelect(0 To 3) As String
Dim iImageState As Integer ''0=Small:1=Max''
Dim GFXCUNO() As Long
Dim iCurrStatus As Integer
Dim bDataSort As Boolean
Public sOrder As String, sIN As String
Dim bStatusReset As Boolean
Dim rCommX1 As Single, rCommX2 As Single, rApproverX1 As Single, rApproverX2 As Single
Dim bTabPop(0 To 4) As Boolean
Dim bReSize As Boolean, bFromFloorplan As Boolean
Dim iUnit As Integer
Dim iSSSort As Integer
Dim sInType(0 To 4) As String
Dim iTabType(0 To 4) As Integer
Public iTabStatus As Integer
Dim sPassedGNode As String, sPassedFBCN As String
Dim sPassedBCC As Long
Public iApproverView As Integer
Public bResetting As Boolean
Public sSearchList As String
Dim rApproverX As Single, rApproverY As Single
Dim bRighted As Boolean
Dim iListStart(0 To 3) As Integer, iGFXCount(0 To 3) As Integer
Dim bEMode As Boolean, bDMode As Boolean
Dim iModeTab As Integer
Dim bApprovedOnly As Boolean
Public bDirsOpen As Boolean
Public pDownloadPath As String

'''Private shlShell As Shell32.Shell
'''Private shlFolder As Shell32.FOLDER
'''Private Const BIF_RETURNONLYFSDIRS = &H1

Const panSpeed = 2
Const iUndoMax = 5 '' 20 ''10

Public Property Get PassBCC() As Long
    PassBCC = sPassedBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    sPassedBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = sPassedFBCN
End Property
Public Property Let PassFBCN(ByVal vNewValue As String)
    sPassedFBCN = vNewValue
End Property

Public Property Get PassGNode() As String
    PassGNode = sPassedGNode
End Property
Public Property Let PassGNode(ByVal vNewValue As String)
    sPassedGNode = vNewValue
End Property

Public Property Get PassDLPath() As String
    PassDLPath = pDownloadPath
End Property
Public Property Let PassDLPath(ByVal vNewValue As String)
    pDownloadPath = vNewValue
End Property





'''
'''Public Property Get PassSHNM() As String
'''    PassSHNM = tSHNM
'''End Property
'''Public Property Let PassSHNM(ByVal vNewValue As String)
'''    tSHNM = vNewValue
'''End Property
'''
'''Public Property Get PassSHYR() As Integer
'''    PassSHYR = tSHYR
'''End Property
'''Public Property Let PassSHYR(ByVal vNewValue As Integer)
'''    tSHYR = vNewValue
'''End Property
'''
'''Public Property Get PassSHCD() As Long
'''    PassSHCD = tSHCD
'''End Property
'''Public Property Let PassSHCD(ByVal vNewValue As Long)
'''    tSHCD = vNewValue
'''End Property





Private Sub cboASHCD_Click()
    Dim tFolder As String
    
    If cboASHCD.Text = "" Then Exit Sub
    If cboCUNO(4).Text <> "" Then cmdRefresh.Enabled = True
    fSHCD(4) = cboASHCD.ItemData(cboASHCD.ListIndex)
'''    Me.MousePointer = 11
'''    flxApprove.Visible = False: picOuter(4).Visible = False
'''    If fSHCD(4) = 0 Then
'''        cboCUNO(4).Text = cboCUNO(4).Text
'''    Else
'''        Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, fSHYR(4), fSHCD(4))
'''    End If

    ''NEED TO VERIFY WHETHER USER IS MAND-RECIP FOR CLIENT-SHOW''
    If fSHCD(4) <> 0 Then
        Call CheckIfClientShowGfxMandRecip(CLng(fBCC(4)), fSHCD(4), UserID)
    Else
        Call CheckIfGfxMandRecip(CLng(fBCC(4)), UserID)
    End If
    
    tFolder = cboFolder.Text
    
    Call GetApprovalFolders(CLng(fBCC(4)), fSHCD(4), fSHYR(4))
    
    If tFolder <> "" Then
        On Error Resume Next
        cboFolder.Text = tFolder
    End If
                
    
    
'''    flxApprove.Visible = True: picOuter(4).Visible = True
'''    Me.MousePointer = 0
End Sub


Private Sub cboCUNO_Click(Index As Integer)
    Dim i As Integer
    Dim bCheck As Boolean
    Dim tFolder As String
    
    
    Screen.MousePointer = 11
    On Error Resume Next
    If cboCUNO(Index).Text <> "" Then
        fBCC(Index) = Right("00000000" & cboCUNO(Index).ItemData(cboCUNO(Index).ListIndex), 8)
        fFBCN(Index) = cboCUNO(Index).List(cboCUNO(Index).ListIndex)
'''        tFBCN = GetBCN(tBCC)
'        For i = 0 To cboCUNO.Count - 1
'            If i <> Index Then
'                If cboCUNO(i).Text = "" Then cboCUNO(i).Text = tFBCN
'            End If
'        Next i
        Screen.MousePointer = 11
        Select Case Index
            Case 0
                tvwGraphics(0).Nodes.Clear
                Call ClearThumbnails0(0)
                Call GetShows(cboSHCD, fSHYR(Index), fBCC(Index))
            Case 1
                Call ClearThumbnails1(0)
                Call LoadClientShows(CLng(fBCC(Index)), fSHYR(Index), iSSSort)
            Case 2
                Call ClearThumbnails2(0)
                Call PopInventory(CLng(fBCC(Index)))
            Case 3
                Call ClearThumbnails3(0)
                Call GetGraphicList(CLng(fBCC(Index)))
            Case 4
                Call CheckIfGfxMandRecip(CLng(fBCC(Index)), UserID)
''''''''                imgSearch.Enabled = True
                lblSearch.Enabled = True
                bApprover = True
'                flxApprove.Visible = False
                flxApprove.Rows = 1
                picOuter(4).Visible = False: picReview.Visible = False
                fSHCD(4) = 0
                fSHYR(4) = 0
                lFID = 0
                fraMulti.Visible = False
                
'                Call GetApprovalGraphics(CLng(fBCC(Index)), sOrder, 0, 0)
                Call GetApprovalShowYears(CLng(fBCC(Index)))
                
                If cboSHYR(4).ListCount = 1 Then
                    cboSHYR(4).Text = cboSHYR(4).List(0)
                ElseIf cboSHYR(4).ListCount > 1 Then
                    On Error Resume Next
                    cboSHYR(4).Text = CurrSHYR
                End If
                
                cmdRefresh.Enabled = True
'''                tFolder = cboFolder.Text
                
                Call GetApprovalFolders(CLng(fBCC(Index)), fSHCD(4), fSHYR(4))
'''                If tFolder <> "" Then
'''                    On Error Resume Next
'''                    cboFolder.Text = tFolder
'''                End If
                
'                flxApprove.Visible = True: picOuter(4).Visible = True
                
                cmdStatusEdit_View.Enabled = True
'''                For i = 0 To flxApprove.Rows - 2
'''                    If lblStat(i).Caption <> lblStat(0) Then
'''                        cmdStatusEdit_View.Enabled = False
'''                        Exit For
'''                    End If
'''                Next i
                
        End Select
'        If bGFXReviewer Then
'            If CheckGFXCUNO(CLng(tBCC)) Then
'                picReview.Width = 6735
'                bApprover = True
'            Else
'                picReview.Width = 2400
'                bApprover = False
'            End If
'        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cboFolder_Click()
    If cboFolder.ListIndex = -1 Then
        lFID = 0
    Else
        lFID = cboFolder.ItemData(cboFolder.ListIndex)
        If cboCUNO(4).Text <> "" Then cmdRefresh.Enabled = True
    End If
    
End Sub

Private Sub cboSHCD_Change()
    If cboSHCD.Text <> "" Then
        Call ClearThumbnails0(0)
        fSHCD(0) = cboSHCD.ItemData(cboSHCD.ListIndex)
        fSHNM(0) = cboSHCD.Text
        If CheckIfFutureShow(fSHYR(0), fSHCD(0)) Then
            Call PopShowGraphics(fBCC(0), fSHYR(0), fSHCD(0), True)
        Else
            Call PopShowGraphics(fBCC(0), fSHYR(0), fSHCD(0), False)
        End If
'''        cmdAssign.Enabled = True
    Else
'''        cmdAssign.Enabled = False
    End If
End Sub

Private Sub cboSHCD_Click()
    If cboSHCD.Text <> "" Then
        Call ClearThumbnails0(0)
        fSHCD(0) = cboSHCD.ItemData(cboSHCD.ListIndex)
        fSHNM(0) = cboSHCD.Text
        fBCC(0) = Right("00000000" & cboCUNO(0).ItemData(cboCUNO(0).ListIndex), 8)
        fFBCN(0) = cboCUNO(0).List(cboCUNO(0).ListIndex)
        If CheckIfFutureShow(fSHYR(0), fSHCD(0)) Then
            Call PopShowGraphics(fBCC(0), fSHYR(0), fSHCD(0), True)
        Else
            Call PopShowGraphics(fBCC(0), fSHYR(0), fSHCD(0), False)
        End If
'''        cmdAssign.Enabled = True
    Else
'''        cmdAssign.Enabled = False
    End If
End Sub

'Private Sub cboSHYR_Change(Index As Integer)
'    If cboSHYR(Index).Text <> "" Then
'        If Index = 0 Then
'            Call ClearThumbnails0(0)
'            cboSHCD.Clear
'        Else
'            Call ClearThumbnails1(0)
'        End If
'        tvwGraphics(Index).Nodes.Clear
'        tSHYR = CInt(cboSHYR(Index).Text)
'        Call GetShowClients(cboCUNO(Index), tSHYR)
'    End If
'End Sub

Private Sub cboSHYR_Click(Index As Integer)
    Dim sClient As String
    If cboSHYR(Index).Text <> "" Then
        If cboCUNO(4).Text <> "" Then cmdRefresh.Enabled = True
        Select Case Index
            Case 0, 1
                sClient = cboCUNO(Index).Text
                If Index = 0 Then
                    Call ClearThumbnails0(0)
                    cboSHCD.Clear
                Else
                    Call ClearThumbnails1(0)
                End If
                cboSHCD.Clear
                tvwGraphics(Index).Nodes.Clear
'''                tSHYR = CInt(cboSHYR(Index).Text)
                fSHYR(Index) = CInt(cboSHYR(Index).Text)
                Call GetShowClients(cboCUNO(Index), fSHYR(Index))
                
                On Error Resume Next
                If sClient <> "" Then cboCUNO(Index).Text = sClient
                
            Case 4
                fSHYR(Index) = CInt(cboSHYR(Index).Text)
                Call GetApprovalShows(CLng(fBCC(Index)), fSHYR(Index))
                
        End Select
    End If
End Sub

Private Sub cboZoom_Change()
    If bZWindow And dZF < xpdf1.zoomPercent Then
'        cmdView(2).UseMaskColor = True
'        cmdMode(0).UseMaskColor = True
'        cmdMode(1).UseMaskColor = False
'        cmdMode(2).UseMaskColor = True
        SetZoomMode (3)
'        bZWindow = False
'        bPan = True
'        xpdf1.enableMouseEvents = CBool(1)
'        xpdf1.enableSelect = False
'        xpdf1.mouseCursor = imgCur(1).Picture
'        Call cmdMode_Click(1)
    Else
        If xpdf1.zoomPercent >= 500 Then
            imgPDF(2).Picture = imlZoomMode.ListImages(7).Picture
            imgPDF(2).Enabled = False
            imgZoom(1).Picture = imlZoomMode.ListImages(15).Picture
            imgZoom(1).Enabled = False
        Else
            imgZoom(1).Picture = imlZoomMode.ListImages(16).Picture
            imgZoom(1).Enabled = True
        End If
    End If
End Sub

Private Sub cboZoom_Click()
    Dim dZoom As Double
    
    If cboZoom.Text = "fit page" Then
        Call imgPDF_Click(0)
'''        xpdf1.Zoom = xpdf1.zoomPage
'''        cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
    ElseIf cboZoom.Text = "fit width" Then
        Call imgPDF_Click(1)
'''        xpdf1.Zoom = xpdf1.zoomWidth
'''        cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
    Else
        dZoom = CDbl(Left(cboZoom.Text, Len(cboZoom.Text) - 1))
        xpdf1.Zoom = dZoom
    End If
    
    If xpdf1.zoomPercent >= 500 Then
        imgPDF(2).Picture = imlZoomMode.ListImages(7).Picture
        imgPDF(2).Enabled = False
        imgZoom(1).Picture = imlZoomMode.ListImages(15).Picture
        imgZoom(1).Enabled = False
    Else
        imgZoom(1).Picture = imlZoomMode.ListImages(16).Picture
        imgZoom(1).Enabled = True
    End If
End Sub

Private Sub chk0_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub chk1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub chk2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub chk3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub chk4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Me.PopupMenu mnuOptArray
End Sub

Private Sub chkApproved_Click(Index As Integer)
    Dim i As Integer
    
    bApprovedOnly = CBool(chkApproved(Index).Value)
    For i = 1 To 3
        chkApproved(i).Value = chkApproved(Index).Value
    Next i
    Select Case bApprovedOnly
        Case True
            defSIN = "30"
        Case False
            If bGPJ Then defSIN = "10, 20, 27, 30" Else defSIN = "20, 27, 30"
    End Select
    
    If Index = iApprovedIndex Then
        On Error Resume Next
        Err.Clear
        i = tvwGraphics(iApprovedIndex).SelectedItem.Index
        If Err Then Exit Sub
'        MsgBox tvwGraphics(iApprovedIndex).SelectedItem.key
        Call tvwGraphics_NodeClick(iApprovedIndex, tvwGraphics(iApprovedIndex).SelectedItem)
    End If
End Sub

Private Sub chkApproved_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    iApprovedIndex = Index
End Sub

Private Sub chkClose_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        chkClose(i).Value = chkClose(Index).Value
    Next i
End Sub


Private Sub cmdFindNext_Click()
    xpdf1.findNext strSearchText
    Debug.Print "Clicked"
End Sub

Public Sub cmdRefresh_Click()
    Dim i As Integer
    
    Me.MousePointer = 11
    
    If optApproverView(0).Value = False And optApproverView(1).Value = False Then
        bResetting = True
        optApproverView(0).Value = True
        bResetting = False
    End If
    
    flxApprove.Visible = False: picOuter(4).Visible = False
    Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, fSHYR(4), fSHCD(4), lFID)
    
    flxApprove.Visible = True: picOuter(4).Visible = True
    picReview.Visible = True
    Me.MousePointer = 0

    cmdRefresh.Enabled = False
End Sub

Public Sub ToggleBackground()
    Dim lErr As Long
    
    lErr = LockWindowUpdate(Me.hwnd)
    If Me.BackColor = vbWhite Then
'        Set Me.Picture = imlSkins.ListImages(1).Picture
        Me.BackColor = vbBlack
'        picJPG.BackColor = vbBlack
'        lblStatus.ForeColor = vbWhite
        lblGraphic.ForeColor = vbWhite
        picMenu2.BackColor = vbBlack
''''''''        lblSettings.ForeColor = vbWhite
        lblBackground.Caption = "White Canvas"
'        picViewer.BackColor = vbBlack
        xpdf1.matteColor = vbBlack
''        picRed.BackColor = vbBlack
''        picJPG.BackColor = vbBlack
        
        picNav.BackColor = vbBlack
        lblNavCnt.ForeColor = vbWhite
'''''''        linNav.BorderColor = vbWhite
        
    Else
'        Set Me.Picture = imlSkins.ListImages(2).Picture
        Me.BackColor = vbWhite
        picMenu2.BackColor = vbWhite
'        picJPG.BackColor = vbWhite
'        lblStatus.ForeColor = vbBlack
        lblGraphic.ForeColor = vbBlack
''''''''        lblSettings.ForeColor = vbBlack
        lblBackground.Caption = "Black Canvas"
'        picViewer.BackColor = vbWhite
        xpdf1.matteColor = vbWhite
''        picRed.BackColor = vbWhite
''        picJPG.BackColor = vbWhite
        
        picNav.BackColor = vbWhite
        lblNavCnt.ForeColor = vbBlack
''''''''        linNav.BorderColor = vbBlack
    End If
    lErr = LockWindowUpdate(0)
End Sub

Private Sub Form_Click()
'    Call ToggleBackground
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Debug.Print "Form KeyPress"
End Sub

'''Private Sub imgBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case Index
''''''        Case 0
''''''            Select Case bDirsOpen
''''''                Case True: Set imgDirs.Picture = imlDirs.ListImages(3).Picture
''''''                Case False: Set imgDirs.Picture = imlDirs.ListImages(1).Picture
''''''            End Select
'''        Case 1
'''            Set imgSearch.Picture = imlDirs.ListImages(5).Picture
'''            Set imgSupDoc.Picture = imlDirs.ListImages(7).Picture
'''            lblImporter.ForeColor = lGeo_Back
'''    End Select
'''End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''''''    lblSettings.ForeColor = lGeo_Back
    lblClose.ForeColor = lGeo_Back
End Sub

Private Sub imgColor_Click()
    Dim lLeft As Long, lTop As Long
    
    lLeft = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + picTools.Left + picRedTools.Left + imgColor.Left
    lTop = Me.Top + ((Me.Height - Me.ScaleHeight) - ((Me.Width - Me.ScaleWidth) / 2)) _
                + picTools.Top + picRedTools.Top + imgColor.Top
    frmPalette.PassLeft = lLeft
    frmPalette.PassTop = lTop
    frmPalette.PassColor = iAnnoColor
    
    frmPalette.Show 1, Me
    
    lAnnoColor = QBColor(iAnnoColor)
    picJPG.ForeColor = lAnnoColor
    picRed.ForeColor = lAnnoColor
    imgColor.BackColor = lAnnoColor
    
End Sub

Private Sub imgDo_Click(Index As Integer)
    Select Case Index
        Case 0 ''UNDO''
'''            If iCurrUndo > 0 Then Call ResetUndo(iCurrUndo)
            Call ResetUndo(lstUndo.ListCount - 1) ''iCurrUndo)
            If lstUndo.ListCount = 1 Then
                imgDo(0).Picture = imlRedMode.ListImages(8).Picture
                imgDo(0).Enabled = False
            End If
            
            
            
''''''            Select Case UCase(lstUndo.List(lstUndo.ListCount - 1))
''''''                Case "LBLRED"
''''''                Case Else
''''''                    imgUndo(lstUndo.ItemData(lstUndo.ListCount - 1)).Picture = LoadPicture("")
''''''                    lstUndo.RemoveItem (lstUndo.ListCount - 1)
''''''            End Select
'''''
'''''
'''''            If lstUndo.ListCount - 1 >= 0 Then
'''''                Call ResetUndo(lstUndo.ListCount - 1)
'''''                lstUndo.RemoveItem (lstUndo.ListCount - 1)
'''''            End If
''''''''            iUndoListIndex = iUndoListIndex - 1
''''''''            Debug.Print "iUndoListIndex = " & iUndoListIndex
''''''''            Call ResetUndo(iUndoListIndex)
'''''
'''''            If lstUndo.ListCount = 0 Then
'''''                imgDo(0).Picture = imlRedMode.ListImages(8).Picture
'''''                imgDo(0).Enabled = False
'''''            End If
    End Select
    
'    MsgBox iUndoListIndex
End Sub

Private Sub imgJPGZoom_Click(Index As Integer)
    Select Case Index
        Case 0 ''RESIZE''
            mnuResizeGraphic_Click
        Case 1 ''MAXIMIZE''
            mnuMaxGraphic_Click
        Case 2 ''FULLSIZE''
            frmHTMLViewer.PassFile = CurrFile
            frmHTMLViewer.PassFrom = Me.Name
            frmHTMLViewer.PassHDR = lblWelcome.Caption
            frmHTMLViewer.PassGID = lGID
            frmHTMLViewer.Show 1, Me
    End Select
End Sub

Private Sub imgMenu_Click()
    Me.PopupMenu mnuRightClick, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
End Sub

Private Sub imgNav_Click(Index As Integer)
    Dim tGID As Long
    Dim tDesc As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iOld As Integer
    
    If bRedded Then
        Select Case ShallWeSave
            Case 0 ''NO''
            Case 1 ''YES''
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    
    iOld = iImage(sst1.Tab)
    Select Case Index
        Case 0: iImage(sst1.Tab) = iImage(sst1.Tab) - 1
        Case 1: iImage(sst1.Tab) = iImage(sst1.Tab) + 1
    End Select
    
    Select Case sst1.Tab
        Case 4
            tGID = lstFiles(sst1.Tab).ItemData(iImage(sst1.Tab))
            tDesc = lstFiles(sst1.Tab).List(iImage(sst1.Tab))
            Call PrepareToOpen(iImage(sst1.Tab))
        Case Else
            tGID = lstFiles(sst1.Tab).ItemData((iListStart(sst1.Tab) + iImage(sst1.Tab)) - 1)
            tDesc = lstFiles(sst1.Tab).List((iListStart(sst1.Tab) + iImage(sst1.Tab)) - 1)
            Call LoadGraphic(sst1.Tab + 10, CStr(tGID), tDesc, _
                        CurrParNode(sst1.Tab), CurrParText(sst1.Tab))
    End Select
            
                    
    
'    strSelect = "SELECT GID, GDESC " & _
'                "FROM GFX_MASTER " & _
'                "WHERE GID = " & tGID
'    Set rst = Conn.Execute(strSelect)
'    If Not rst.EOF Then
'        Call LoadGraphic(sst1.Tab + 10, CStr(tGID), Trim(rst.Fields("GDESC")), _
'                    CurrParNode(sst1.Tab), CurrParText(sst1.Tab))
'        rst.Close: Set rst = Nothing
'    Else
'        rst.Close: Set rst = Nothing
'        MsgBox "File not found", vbCritical, "Sorry..."
'        iImage(sst1.Tab) = iOld
'    End If
    
End Sub

Private Sub imgPage_Click(Index As Integer)
    Select Case Index
        Case 0 ''BACK''
            xpdf1.currentPage = xpdf1.currentPage - 1
        Case 1 ''FORWARD''
            xpdf1.currentPage = xpdf1.currentPage + 1
    End Select
End Sub

Private Sub imgPageMode_Click(Index As Integer)
    iPageMode = Index
    xpdf1.continuousMode = CBool(iPageMode)
    Select Case Index
        Case 0 ''SINGLE PAGE MODE''
            imgPageMode(0).Picture = imlPageMode.ListImages(1).Picture
            imgPageMode(1).Picture = imlPageMode.ListImages(4).Picture
            If xpdf1.NumPages > 1 Then
                imgPageMode(1).Enabled = True
            Else
                imgPageMode(1).Enabled = False
            End If
        Case 1 ''CONTINUOUS PAGE MODE''
            imgPageMode(0).Picture = imlPageMode.ListImages(2).Picture
            imgPageMode(1).Picture = imlPageMode.ListImages(3).Picture
    End Select
    
End Sub

Private Sub imgPDF_Click(Index As Integer)
    SetZoomMode (Index)
    xpdf1.Visible = True
    
    cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
End Sub

Private Sub imgRed_Click(Index As Integer)
    iUndoMin = Index
    Call mnuGRedSketch_Click
End Sub

Private Sub imgRedMode_Click(Index As Integer)
    Select Case Index
        Case 0: Call mnuGRedSketch_Click
        Case 1: Call mnuGRedText_Click
    End Select
End Sub

Private Sub imgRedReload_Click()
    Call mnuGRedEnd_Click
End Sub

Private Sub imgSelect_Click()

End Sub

Private Sub imgSearchPDF_Click(Index As Integer)
    Select Case Index
        Case 0
            strSearchText = InputBox("What word or phrase would you like to search for?", "Search PDF...")
            If strSearchText <> "" Then
                
                xpdf1.Find strSearchText
                cmdFindNext.SetFocus
                cmdFindNext.Default = True
                imgSearchPDF(1).Visible = True
            Else
                cmdFindNext.Default = False
                imgSearchPDF(1).Visible = False
            End If
        Case 1
            xpdf1.findNext strSearchText
    End Select
End Sub

Private Sub imgUtility_Click(Index As Integer)
'''    Call ClearUndo(1)
    Select Case Index
        Case 0 ''SAVE''
            Call mnuGRedSave_Click
        Case 1 ''CLEAR''
            Call mnuGRedClear_Click
''''            bRedded = False
''''            If RedFile <> "" Then
''''                Call mnuGRedLoad_Click
''''            Else
''''                bRedMode = False
''''                If RedMode Then
''''                    Call mnuGRedSketch_Click
''''                ElseIf TextMode Then
''''                    Call mnuGRedText_Click
''''                End If
''''            End If
'            Call mnuGRedClear_Click
        Case 2 ''DELETE''
            Call mnuGRedDelete_Click
    End Select
End Sub

Private Sub imgZoom_Click(Index As Integer)
    Dim iZ As Integer, i As Integer
    
    iZ = CInt(xpdf1.zoomPercent)
    
    Select Case Index
        Case 0 ''ZOOM OUT''
            i = lstZoom.ListCount - 1
            Do Until lstZoom.ItemData(i) < iZ And i >= 0
                i = i - 1
            Loop
            xpdf1.Zoom = lstZoom.ItemData(i)
            cboZoom.Text = lstZoom.List(i)
            If xpdf1.zoomPercent <= 10 Then
                imgZoom(0).Picture = imlZoomMode.ListImages(13).Picture
                imgZoom(0).Enabled = False
            Else
                imgZoom(0).Picture = imlZoomMode.ListImages(14).Picture
                imgZoom(0).Enabled = True
            End If
            If xpdf1.zoomPercent < 500 Then
                imgPDF(2).Picture = imlZoomMode.ListImages(8).Picture
                imgPDF(2).Enabled = True
            End If
                
        Case 1 ''ZOOM IN''
            i = 0
            Do Until lstZoom.ItemData(i) > iZ And i < lstZoom.ListCount
                i = i + 1
            Loop
            xpdf1.Zoom = lstZoom.ItemData(i)
            cboZoom.Text = lstZoom.List(i)
            If xpdf1.zoomPercent >= 500 Then
                imgPDF(2).Picture = imlZoomMode.ListImages(7).Picture
                imgPDF(2).Enabled = False
                imgZoom(1).Picture = imlZoomMode.ListImages(15).Picture
                imgZoom(1).Enabled = False
            Else
                imgPDF(2).Picture = imlZoomMode.ListImages(8).Picture
                imgPDF(2).Enabled = True
                imgZoom(1).Picture = imlZoomMode.ListImages(16).Picture
                imgZoom(1).Enabled = True
            End If
    End Select
End Sub

Private Sub lblBackground_Click()
    Call ToggleBackground
End Sub

'''Private Sub imgDirs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case bDirsOpen
'''        Case True: Set imgDirs.Picture = imlDirs.ListImages(4).Picture
'''        Case False: Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'''    End Select
'''End Sub

'''Private Sub imgSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Set imgSearch.Picture = imlDirs.ListImages(6).Picture
'''End Sub
'''
'''Private Sub imgSupDoc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Set imgSupDoc.Picture = imlDirs.ListImages(8).Picture
'''End Sub

'''Private Sub chkReviewClients_Click()
'''    Dim i As Integer
'''
'''    If chkReviewClients.value = 1 Then
'''        Call PopReviewClients(UserID)
'''    Else
'''        Call PopClientsWithGraphics(cboCUNO(3), tvwGraphics(3))
'''    End If
'''    Call ClearThumbnails3(0)
'''End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub


'''Private Sub cmdAssign_Click()
'''    cmdAssign.Visible = False
''''''''    On Error Resume Next
''''''''    If cboCUNO(1).Text <> cboCUNO(0).Text Then
''''''''        cboCUNO(1).Text = cboCUNO(0).Text
''''''''    End If
'''    cmdAssign.Refresh
'''    bAddMode = True
'''    With frmAssign
'''        .PassBCC = tBCC
'''        .PassFBCN = tFBCN
'''        .PassSHYR = tSHYR
'''        .PassSHCD = tSHCD
'''        .PassSHNM = tSHNM
'''        On Error Resume Next
'''        .Show , Me
'''        If Err Then
'''            MsgBox "At this time, you cannot access the Assignment Interface when coming in from " & _
'''                        "the Floorplan Viewer.  To access the Assignment Interface, close both the " & _
'''                        "Graphics Viewer and the Floorplan Viewer, and then come directly into " & _
'''                        "the Graphics Interface from the opening screen.", vbExclamation, "Sorry..."
'''            Err.Clear
'''        End If
'''    End With
'''End Sub

Private Sub imgDirs_Click()
    Dim i As Integer
    If picTabs.Visible = False Then ''sst1.Visible = False Then
'        picTools.Visible = False
        sst1.Visible = True
        picTabs.Visible = True ''sst1.Visible = True
        imgDirs.ToolTipText = "Click to toggle to Image File"
        bDirsOpen = True
        Set imgDirs.Picture = imlDirs.ListImages(1).Picture
        imgComm.Visible = False
        lblWelcome.Visible = False
        Me.lblGraphic.Visible = False
'        If sst1.Tab = 4 Then
''''            Me.AutoRedraw = False
''''            picTabs.AutoRedraw = False
''            picTabs.Refresh
'            For i = 0 To imx4.Count - 1
'                imx4(i).Refresh
'
'            Next i
'        End If
    Else
        picTabs.Visible = False ''sst1.Visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
        If bPerm(26) And bPicLoaded Then imgComm.Visible = True
        lblWelcome.Visible = True
        Me.lblGraphic.Visible = True
    End If
End Sub


Private Sub imgImporter_Click()
    Dim RetVal
    Dim sTFile As String
'''    sTFile = "M:\Temp\Exporter.gpj"
    sTFile = sGIPath & ":\Program Files\GraphicExporter\Exporter.gpj"
'''    sTFile = "C:\Temp\Exporter.gpj"
    Open sTFile For Output As #1
    Write #1, Shortname, LogAddress
    Close #1
    
'''    RetVal = Shell("C:\Program Files\GraphicExporter\GraphicExporter.exe", 1)
'''    RetVal = Shell("M:\Program Files\GraphicExporter\GraphicExporter.exe", 1)
    RetVal = Shell(sGIPath & ":\Program Files\GraphicExporter\GraphicExporter.exe", 1)
'''''    RetVal = Shell("D:\Data\VB Projects\GPJAnnotator\GraphicExporter\GraphicExporter.exe", 1)
    
End Sub

'''Private Sub cmdGfxApprove_Click()
'''    Dim i As Integer, Index As Integer, iErr As Integer
'''    Dim iNewStatus(0 To 3) As Integer
'''    Dim sGStatus(0 To 30) As String
'''    Dim strUpdate As String, sComm As String
'''
'''    Screen.MousePointer = 11
'''
'''    '///// FILE STATUS VARIABLES \\\\\
'''    sGStatus(0) = "DE-ACTIVATED"
'''    sGStatus(10) = "INTERNAL"
'''    sGStatus(20) = "CLIENT DRAFT"
'''    sGStatus(30) = "APPROVED"
'''
'''    iNewStatus(0) = 5
'''    iNewStatus(1) = 15
'''    iNewStatus(2) = 25
'''
'''    Select Case iCurrStatus
'''        Case 10: iNewStatus(3) = 2
'''        Case 20: iNewStatus(3) = 3
'''        Case 30: iNewStatus(3) = 4
'''    End Select
'''
'''    For i = 0 To 3
'''        If optGfxApprove(i).value = True Then
'''            Index = i
'''            Exit For
'''        End If
'''    Next i
'''
'''    Conn.BeginTrans
'''    On Error GoTo ErrorTrap
'''
'''    strUpdate = "UPDATE " & GFXMas & " " & _
'''                "SET GSTATUS = " & iNewStatus(Index) & ", " & _
'''                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
'''                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
'''                "WHERE GID = " & lRefID
'''    Conn.Execute (strUpdate)
'''
'''    sComm = txtApprove.Text
'''    iErr = InsertComment(lRefID, Index, sComm)
'''    If iErr > 0 Then GoTo ErrorTrap
'''
'''    Conn.CommitTrans
''''''    picGfxApprove.Visible = False
'''    cmdGfxApproveHide.Visible = False
'''
'''    ''CHANGE LBLGRAPHIC''
'''    i = InStr(1, lblGraphic.Caption, " ")
'''    If iNewStatus(Index) >= 5 Then
'''        lblGraphic.Caption = sGStatus(iNewStatus(Index) + 5) & Mid(lblGraphic.Caption, i)
'''        lblStatus.Caption = "STATUS:  " & sGStatus(iNewStatus(Index) + 5) & " (Last Status Update " & _
'''                    format(Now, "MMMM D, YYYY") & ")"
'''    Else
'''        lblGraphic.Caption = "DE-ACTIVATED" & Mid(lblGraphic.Caption, i)
'''        lblStatus.Caption = "STATUS:  DE-ACTIVATED (" & format(Now, "MMMM D, YYYY") & ")"
'''    End If
'''    lblStatus.Visible = True
'''
'''    cboCUNO(4).Text = cboCUNO(4).Text
'''    cmdStatusEdit_Notify.Enabled = CheckForNotify
'''
'''    Screen.MousePointer = 0
'''
'''Exit Sub
'''ErrorTrap:
'''    Conn.RollbackTrans
'''    Screen.MousePointer = 0
'''    MsgBox "Error Encountered during Status Change." & vbNewLine & vbNewLine & _
'''                "Error:  " & Err.Description, vbCritical, "Status Change Aborted..."
'''    Err.Clear
'''End Sub

''''''''Private Sub lblfullsize_Click()
''''''''    If picJPG.Visible Then
''''''''        frmHTMLViewer.PassFile = CurrFile
''''''''        frmHTMLViewer.PassFrom = Me.Name
''''''''        frmHTMLViewer.PassHDR = lblWelcome.Caption
''''''''        frmHTMLViewer.PassGID = lGID
''''''''        frmHTMLViewer.Show 1, Me
''''''''    Else
''''''''        SetZoomMode (1)
''''''''        cboZoom.Text = CInt(Xpdf1.zoomPercent) & "%"
''''''''    End If
''''''''End Sub

'''Private Sub cmdGfxApproveHide_Click()
'''    picGfxApprove.Visible = True
'''    cmdGfxApproveHide.Visible = False
'''End Sub

Private Sub cmdHelp_Click()
    If picHelp.Visible = False Then
        web1.Navigate2 App.Path & "/Graphic Approval Process.htm"
        picHelp.Visible = True
    Else
        picHelp.Visible = False
        web1.Navigate2 ""
    End If
End Sub

Private Sub cmdHelpClose_Click()
    picHelp.Visible = False
End Sub

Private Sub lblGraphic_Change()
    lblRedline.Left = lblGraphic.Left + lblGraphic.Width + 600
End Sub



'''''Private Sub lblImporter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''''    lblImporter.ForeColor = vbWhite
'''''End Sub

Private Sub lblkeyedit_Click()
    frmKeywordEdit.PassGID = lGID
    frmKeywordEdit.PassFrom = "GH"
    frmKeywordEdit.Show 1, Me

End Sub

Private Sub lblMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Me.PopupMenu mnuRightClick, 0, cmdMenu.Left, cmdMenu.Top + cmdMenu.Height
    Me.PopupMenu mnuRightClick, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
End Sub

Private Sub lblPage_Click()
    Dim iTempPage As Integer
    
    iTempPage = iPDFPage
    frmGetPage.PassMax = xpdf1.NumPages
    frmGetPage.PassVal = iPDFPage
    frmGetPage.Show 1, Me
    
    If iTempPage <> iPDFPage Then
        xpdf1.currentPage = iPDFPage
    End If
End Sub

Private Sub lblRed_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Source.Move lblRed(Index).Left + X - x0, lblRed(Index).Top + Y - y0
End Sub

Private Sub lblRed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        iRed = Index
        x0 = X
        y0 = Y
        lblRed(iRed).WordWrap = False
        lblRed(iRed).Drag 1
    ElseIf Button = vbRightButton Then
        iRed = Index
        Me.PopupMenu mnuText
    End If
End Sub

Private Sub lblRedline_Click()
    If bRedded Then
        Select Case ShallWeSave
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    Call mnuGRedLoad_Click
End Sub

'''Private Sub lblRedMode_Click(Index As Integer)
'''    Select Case Index
'''        Case 0: Call mnuGRedSketch_Click
'''        Case 1: Call mnuGRedText_Click
'''    End Select
'''End Sub

''''''''Private Sub lblResize_Click()
''''''''    Select Case lblResize.Caption
''''''''        Case "Resize"
'''''''''''            lblresize.Caption = "Maximize Graphic"
''''''''            mnuResizeGraphic_Click
''''''''        Case "Maximize"
'''''''''''            lblresize.Caption = "Resize Graphic"
''''''''            mnuMaxGraphic_Click
''''''''    End Select
''''''''End Sub

'''Private Sub cmdStatusEdit_All_Click()
'''    Dim i As Integer
'''    Dim dTop As Double, dLeft As Double
'''
'''    mnuName(0).Caption = "Execution will Edit the Status of all ( " & iGFXCount & " ) Files"
'''    mnuName(1).Caption = "in the '" & tvwGraphics(3).SelectedItem.Parent.Text & " - " & _
'''                tvwGraphics(3).SelectedItem.Text & "' Folder"
'''
'''    Select Case Screen.Width / Screen.TwipsPerPixelX
'''        Case Is >= 1024
'''            dLeft = sst1.Left + picReview.Left + _
'''                        ((picReview.Width - picReview.ScaleWidth) / 2) + cmdStatusEdit_All.Left
'''        Case Else
'''            dLeft = sst1.Left + picReview.Left + ((picReview.Width - picReview.ScaleWidth) / 2) + _
'''                        cmdStatusEdit_All.Left + cmdStatusEdit_All.Width
'''    End Select
'''    dTop = sst1.Top + picReview.Top + ((picReview.Height - picReview.ScaleHeight) / 2) + _
'''                cmdStatusEdit_All.Height
'''
'''    mnuName(0).Tag = "ALL"
'''
'''    Me.PopupMenu mnuResetStatus, 0, dLeft, dTop
'''SkipIt:
'''End Sub

Private Sub lblRN_Click()
    bHideRN = True
    Call RedNoteVis(False)
End Sub

Private Sub lblSearch_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sDNode As String
    Dim i As Integer
    
    lOpenInViewer = 0
    If picTabs.Visible = True And sst1.Tab = 4 And bApprover Then
        frmSearch.PassFrom = "GA"
        frmSearch.PassBCC = fBCC(4)
    Else
        frmSearch.PassFrom = "GH"
    End If
    frmSearch.Show 1
    If lOpenInViewer > 0 Then
        ''LOAD THE IMAGE''
'''        MsgBox "Loading the Image..."
        sst1.Tab = 3
        If cboCUNO(3).Text <> sPassedFBCN Then
            cboCUNO(3).Text = sPassedFBCN
            fFBCN(3) = sPassedFBCN
            fBCC(3) = Right("000000000000" & sPassedBCC, 12)
        End If
        
        strSelect = "SELECT TO_CHAR(ADDDTTM, 'MM') AS M2, " & _
                    "TO_CHAR(ADDDTTM, 'YYYY') AS Y4, " & _
                    "NVL(FLR_ID, 0) AS FLR_ID " & _
                    "FROM ANNOTATOR.GFX_MASTER " & _
                    "WHERE GID > 0 " & _
                    "AND GID = " & sPassedGNode
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            If rst.Fields("FLR_ID") > 0 Then
                sDNode = "f" & rst.Fields("FLR_ID")
            Else
                sDNode = "d" & Trim(rst.Fields("M2")) & Trim(rst.Fields("Y4"))
            End If
        End If
        rst.Close: Set rst = Nothing
        
        On Error Resume Next
        If tvwGraphics(3).SelectedItem.Key <> sDNode Then
            If Err Then Err.Clear
            tvwGraphics(3).Nodes(sDNode).Selected = True
            Call tvwGraphics_NodeClick(3, tvwGraphics(3).Nodes(sDNode))
        End If
        For i = imx3.LBound To imx3.UBound
            If imx3(i).Tag = CStr(sPassedGNode) Then
                Call imx3_Click(i)
                Exit For
            End If
        Next i
    End If
    
'''    Set imgSearch.Picture = imlDirs.ListImages(5).Picture
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblClose.ForeColor = vbWhite
End Sub


'''Private Sub cmdStatusEdit_Notify_Click()
'''    frmNotify.Show 1
'''    cmdStatusEdit_Notify.Enabled = CheckForNotify
'''End Sub



Private Sub cmdStatusEdit_View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, iCnt As Integer
    Dim dTop As Double, dLeft As Double
     
    
    iCnt = flxApprove.Rows - 1
    
    Select Case iCnt
        Case 0: GoTo SkipIt
        Case 1
            frmGfxApprove.PassHDR = "Editing the one File in Current View"
'''            mnuName(0).Caption = "Editing the one File in Current View"
'''            mnuName(1).Caption = ""
        Case Else
            frmGfxApprove.PassHDR = "Execution will Edit the ( " & iCnt & " ) " & _
                    "Files in the View"
'''            mnuName(0).Caption = "Execution will Edit the ( " & iCnt & " ) " & _
'''                    "Files in the Current View"
'''            mnuName(1).Caption = ""
    End Select
    
'''    Select Case Screen.Width / Screen.TwipsPerPixelX
'''        Case Is >= 1024
'''            frmGfxApprove.PassX = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + _
'''                        picTabs.Left + picReview.Left + _
'''                        ((picReview.Width - picReview.ScaleWidth) / 2) + cmdStatusEdit_View.Left
''''''            dLeft = sst1.Left + picReview.Left + _
''''''                        ((picReview.Width - picReview.ScaleWidth) / 2) + cmdStatusEdit_View.Left
'''        Case Else
            frmGfxApprove.PassX = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + _
                        picTabs.Left + picReview.Left + ((picReview.Width - picReview.ScaleWidth) / 2) + _
                        cmdStatusEdit_View.Left + cmdStatusEdit_View.Width - 5400 '''frmGfxApprove.Width
'''            dLeft = sst1.Left + picReview.Left + ((picReview.Width - picReview.ScaleWidth) / 2) + _
'''                        cmdStatusEdit_View.Left + cmdStatusEdit_View.Width
'''    End Select
    frmGfxApprove.PassY = Me.Top + (Me.Height - Me.ScaleHeight) - ((Me.Width - Me.ScaleWidth) / 2) + _
                picTabs.Top + picReview.Top + picReview.Height
                ''' ((picReview.Height - picReview.ScaleHeight) / 2) + _
                cmdStatusEdit_View.Height
'''    dTop = sst1.Top + picReview.Top + ((picReview.Height - picReview.ScaleHeight) / 2) + _
'''                cmdStatusEdit_View.Height
    
'''    mnuName(0).Tag = "VIEW"
'''    mnuOpen.Visible = False
'''    mnuStatusReset(0).Visible = True
'''    mnuStatusReset(1).Visible = True
'''    mnuStatusReset(2).Visible = True
    Select Case UCase(lblStat(0).Caption)
        Case "INTERNAL": frmGfxApprove.PassVal = 0
        Case "CLIENT DRAFT": frmGfxApprove.PassVal = 1
        Case "APPROVED": frmGfxApprove.PassVal = 2
        Case "RETURNED": frmGfxApprove.PassVal = 4
'''        Case "INTERNAL": mnuStatusReset(0).Visible = False
'''        Case "CLIENT DRAFT": mnuStatusReset(1).Visible = False
'''        Case "APPROVED": mnuStatusReset(2).Visible = False
    End Select
    frmGfxApprove.PassBCC = CLng(fBCC(4))
    frmGfxApprove.PassFBCN = fFBCN(4)
    
    frmGfxApprove.PassType = "VIEW"
    frmGfxApprove.Show 1, Me
'''    Me.PopupMenu mnuResetStatus, 0, dLeft, dTop
SkipIt:
End Sub

Private Sub cmdStatusEdit_View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mnuOpen.Visible = True
End Sub

Private Sub imgSupdoc_Click()
    Dim sSupDocPath As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If imgSupDoc.Tag <> "" Then
        sSupDocPath = "\\DETMSFS01\GPJAnnotator\Graphics\SupDocs\"
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
        
'''        Set imgSupDoc.Picture = imlDirs.ListImages(7).Picture
    End If

End Sub

'''Private Sub cmdXData_Click()
'''    If picExpand(0).Visible Then picExpand_Click (0) Else picExpand_Click (1)
'''End Sub


'''Private Sub Command2_Click()
'''    frmHTMLViewer.PassFile = App.Path & "\TestDoc.htm"
'''    frmHTMLViewer.Show 1, Me
'''End Sub

'''Private Sub Command2_Click()
'''    Dim i As Integer, iCol As Integer
'''    i = imx4.Count
'''
'''    Load shp4(i)
'''    shp4(i).Top = shp4(i).Top + (i * 1080) '''1500)
'''    If i Mod 2 = 1 Then
'''        shp4(i).BackColor = vb3DLight
'''        shp4(i).BorderColor = vb3DLight
'''    End If
'''    shp4(i).Visible = True
'''
''''''    Load shp5(i)
''''''    shp5(i).Top = shp5(i).Top + (i * 1080) '''1500)
''''''    If i Mod 2 = 1 Then
''''''        shp5(i).BackColor = vb3DLight
''''''        shp5(i).BorderColor = vb3DLight
''''''    End If
''''''    shp5(i).Visible = True
'''
'''    Load imx4(i)
'''    imx4(i).Top = imx4(i).Top + (i * 1080) '''1500)
'''    If i Mod 2 = 1 Then imx4(i).BackColor = vb3DLight
'''    imx4(i).Visible = True
'''
'''    Load imgStat(i)
'''    imgStat(i).Top = imgStat(i).Top + (i * 1080) '''1500)
'''    imgStat(i).Visible = True
'''    imgStat(i).ZOrder
'''
'''    Load lblStat(i)
'''    lblStat(i).Top = lblStat(i).Top + (i * 1080) '''1500)
'''    lblStat(i).Visible = True
'''    lblStat(i).ZOrder
'''
'''    If i Mod 2 = 1 Then
'''        With flxApprove
'''            .Row = i + 1
'''            For iCol = 1 To .Cols - 1
'''                .Col = iCol: .CellBackColor = vb3DLight
'''            Next iCol
'''        End With
'''    End If
'''End Sub

Private Sub flxApprove_Click()
'''    Debug.Print "Clicked Row " & flxApprove.RowSel & ", Col " & flxApprove.ColSel
'''    sOrder = "ORDER BY GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
    
    If bDataSort Then
        Exit Sub
    ElseIf bRighted Then
        Exit Sub
    ElseIf flxApprove.Rows = 1 Then
        Exit Sub
    Else
        Select Case flxApprove.ColSel
            Case 4
'                MsgBox "View Comments for Row " & flxApprove.RowSel
                With frmComments
                    .PassREFID = flxApprove.TextMatrix(flxApprove.RowSel, 0)
                    .PassTable = sTable
                    .PassIType = 1
                    .PassBCC = fBCC(4)
                    .PassFBCN = fFBCN(4)
                    .PassSHCD = fSHCD(4)
                    .PassMessPath = UCase(Trim(cboCUNO(4).Text)) & " " & _
                                "Client Graphics:  " & _
                                flxApprove.TextMatrix(flxApprove.RowSel, 3)
                    .PassMessSub = "Testing 2"
                    .PassForm = "frmGraphics"
                    .PassGPath = imx4(flxApprove.RowSel - 1).FileName
                    .Show 1
                End With
            Case 6
                If Not bGfxMandRecip Then Exit Sub
                With frmApprover
                    .PassBCC = fBCC(4)
                    .PassHDR = flxApprove.TextMatrix(flxApprove.RowSel, 3)
                    .PassName = flxApprove.TextMatrix(flxApprove.RowSel, 6)
                    .PassGID = flxApprove.TextMatrix(flxApprove.RowSel, 0)
                    
                    .PassX = rApproverX
                    .PassY = rApproverY
                    .Show 1
                End With
        End Select
    End If
End Sub

Private Sub flxApprove_DblClick()
    Dim i As Integer, iRow As Integer
    
    Debug.Print "Dbl-Clicked Row " & flxApprove.RowSel & ", Col " & flxApprove.ColSel
    If bDataSort Then
        Select Case flxApprove.ColSel
            Case 1
                sOrder = "ORDER BY GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
                lblMess.Caption = "...Re-Sorting to Default (by Status)..."
            Case 2
                sOrder = "ORDER BY GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
                lblMess.Caption = "...Re-Sorting by Graphics Status..."
            Case 3
                sOrder = "ORDER BY UPPER(GM.GDESC), GM.GSTATUS, GM.GTYPE"
                lblMess.Caption = "...Re-Sorting by File Name..."
            Case 5
                sOrder = "ORDER BY GM.GTYPE, GM.GSTATUS, UPPER(GM.GDESC)"
                lblMess.Caption = "...Re-Sorting by Graphics Type..."
            Case 6
                sOrder = "ORDER BY APPROVER, GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
                lblMess.Caption = "...Re-Sorting by Approver..."
            Case 7
                sOrder = "ORDER BY GM.ADDDTTM"
                lblMess.Caption = "...Re-Sorting by Post Date..."
            Case 8
                sOrder = "ORDER BY GM.ADDUSER, GM.GSTATUS, GM.GTYPE, UPPER(GM.GDESC)"
                lblMess.Caption = "...Re-Sorting by Poster..."
        End Select
        picMess.Visible = True: picMess.Refresh
        flxApprove.Visible = False: picOuter(4).Visible = False
        Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, fSHYR(4), fSHCD(4), lFID)
        picMess.Visible = False
        flxApprove.Visible = True: picOuter(4).Visible = True
'''    Else
'''        Select Case flxApprove.ColSel
'''            Case 4
'''                MsgBox "View Comments for Row " & flxApprove.RowSel
'''        End Select
    End If
End Sub


Private Sub flxApprove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        bRighted = True
        Exit Sub
    Else
        bRighted = False
    End If
    
    If Y < flxApprove.RowHeight(0) Then
        bDataSort = True
    Else
        bDataSort = False
        
        rApproverX = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + _
                    picTabs.Left + flxApprove.Left + X
        rApproverY = Me.Top + (Me.Height - Me.ScaleHeight) + _
                    picTabs.Top + flxApprove.Top + Y
    End If
End Sub

Private Sub flxApprove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= rCommX1 And X <= rCommX2 And Y > flxApprove.RowHeight(0) Then
        flxApprove.MousePointer = 99
    ElseIf bGfxMandRecip And X >= rApproverX1 And X <= rApproverX2 And Y > flxApprove.RowHeight(0) Then
        flxApprove.MousePointer = 99
    Else
        flxApprove.MousePointer = 0
    End If
End Sub

Private Sub flxApprove_Scroll()
    picInner(4).Top = ((flxApprove.TopRow - 1) * 1080) * -1
'    picInner(5).Top = ((flxApprove.TopRow - 1) * 1080) * -1
End Sub

Private Sub Form_Load()
    Dim i As Integer, iCol As Integer, iRow As Integer, iTabs As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bFound As Boolean
    Dim iDebug As Integer
    
    If bDebug Then MsgBox "Opening Graphics Handler"
    iDebug = 0
    
    Screen.MousePointer = 11
    
    bAcro = False
    
'''    Set pdfGraphic = Controls.Add("AcroPDF.PDF.1", "pdfGraphic")
'''    pdfGraphic.ZOrder 1
'''    pdfGraphic.Visible = False
    
    bDirsOpen = True
    picTabs.Left = 0
    picTabs.Top = 600
    
    ''SET VARIABLES''
    For i = 0 To 4
        fBCC(i) = ""
        fFBCN(i) = ""
        fSHYR(i) = 0
        fSHCD(i) = 0
        fSHNM(i) = ""
    Next i
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    sOrder = "ORDER BY UPPER(GM.GDESC), GM.GSTATUS, GM.GTYPE"
    sNoShows = "No Graphics awaiting review have been assigned to a Show"
    
    If bGPJ Then defSIN = "10, 20, 27, 30" Else defSIN = "20, 27, 30"
    If bGPJ Then defEditSIN = "10, 20, 27" Else defEditSIN = "20, 27"
    
    ''CHECK IF COMING FROM FLOORPLAN''
    If tBCC <> "" Then bFromFloorplan = True
    
    ''SET COLOR CONSTANTS''
''    lBColor(0) = &HC0C0C0
''    lBColor(5) = &H8080FF
''    lBColor(10) = &H8080FF
''    lBColor(15) = &H80FFFF
''    lBColor(20) = &H80FFFF
''    lBColor(25) = &H80FF80
''    lBColor(27) = &H80C0FF
''    lBColor(30) = &H80FF80
    lBColor(0) = &HC0C0C0
    lBColor(5) = vbRed
    lBColor(10) = vbRed
    lBColor(15) = vbYellow
    lBColor(20) = vbYellow
    lBColor(25) = vbGreen
    lBColor(27) = RGB(64, 128, 255) '' RGB(255, 127, 0)
    lBColor(30) = vbGreen
    
    '///// GRAPHIC TYPES \\\\\
    GfxType(1) = "Digital Photos"
    GfxType(2) = "Graphic Files"
    GfxType(3) = "Graphic Layouts"
    GfxType(4) = "Presentation Files"
    
    ''///// SET SHOW-SEASON SORT \\\\\''
    iSSSort = 1
    
    ''///// SET DEFAULT APPROVER VIEW \\\\\''
    iApproverView = 0
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
    ''///// SET TABTYPES & DEFAULT INTYPES \\\\\''
    For i = 0 To 4
        iTabType(i) = 0
        sInType(i) = "1, 2, 3, 4"
    Next i
        
    If bGPJ Then
        sIN = "10, 20, 27"
    Else
        sIN = "20, 27"
        imgStatus(1).Visible = False
        imgStatus(2).Left = imgStatus(2).Left - 720
        imgStatus(3).Left = imgStatus(3).Left - 720
        imgStatus(4).Left = imgStatus(4).Left - 720
        cmdStatusEdit_View.Left = cmdStatusEdit_View.Left - 720
        picReview.Width = picReview.Width - 720
    End If
    
    maxX = Me.ScaleWidth - 240
    maxY = Me.ScaleHeight - 795 - picTools.Height - 60 ''120
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
    iTabs = 4
'''    sst1.TabVisible(0) = False
'''    iTabs = iTabs - 1

    If bPerm(51) = False Then
        sst1.TabVisible(0) = False
'        iTabs = iTabs - 1
    Else
        sst1.TabVisible(0) = True
    End If
    If bPerm(52) = False Then
        sst1.TabVisible(1) = False
        iTabs = iTabs - 1
    End If
    If bPerm(53) = False Then
        sst1.TabVisible(2) = False
        iTabs = iTabs - 1
    End If
    If bGFXReviewer And iView = 1 Then
        sst1.TabVisible(4) = True
        iTabs = iTabs + 1
    End If
    
    sst1.TabVisible(0) = False
    sst1.TabsPerRow = iTabs - 1
    
    For i = imx0.LBound To imx0.UBound
        imx0(i).Width = lIconX
        imx0(i).Height = lIconY
    Next i
    For i = imx1.LBound To imx1.UBound
        imx1(i).Width = lIconX
        imx1(i).Height = lIconY
    Next i
    For i = imx2.LBound To imx2.UBound
        imx2(i).Width = lIconX
        imx2(i).Height = lIconY
    Next i
    For i = imx3.LBound To imx3.UBound
        imx3(i).Width = lIconX
        imx3(i).Height = lIconY
    Next i
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
'    iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
'    iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
'
'    imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
'    spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
'    imageX = CLng((picInner(0).ScaleWidth - 240) / iCols) '' (imageY / 3) * 4
'    spaceX = imageX ''CLng(picInner(0).ScaleWidth / iCols) '' imageX + 720 ''240
    
'''    For i = 0 To (iCols * iRows) - 1 ''19
'''        If i >= imx0.Count Then
'''            Load imx0(i)
'''            Set imx0(i).Container = picInner(0)
'''        End If
'''        iCol = Int(i / iRows): iRow = i Mod iRows
'''        imx0(i).Width = lIconX
'''        imx0(i).Height = lIconY
''''''        imx0(i).Left = 480 + (iCol * spaceX)
'''        imx0(i).Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
'''        imx0(i).Top = 120 + (iRow * spaceY)
'''
'''        If i >= lbl0.Count Then Load lbl0(i)
'''
'''        If i >= chkMulti.Count Then Load chkMulti(i)
'''    Next i
    
'    imageY = (picInner(1).Height - hsc1(1).Height - 240 - 900) / iRows
'    spaceY = imageY + 270 '''300
'    imageX = (imageY / 3) * 4
'    spaceX = imageX + 720 ''240
    
'    For i = 0 To 19
'        If i >= imx0.Count Then Load imx0(i)
'        If i >= imx1.Count Then Load imx1(i)
'        If i >= imx2.Count Then Load imx2(i)
'        If i >= imx3.Count Then Load imx3(i)
'        iCol = Int(i / iRows): iRow = i Mod iRows
'        imx1(i).Left = 480 + (iCol * spaceX)
'        imx1(i).Top = 120 + (iRow * spaceY)
'
'        If i >= lbl0.Count Then Load lbl0(i)
'        If i >= lbl1.Count Then Load lbl1(i)
'        If i >= lbl2.Count Then Load lbl2(i)
'        If i >= lbl3.Count Then Load lbl3(i)
'
'        If i >= shp1.Count Then Load shp1(i)
'        If i >= shp2.Count Then Load shp2(i)
'        If i >= shp3.Count Then Load shp3(i)
'
'    Next i

'''''    picInner(0).Width = ((iCol + 1) * spaceX) + 240
'''''    picInner(1).Width = ((iCol + 1) * spaceX) + 240
'''''    picInner(2).Width = ((iCol + 1) * spaceX) + 240
'''''    picInner(3).Width = ((iCol + 1) * spaceX) + 240
    
    
    sGPath = "\\DETMSFS01\GPJAnnotator\Graphics\"
    sVPath = sGPath & "Versions\"

    sMessPath = ""
    rAsp = 0
    bAddMode = False
    bPicLoaded = False
    lNewLockId = 0
    
    Me.WindowState = AppWindowState
    
    picJPG.Top = 1200 '' 1380 ''675
    picJPG.Left = 1320 '' 1260 ''120
'''    picGfxApprove.Top = picJPG.Top + 120
    dTop = picJPG.Top
    dLeft = picJPG.Left
    
'''    If bAcro Then
        picPDF.Top = dTop
        picPDF.Left = dLeft
        
        xpdf1.Top = 0 ''15 '' picPDF.Top
        xpdf1.Left = 0 '' 15 ''picPDF.Left
        xpdf1.matteColor = Me.BackColor
        
        picRed.Top = xpdf1.Top
        picRed.Left = xpdf1.Left
    
        
        
'''        For i = 0 To 15
'''            picColor(i).BackColor = QBColor(i)
'''            picColor(i).Left = 30 + (180 * i)
'''            If picColor(i).BackColor = vbRed Then
'''                shpHL.Left = picColor(i).Left - 30
'''            End If
'''        Next i
        iAnnoColor = 12
        lAnnoColor = QBColor(iAnnoColor)
        
        
        
        '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
        PopZooms
    
        xpdf1.matteColor = Me.BackColor
        
    
'''    End If

'''    pdfGraphic.setShowToolbar (False)
    
'    pdfGraphic.setShowToolbar
'''    lblByGeorge(0).ForeColor = lGeo_Back '' RGB(56, 88, 14) ''  RGB(30, 30, 21)
'''    lblByGeorge(1).ForeColor = lGeo_Fore '' RGB(111, 175, 28) '' RGB(100, 100, 68)
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
    '///// FIRST, POP SHYR COMBOS \\\\\
    CurrSHYR = CInt(Format(Now, "YYYY"))
    For i = -2 To 2
        cboSHYR(0).AddItem CurrSHYR + i
        cboSHYR(1).AddItem CurrSHYR + i
    Next i
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    '///// NEXT, POP INVENTORY CLIENT LIST \\\\\
    Call PopClientsWithInventory(cboCUNO(2))
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    '///// NEXT, POP CLIENTS WITH DB GRAPHICS \\\\\
    Call PopClientsWithGraphics(cboCUNO(3), tvwGraphics(3))
    
    '///// NEXT, PASS IN VALUES (IF THEY EXIST) \\\\\
    Err = 0
    On Error Resume Next
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    ''ADD TEST FOR PASSING IN VARS''
    If tSHYR = 0 And SHYR <> 0 Then
        tSHYR = SHYR
        For i = 0 To 4
            fSHYR(i) = SHYR
        Next i
    End If
    
'    If tSHYR <> 0 Then
'        cboSHYR(0).Text = tSHYR
'        If Err Then
'            cboSHYR(0).Text = CurrSHYR
'            Err.Clear
'        End If
'        cboSHYR(1).Text = tSHYR
'        If Err Then
'            cboSHYR(1).Text = CurrSHYR
'            Err.Clear
'        End If
'    Else
'        cboSHYR(0).Text = CurrSHYR
'        cboSHYR(1).Text = CurrSHYR
'    End If
    
    '///// NOW PASS IN CUNOS IF THEY EXIST \\\\\
    
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
    ''ADD TEST FOR PASSING IN VARS''
    If tBCC = "" And BCC <> "" Then
        tBCC = BCC
        tFBCN = FBCN
        For i = 0 To 4
            fBCC(i) = BCC
            fFBCN(i) = FBCN
        Next i
    Else
        If Not bFromFloorplan And bPassIn = False And defCUNO > 0 Then
            ''CHECK FOR STORED PROFILE''
            ''HARD-CODE FOR TIME BEING''
            tBCC = defCUNO: lBCC_Def = tBCC
            tFBCN = defFBCN: sFBCN_Def = tFBCN
            For i = 0 To 4
                fBCC(i) = tBCC
                fFBCN(i) = tFBCN
            Next i
'            tBCC = 1161 ''1190
'            tFBCN = "SAAB CARS USA, INC." '' "Toyota Motor Sales USA, Inc."
        End If
    End If
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
'    If tBCC <> "" And tFBCN <> "" Then
'        cboCUNO(0).Text = tFBCN
'        cboCUNO(1).Text = tFBCN
'        cboCUNO(2).Text = tFBCN
'        cboCUNO(3).Text = tFBCN
'    End If
    
    '///// NOW PASS IN SHCD IF COMING FROM FLOORPLAN \\\\\
    
    ''ADD TEST FOR PASSING IN VARS''
    If tSHCD = 0 And SHCD <> 0 Then
        tSHCD = SHCD
        tSHNM = SHNM
        For i = 0 To 4
            fSHCD(i) = tSHCD
            fSHNM(i) = tSHNM
        Next i
    End If

'    If tSHCD <> 0 And tSHNM <> "" Then
'        cboSHCD.Text = tSHNM
'    End If
    
'''    picMenu.Left = picJPG.Left + 120
'''    picMenu.Top = picJPG.Top + 120
'''    cmdMenu.Left = picJPG.Left + 120
'''    cmdMenu.Top = picJPG.Top + 120
    bMenuButton = True
    
    For i = 0 To 3
        picOuter(i).Left = 4500
    Next i
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
''    sst1.Tab = 3
    
    sTabDesc = sst1.TabCaption(sst1.Tab)
    bGfxOpen = True
    
    '///// TIME TO SET VIEW BASED ON PERMS \\\\\
    If bPerm(31) Then imgImporter.Visible = True Else imgImporter.Visible = False
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    Call ResetControls(iView)
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    Select Case iView
        Case 0
            mnuSelByName.Checked = True
            mnuSelByImage(2).Checked = False
            mnuSelByImage(3).Checked = False
        Case 1
            mnuSelByName.Checked = False
            mnuSelByImage(2).Checked = True
            mnuSelByImage(3).Checked = False
    End Select
    If iView = 0 Then
'''''        If bPerm(24) Or bPerm(25) Then
'''''            cmdAssign.Visible = True
'''''            cboCUNO(0).Width = 3375
'''''            cboSHCD.Width = 4335
'''''        Else
'''''            cmdAssign.Visible = False
            cboCUNO(0).Width = 4755
            cboSHCD.Width = 5715
'''''        End If
    ElseIf iView = 1 Then
'''''        If bPerm(24) Or bPerm(25) Then cmdAssign.Visible = True Else cmdAssign.Visible = False
        cboCUNO(0).Width = 3375
        cboSHCD.Width = 4335
    End If
    If Not bPerm(39) Then
        mnuRedlining.Visible = False
        mnuDash02.Visible = False
        imgRed(0).Visible = False
        imgRed(1).Visible = False
    End If
    If bPerm(45) Then
        mnuDownload.Visible = True
    Else
        mnuDownload.Visible = False
    End If
    
    If Not bPerm(59) Then
        imgKeyEdit.Visible = False
        picMenu2.Height = imgKeyEdit.Top '' imgFullSize.Top + imgFullSize.Height
    End If
        
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(False)
    If bENABLE_PRINTERS Then
        mnuGPrint(0).Visible = True ''': mnuGPrint(1).Visible = True
    Else
        mnuGPrint(0).Visible = False: mnuGPrint(1).Visible = False
    End If
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
'''    For i = 0 To 3
'''        picOuter(i).Left = 4500
'''    Next i
    
    
    '\\\\\ -------------------------------------------------------- /////
    
'''''    If bGFXReviewer And iView = 1 Then
'''''        picReview.Visible = True
''''''''        sst1.TabCaption(3) = "Client Graphics / Approval Interface"
'''''        sst1.TabCaption(4) = "Approval Interface"
'''''    Else
'''''        picReview.Visible = False
'''''    End If
'''''    picReview.Width = 2400
    
    '///// CREATE ARRAY OF GFX REVIEW CLIENTS \\\\\'
    If bGFXReviewer Then
        Call SetGFXCUNOArray(UserID)
'''        cmdStatusEdit_Notify.Enabled = CheckForNotify
    Else
        sst1.TabVisible(4) = False
        If bFromFloorplan Then sst1.Tab = 0 Else sst1.Tab = 3
    End If
    
    picOuter(4).Top = flxApprove.Top + flxApprove.RowHeight(0)
'''    picOuter(4).Height = flxApprove.Height - flxApprove.RowHeight(0)
    
    picInner(4).Top = 0
    
    With flxApprove
        .Row = 0
        .Col = 1: .CellAlignment = 4: .Text = "Image"
        .Col = 2: .CellAlignment = 4: .Text = "Status"
        .Col = 3: .CellAlignment = 4: .Text = "File Name": .ColAlignment(3) = 1
        .Col = 4: .CellAlignment = 4: .Text = "Comments"
        .Col = 5: .ColAlignment(5) = 4: .Text = "Graphic Type"
        .Col = 6: .ColAlignment(6) = 4: .Text = "Approver"
        .Col = 7: .CellAlignment = 4: .Text = "Original Posting Date"
        .Col = 8: .ColAlignment(8) = 4: .Text = "Poster"
    End With
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    
'''    Randomize
'''    For i = 1 To flxApprove.Rows - 1
'''        flxApprove.RowHeight(i) = 1080 '''1500 '''picRecord(0).Height
'''        For iCol = 2 To 5
'''            Select Case iCol
'''                Case 4
'''                    flxApprove.Row = i: flxApprove.Col = iCol
'''                    Set flxApprove.CellPicture = imgMail(Int(2 * Rnd)).Picture
'''                    flxApprove.CellPictureAlignment = 4
'''                Case Else
'''                    flxApprove.TextMatrix(i, iCol) = _
'''                                "Data Feed Row " & i & _
'''                                ", Column " & iCol & "."
'''            End Select
'''        Next iCol
'''    Next i
'''    picInner(4).Height = (flxApprove.Rows - 1) * 1080 '''1500
'''    picInner(5).Height = (flxApprove.Rows - 1) * 1080 '''1500
    
    If bPassIn Then
        picNav.Visible = False
'''        MsgBox "Attempting to Pass In"
        strSelect = "SELECT GM.AN8_CUNO, C.ABALPH, GM.GID, " & _
                    "GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
                    "NVL(GM.GAPPROVER_ID,0) AS GAPPROVER_ID, " & _
                    "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
                    "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
                    "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
                    "FROM " & GFXMas & " GM, " & F0101 & " C " & _
                    "WHERE GM.GID = " & sPassInValue & " " & _
                    "AND GM.AN8_CUNO = C.ABAN8"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            
            iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
            iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
            
            imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
            spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
            imageX = CLng((picInner(0).ScaleWidth) / iCols) '' (imageY / 3) * 4
            spaceX = imageX
            
            
            ''DO THE STUFF''
            On Error Resume Next
            Err = 0
            If rst.Fields("GSTATUS") = 10 Or rst.Fields("GSTATUS") = 27 Or rst.Fields("GSTATUS") = 20 Then
                If Not bGPJ And Not bGFXReviewer Then
                    rst.Close: Set rst = Nothing
                    MsgBox "You do not currently have rights to Review and Approve this graphic file.  " & _
                                "Please, contact your Account Team and have them reset your access rights.", vbExclamation, "Sorry..."
                    GoTo NonApprover
                End If
                
                cboCUNO(4).Text = Trim(rst.Fields("ABALPH"))
                If cboCUNO(4).Text = Trim(rst.Fields("ABALPH")) Then Err.Clear
                
                If rst.Fields("GAPPROVER_ID") <> UserID Then
'                    bResetting = True
                    iApproverView = 1
                    optApproverView(1).Value = True
                    bResetting = False
                Else ''MEANS THIS IS APPROVER''
                    iApproverView = 0
                    optApproverView(0).Value = True
                    bResetting = False
                End If
                
                
                If Err Then GoTo NonApprover
                sst1.Tab = 4
                For i = 1 To flxApprove.Rows - 1
                    If flxApprove.TextMatrix(i, 0) = rst.Fields("GID") Then
                        imx4_Click (i - 1)
                        bPassIn = False
                        iPassIn = 0
                        Exit For
                    End If
                Next i
                
                Debug.Print "Through with row search"
            Else
NonApprover:
                sst1.Tab = 3
                cboCUNO(3).Text = Trim(rst.Fields("ABALPH"))
''                call tvwGraphics(3).Nodes("D" & Trim(rst.Fields("M2")) & Trim(rst.Fields("Y4"))).Selected = True
                Call tvwGraphics_NodeClick(3, tvwGraphics(3).Nodes("D" & Trim(rst.Fields("M2")) & Trim(rst.Fields("Y4"))))
                
                bFound = False
                For i = 0 To imx3.Count - 1
'                    If imx3(i).Visible = True Then
                        If imx3(i).Tag = sPassInValue Then
                            Call imx3_Click(i)
                            bPassIn = False
                            iPassIn = 0
                            bFound = True
                            Exit For
                        End If
'                    End If
                Next i
                
                If Not bFound Then
                    ''THE FILE IS NOT IN THE CURRENT VIEW - LOAD IT''
                    Call LoadGraphic(13, sPassInValue, "Testing", "", "")
                    
                    
                End If
                    
                    
                ''DIG DOWN THRU TREE FOR CORRECT NODE''
                
            End If
        End If
        rst.Close: Set rst = Nothing
    
    Else
        imgDirs_Click
    End If
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
    cmdFindNext.Left = -1000
    
    lblRN.Caption = "REDLINE NOTE:" & vbNewLine & vbNewLine & _
                "Do not resize the window while you are in Redline Mode.  " & _
                vbNewLine & _
                "Redline Annotations may be lost during the regeneration of the window."
    shpRN.Height = lblRN.Height + ((lblRN.Top - shpRN.Top) * 2)
                
    Screen.MousePointer = 0
    
    
    '********* DEBUG *********'
    If bDebug Then
        iDebug = iDebug + 1
        MsgBox "Debug Point: " & iDebug
    End If
    '********* DEBUG *********'
    
    
'''    imgDirs_Click
    
End Sub

'''Private Sub Form_Paint()
'''    MsgBox "Painting"
'''End Sub

'''Private Sub Form_Paint()
''''''''    ResizeThePicture
'''End Sub

Private Sub Form_Resize()
    Dim i As Integer, iTab As Integer
    Dim lErr As Long
    Dim bTabVis As Boolean
    Dim iCol As Long, iRow As Long
    
    If bRedMode And Me.WindowState <> 1 Then
        MsgBox "Window cannot be resized while in Redline Mode.  To resize, " & _
                    "return to the posted image, then resize.", vbExclamation, "Sorry..."
        Exit Sub
    End If
    
    Debug.Print Me.Width & ", " & Me.Height
    
    lErr = LockWindowUpdate(Me.hwnd)
    
    If Me.WindowState <> 1 Then
        bResizing = True
        
        shpHDR.Width = Me.ScaleWidth
        
        If picTabs.Visible Then
            bTabVis = True
            picTabs.Visible = False
            bDirsOpen = False
        End If
        If Me.Width > 12000 And Me.Height > 8400 Then
            maxX = Me.ScaleWidth - 1320 - 120 '' - 1260 - 120 '' 240
            maxY = Me.ScaleHeight - 1200 - picTools.Height - 120 '' 240 '' 1380 - 240 '' 795 - 120
            
            If bAcro Then
                picPDF.Width = maxX
                picPDF.Height = maxY
'''                pdfGraphic.Width = maxX
'''                pdfGraphic.Height = maxY
            End If
            lblStatus.Left = 120
'            lblStatus.Top = dTop + maxY '' pdfGraphic.Height
            
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
                    If picJPG.Width > maxX Or picJPG.Height > maxY Then
                        With picJPG
                            .Height = maxY
                            .Width = .Height * rAsp ''''' / rFAsp)
                            .Top = dTop
                            .Left = (Me.ScaleWidth - .Width) / 2
                        End With
                    Else
                        With picJPG
                            .Top = dTop + ((maxY - picJPG.Height) / 2)
                            .Left = dLeft + ((maxX - picJPG.Width) / 2)
                        End With
                    End If
            End Select
            
            If bPicLoaded Then
                If picJPG.Visible Then
                    If imgSize.Picture <> 0 Then
                        picJPG.PaintPicture imgSize.Picture, 0, 0, picJPG.Width, picJPG.Height
                        Call SetImageState
                    End If
                End If
            End If
            
'''            '///// POSITION LBLBYGEORGEs \\\\\\
'''            lblByGeorge(0).Left = picJPG.Width - lblByGeorge(0).Width - 240
''''''            lblByGeorge(0).Left = (Me.ScaleWidth - (picTabs.Left * 2)) - 240 - lblByGeorge(0).Width
'''            lblByGeorge(0).Top = picJPG.Height - 240 - lblByGeorge(0).Height
'''            lblByGeorge(1).Left = 240
'''            lblByGeorge(1).Top = lblByGeorge(0).Top + 840
            
            AppWindowState = Me.WindowState
'''            cmdClose.Left = Me.ScaleWidth - 120 - cmdClose.Width
'''            cmdClose.Top = 300 ''120
'''            cmdClose.Refresh
'''            cmdSettings.Left = cmdClose.Left
'''            chkStoreClient.Left = Me.ScaleWidth - 120 - chkStoreClient.Width
            
            imgClose.Left = Me.ScaleWidth - imgClose.Width
''''''''            lblSettings.Left = Me.ScaleWidth - 180 - lblSettings.Width '' imgClose.Left + (imgClose.Width / 2) - (lblSettings.Width / 2)
            lblClose.Left = Me.ScaleWidth - 300 - lblClose.Width '' imgClose.Left + (imgClose.Width / 2) - (lblClose.Width / 2)
            picNav.Left = Me.ScaleWidth - 240 - picNav.Width
            
'''            lblImporter.Left = imgClose.Left - 60 - lblImporter.Width
'''            lblImporter.Top = 180 ''120
            
''''''''            imgSearch.Left = lblClose.Left - 120 - imgSearch.Width
            If imgImporter.Visible Then
''''''''                imgImporter.Left = imgSearch.Left - 120 - imgImporter.Width
                imgImporter.Left = lblClose.Left - 120 - imgImporter.Width
            End If
            
            
'''            If lblImporter.Visible Then
'''                imgSearch.Left = lblImporter.Left - 60 - imgSearch.Width
''''''                imgSettings.Left = imgSearch.Left - 30 - imgSettings.Width
'''                imgSupDoc.Left = imgSearch.Left - 60 - imgSupDoc.Width
'''            Else
'''                imgSearch.Left = imgClose.Left - 60 - imgSearch.Width
''''''                imgSettings.Left = imgSearch.Left - 30 - imgSettings.Width
'''                imgSupDoc.Left = imgSearch.Left - 60 - imgSupDoc.Width
'''            End If
''''''            imgSettings.Top = 120
'''            imgSupDoc.Top = 60 ''120
            
            imgBack(1).Left = imgSupDoc.Left - 300
            imgBack(1).Width = imgClose.Left - imgBack(1).Left + 300
            
                
'''            picGfxApprove.Left = cmdClose.Left + cmdClose.Width - 120 - picGfxApprove.Width
            
            picPrint.Left = (Me.Width - picPrint.Width) / 2
            picPrint.Top = (Me.Height - picPrint.Height) / 2
            
            If bPicLoaded Then
                '///// CHECK IF IMAGE IS STRETCHED \\\\\
                If picJPG.Width > imgSize.Width Then
                    mnuResizeGraphic.Visible = True
                    mnuMaxGraphic.Visible = True
                    mnuResizeGraphic.Enabled = True
                    mnuMaxGraphic.Enabled = False
''''''''                    lblResize.Caption = "Resize"
''''''''                    lblResize.Enabled = True
                    Call ResetJPGZoom(1, 2, 0)
'                ElseIf picJPG.Width = imgSize.Width Then
                    
                Else
                    mnuResizeGraphic.Visible = True
                    mnuMaxGraphic.Visible = True
                    mnuResizeGraphic.Enabled = True
                    mnuMaxGraphic.Enabled = True
''''''''                    lblResize.Caption = "Resize"
''''''''                    lblResize.Visible = True
''''''''                    lblResize.Enabled = True
                    Call ResetJPGZoom(0, 2, 1)
                End If
            End If
            
            iTab = sst1.Tab
            If iView = 1 Then
                picTabs.Height = Me.ScaleHeight - picTabs.Top
                sst1.Height = picTabs.ScaleHeight
                picTabs.Width = Me.ScaleWidth
                sst1.Width = picTabs.ScaleWidth ''' Me.ScaleWidth - (picTabs.Left * 2) '' maxX
'''                picTabs.Width = sst1.Width
                
                
                
                
            End If
            bReSize = True
            For i = 0 To 3
                tvwGraphics(i).Height = sst1.Height - 1155 - 480 '' tvwGraphics(i).Top - 480
                picOuter(i).Height = tvwGraphics(i).Height
                hsc1(i).Top = picOuter(i).Top + picOuter(i).Height - hsc1(i).Height
                picInner(i).Height = picOuter(i).ScaleHeight
                Debug.Print "picInner(" & i & "),height = " & picInner(i).Height
                picOuter(i).Width = picTabs.ScaleWidth - 4500 - 180
                picInner(i).Width = picOuter(i).ScaleWidth
                hsc1(i).Width = picOuter(i).Width
                lblCnt(i).Left = picOuter(i).Left + picOuter(i).Width - lblCnt(i).Width
                
'''                If bPopped(i) Then
'''                    Call tvwGraphics_NodeClick(i, tvwGraphics(i).SelectedItem)
'''                End If
                
                If sst1.TabVisible(i) Then
                    sst1.Tab = i
                    picOuter(i).Left = 4500
'                    If (Me.ScaleWidth - (picTabs.Left * 2)) - picOuter(i).Left - 180 > 0 Then
'                        picOuter(i).Width = (Me.ScaleWidth - (picTabs.Left * 2)) - picOuter(i).Left - 180
'                    End If
'                    hsc1(i).Width = picOuter(i).Width
'                    lblCnt(i).Left = picOuter(i).Left + picOuter(i).Width - lblCnt(i).Width
                End If
            Next i
'''''            Me.picIconSize.Top = picTabs.ScaleHeight - 240 - (picIconSize.Height / 2)
            
            For i = chkClose.LBound To chkClose.UBound
                chkClose(i).Top = sst1.Height - 240 - (chkClose(i).Height / 2)
                chkClose(i).Visible = False
            Next i
            For i = chkApproved.LBound To chkApproved.UBound
                chkApproved(i).Top = sst1.Height - 240 - (chkApproved(i).Height / 2)
            Next i
            For i = lblPipe.LBound To lblPipe.UBound
                lblPipe(i).Top = sst1.Height - 240 - (lblPipe(i).Height / 2)
            Next i
            For i = lblFirst.LBound To lblFirst.UBound
                lblFirst(i).Top = sst1.Height - 240 - (lblFirst(i).Height / 2)
            Next i
            For i = lblPrevious.LBound To lblPrevious.UBound
                lblPrevious(i).Top = sst1.Height - 240 - (lblPrevious(i).Height / 2)
            Next i
            For i = lblNext.LBound To lblNext.UBound
                lblNext(i).Top = sst1.Height - 240 - (lblNext(i).Height / 2)
            Next i
            For i = lblLast.LBound To lblLast.UBound
                lblLast(i).Top = sst1.Height - 240 - (lblLast(i).Height / 2)
            Next i
            For i = lblList.LBound To lblList.UBound
                lblList(i).Top = sst1.Height - 240 - (lblList(i).Height / 2)
            Next i
            For i = lblCnt.LBound To lblCnt.UBound
                lblCnt(i).Top = sst1.Height - 240 - (lblCnt(i).Height / 2)
            Next i
            
            iRows = Int((picInner(0).Height - 120) / (imx0(0).Height + 360))
            iCols = Int(picInner(0).Width / (imx0(0).Width + 480))
            
            imageY = (picInner(0).ScaleHeight - 120) / iRows '' (imx0(0).Height + 480)  '' (picInner(0).Height - hsc1(0).Height - 240 - 900) / iRows
            spaceY = imageY '' (imx0(0).Height + 480) '' imageY + 270 '''300
            imageX = CLng((picInner(0).ScaleWidth) / iCols) '' (imageY / 3) * 4
            spaceX = imageX ''CLng(picInner(0).ScaleWidth / iCols) '' imageX + 720 ''240
            
            
            For i = imx0.LBound To imx0.UBound
                iCol = Int(i / iRows): iRow = i Mod iRows
                imx0(i).Left = ((imageX - imx0(0).Width) / 2) + (iCol * spaceX)
                imx0(i).Top = 120 + (iRow * spaceY)
                lbl0(i).Left = imx0(i).Left + (imx0(i).Width / 2) - (lbl0(i).Width / 2)
                lbl0(i).Top = imx0(i).Top + imx0(i).Height + 60
            Next i
            For i = imx1.LBound To imx1.UBound
                iCol = Int(i / iRows): iRow = i Mod iRows
                imx1(i).Left = ((imageX - imx1(0).Width) / 2) + (iCol * spaceX)
                imx1(i).Top = 120 + (iRow * spaceY)
                lbl1(i).Left = imx1(i).Left + (imx1(i).Width / 2) - (lbl1(i).Width / 2)
                lbl1(i).Top = imx1(i).Top + imx1(i).Height + 60
                shp1(i).Left = lbl1(i).Left + 60
                shp1(i).Top = lbl1(i).Top + 45
            Next i
            For i = imx2.LBound To imx2.UBound
                iCol = Int(i / iRows): iRow = i Mod iRows
                imx2(i).Left = ((imageX - imx2(0).Width) / 2) + (iCol * spaceX)
                imx2(i).Top = 120 + (iRow * spaceY)
                lbl2(i).Left = imx2(i).Left + (imx2(i).Width / 2) - (lbl2(i).Width / 2)
                lbl2(i).Top = imx2(i).Top + imx2(i).Height + 60
                shp2(i).Left = lbl2(i).Left + 60
                shp2(i).Top = lbl2(i).Top + 45
            Next i
            For i = imx3.LBound To imx3.UBound
                iCol = Int(i / iRows): iRow = i Mod iRows
                imx3(i).Left = ((imageX - imx3(0).Width) / 2) + (iCol * spaceX)
                imx3(i).Top = 120 + (iRow * spaceY)
                lbl3(i).Left = imx3(i).Left + (imx3(i).Width / 2) - (lbl3(i).Width / 2)
                lbl3(i).Top = imx3(i).Top + imx3(i).Height + 60
                shp3(i).Left = lbl3(i).Left + 60
                shp3(i).Top = lbl3(i).Top + 45
            Next i
            
            picMess.Left = (Me.ScaleWidth - picMess.Width) / 2
            
            picWait.Left = picTabs.Left + 4500 + _
                        ((picOuter(iTab).Width - picWait.Width) / 2)
            picWait.Top = picTabs.Top + picOuter(1).Top + _
                        ((picOuter(iTab).Height - picWait.Height) / 2)
            If picWait.Visible Then picWait.Refresh
            
            
            For i = 0 To 3
                If bPopped(i) Then
                    Call tvwGraphics_NodeClick(i, tvwGraphics(i).SelectedItem)
                End If
            Next i
            
            sst1.Tab = iTab
            bReSize = False
            
            With flxApprove
                .Width = sst1.Width - 360
                .ColWidth(0) = 0
                .ColWidth(1) = 1515
                .ColWidth(2) = 1815
                .ColWidth(3) = (flxApprove.Width - 3330 - 240 - 960 - 1080 - 1200 - 1080) / 2
                .ColWidth(4) = 960
                .ColWidth(5) = 1200
                .ColWidth(6) = 1080 ''960 '''1200
                .ColWidth(7) = (flxApprove.Width - 3330 - 240 - 960 - 1080 - 1200 - 1080) / 2
                .ColWidth(8) = 1080 ''960'''1200
                .ColWidth(9) = 240
            End With
            rCommX1 = CSng(flxApprove.ColPos(4))
            rCommX2 = CSng(flxApprove.ColPos(5) - 1)
            rApproverX1 = CSng(flxApprove.ColPos(6))
            rApproverX2 = CSng(flxApprove.ColPos(7) - 1)
            
            flxApprove.Height = picTabs.ScaleHeight - flxApprove.Top - 180 '' flxApprove.Left
            picOuter(4).Height = flxApprove.Height - flxApprove.RowHeight(0)
            
            picReview.Left = flxApprove.Left + flxApprove.Width - picReview.Width
'            picReview.Top = fraRefresh.Top + cmdRefresh.Top + cmdRefresh.Height - picReview.Height
            picReview.Top = cboFolder.Top + cboFolder.Height - picReview.Height
            
            lblFileCount.Left = flxApprove.Left + flxApprove.Width - lblFileCount.Width
            
            picHelp.Top = flxApprove.Top '''tvwGraphics(3).Top
            picHelp.Left = flxApprove.Left '''tvwGraphics(3).Left
            picHelp.Width = flxApprove.Width '''picOuter(3).Left + picOuter(3).Width - tvwGraphics(3).Left
            picHelp.Height = flxApprove.Height '''tvwGraphics(3).Height
            cmdHelpClose.Left = picHelp.ScaleWidth - cmdHelpClose.Width - 360
            cmdHelpClose.Top = 120
            web1.Width = picHelp.ScaleWidth '''- (web1.Left * 2)
            web1.Height = picHelp.ScaleHeight '''- 540 - web1.Left
            
'            picXData.Left = Me.ScaleWidth - picXData.Width - 90
'            picXData.Top = Me.ScaleHeight - picXData.Height - 90
'            picXD.Top = Me.ScaleHeight - picXD.Height - 30
'            picXD.Left = Me.ScaleWidth - picXD.Width - 30
            
            
            
            picType.Left = sst1.Left + sst1.Width - 180 - picType.Width
            
            Select Case iModeTab
                Case 4
                    fraMulti.Left = picReview.Left - fraMulti.Width
                    fraMulti.Top = picReview.Top + picReview.Height - fraMulti.Height
                Case Else
                    fraMulti.Left = picType.Left - fraMulti.Width
                    fraMulti.Top = picType.Top + picType.Height - fraMulti.Height
            End Select
            
'''            picWait.Left = picTabs.Left + 4500 + _
'''                        ((picOuter(1).Width - picWait.Width) / 2)
'''            picWait.Top = picTabs.Top + picOuter(1).Top + _
'''                        ((picOuter(1).Height - picWait.Height) / 2)
            
            
            picPDF.Width = maxX '' Me.ScaleWidth - picPDF.Left - 120
            picPDF.Height = maxY '' Me.ScaleHeight - picPDF.Top - 120
            
            picTools.Top = Me.ScaleHeight - picTools.Height '' picPDF.ScaleHeight - picTools.Height
            picTools.Width = Me.ScaleWidth '' picPDF.ScaleWidth
            
            
            
            
            If picTools.ScaleWidth / 2 > picPDFTools.Width Then
                picPDFTools.Left = picTools.ScaleWidth / 2
            Else
                picPDFTools.Left = picTools.ScaleWidth - picPDFTools.Width
            End If
            picRedTools.Left = picPDFTools.Left
            picJPGTools.Left = picPDFTools.Left
            
            xpdf1.Width = picPDF.ScaleWidth ''- (Xpdf1.Left * 2) '' Me.ScaleWidth - Xpdf1.Left - 120
            xpdf1.Height = picPDF.ScaleHeight ''- picTools.Height - (Xpdf1.Top * 2) '' Me.ScaleHeight - Xpdf1.Top - 120
            
            picROuter.Height = xpdf1.Height
            picROuter.Width = xpdf1.Width - vsc1.Width
            
            cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
            
            Me.picRedTools.Top = 0
            
            vsc1.Left = picPDF.ScaleWidth - vsc1.Width
            vsc1.Height = picROuter.Height
            picRed.Width = xpdf1.Width - vsc1.Width
            picRed.Height = xpdf1.Height
            
            shpRN.Top = picTools.Top - shpRN.Left - shpRN.Height
            lblRN.Top = shpRN.Top + 120
            
'            picPDFMenu.Top = Me.ScaleHeight - picPDFMenu.Height
            
''            picPDFOptions.Top = picPDF.Top - picPDFOptions.Height
''            picPDFOptions.Left = picPDF.Left + picPDF.Width - picPDFOptions.Width
            
            txtRed.Left = Me.Width
    
            AppWindowState = Me.WindowState
        End If
    End If
    
    bResizing = False
    
    If bTabVis Then
        picTabs.Visible = True
        bDirsOpen = True
    End If
    
    lErr = LockWindowUpdate(0)
End Sub

Public Sub PopInventory(lCUNO As Long)
    Dim sNode As String, sParent As String, sDesc As String, sDescPar As String, sStat As String
    Dim sKNode As String, sENode As String, sTNode As String, sGNode As String, sFNode As String
    Dim nodX As Node
    Dim iFile As Integer, i As Integer, iNode As Integer, iNo As Integer, iType As Integer
    Dim lParent As Long, lElem As Long
    Dim strSelect As String, strNest As String
    Dim rst As ADODB.Recordset
    Dim sGStatus(0 To 30) As String
    Dim iGStatus(0 To 30) As Integer
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(0) = "DE-ACTIVATED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    iGStatus(0) = 10
    iGStatus(10) = 7
    iGStatus(20) = 8
    iGStatus(30) = 9
    
    
    tvwGraphics(2).Visible = False
    tvwGraphics(2).Nodes.Clear
    tvwGraphics(2).ImageList = ImageList1
    sKNode = "": sENode = "": sTNode = "": sGNode = ""
    sNode = "": sParent = ""

    '///// FIRST GET KITS \\\\\
    strSelect = "SELECT K.KITID, K.KITFNAME " & _
                "FROM IGLPROD.IGL_KIT K " & _
                "Where K.AN8_CUNO = " & lCUNO & " " & _
                "AND K.KSTATUS > 0 " & _
                "ORDER BY K.KITREF"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        sKNode = "k" & rst.Fields("KITID")
        sDesc = UCase(Trim(rst.Fields("KITFNAME"))) & " Kit"
        sDescPar = UCase(Trim(rst.Fields("KITFNAME")))
        Set nodX = tvwGraphics(2).Nodes.Add(, , sKNode, sDesc, 6)
        rst.MoveNext
    Loop
    rst.Close
    
    ''NOW, CHECK FOR KIT FOLDERS''
    strSelect = "SELECT DISTINCT K.KITID, GM.FLR_ID, GF.FLRDESC, " & _
                "GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                "FROM IGLPROD.IGL_KIT K, ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                "Where K.AN8_CUNO = " & lCUNO & " " & _
                "AND K.KITID = GE.ELTID " & _
                "AND GE.GID = GM.GID " & _
                "AND GM.GSTATUS IN (" & defSIN & ") " & _
                "AND GM.FLR_ID > 0 " & _
                "AND GM.FLR_ID = GF.FLR_ID"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case iView
            Case 1
                If sKNode <> "k" & rst.Fields("KITID") Then
                    sKNode = "k" & rst.Fields("KITID")
                    sFNode = "f" & rst.Fields("KITID") & "-" & rst.Fields("FLR_ID")
                    sDesc = Trim(rst.Fields("FLRDESC"))
                    iType = rst.Fields("FLAG") + 14
                    i = tvwGraphics(2).Nodes(sKNode).Index + 1
                    Set nodX = tvwGraphics(2).Nodes.Add(sKNode, tvwChild, sFNode, sDesc, iType)
                End If
        End Select
        rst.MoveNext
    Loop
    rst.Close
    
    ''GET ELEMENTS''
    sKNode = "": sDescPar = ""
    strSelect = "SELECT K.KITID, K.KITFNAME, E.ELTID, E.ELTFNAME, E.ELTDESC " & _
                "FROM " & IGLKit & " K, " & IGLElt & " E " & _
                "WHERE K.AN8_CUNO = " & lCUNO & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ESTATUS > 2 " & _
                "ORDER BY K.KITREF, E.ELTFNAME"
    Debug.Print strSelect
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sKNode <> "k" & rst.Fields("KITID") Then
            sKNode = "k" & rst.Fields("KITID")
            sDescPar = Trim(rst.Fields("KITFNAME"))
        End If
        If sENode <> "e" & rst.Fields("ELTID") Then
            sENode = "e" & rst.Fields("ELTID")
            sDesc = sDescPar & "-" & UCase(Trim(rst.Fields("ELTFNAME"))) & "  " & _
                        UCase(Trim(rst.Fields("ELTDESC")))
            Set nodX = tvwGraphics(2).Nodes.Add(sKNode, tvwChild, sENode, sDesc, 6)
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    '///// NOW, GET ELEMENT GRAPHICS \\\\\
    strNest = "SELECT E.ELTID " & _
                "FROM " & IGLKit & " K, " & IGLElt & " E " & _
                "WHERE K.AN8_CUNO = " & lCUNO & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ESTATUS > 2"
    If bPerm(29) Then
        If iView = 0 Then
            strSelect = "SELECT GE.ELTID, GE.ES_ID, GM.GDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GE.ELTID IN (" & strNest & ") " & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & defSIN & ")"
        Else
'''            strSelect = "SELECT DISTINCT GE.ELTID, GM.GTYPE, GM.GSTATUS " & _
'''                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
'''                        "WHERE GE.ELTID IN (" & strNest & ") " & _
'''                        "AND GE.GID = GM.GID " & _
'''                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
'''                        "ORDER BY GM.GTYPE, GM.GSTATUS"
'''            strSelect = "SELECT DISTINCT GE.ELTID, GM.GTYPE " & _
'''                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
'''                        "WHERE GE.ELTID IN (" & strNest & ") " & _
'''                        "AND GE.GID = GM.GID " & _
'''                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
'''                        "ORDER BY GM.GTYPE"
            strSelect = "SELECT DISTINCT GE.ELTID " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GE.ELTID IN (" & strNest & ") " & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & defSIN & ")"
            
        End If
    Else
        If iView = 0 Then
            strSelect = "SELECT GE.ELTID, GE.ES_ID, GM.GDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GID = GE.GID"
        Else
'''            strSelect = "SELECT DISTINCT GE.ELTID, GM.GTYPE, GM.GSTATUS " & _
'''                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GTYPE <> 3 " & _
'''                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
'''                        "AND GM.GID = GE.GID " & _
'''                        "ORDER BY GM.GTYPE, GM.GSTATUS"
'''            strSelect = "SELECT DISTINCT GE.ELTID, GM.GTYPE " & _
'''                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GTYPE <> 3 " & _
'''                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
'''                        "AND GM.GID = GE.GID " & _
'''                        "ORDER BY GM.GTYPE"
            strSelect = "SELECT DISTINCT GE.ELTID " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GID = GE.GID"
        End If
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
'''        sENode = "e" & rst.Fields("ELTID")
        Select Case iView
            Case 0
                sENode = "e" & rst.Fields("ELTID")
                sGNode = "g" & rst.Fields("ES_ID")
                sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
                iType = rst.Fields("GTYPE")
                Set nodX = tvwGraphics(2).Nodes.Add(sENode, tvwChild, sGNode, sDesc, iType)
                nodX.Parent.Image = 5
                nodX.Parent.Parent.Image = 5
            Case 1
'''                Select Case Len(rst.Fields("GSTATUS"))
'''                    Case 1
'''                        sStat = "0" & CStr(rst.Fields("GSTATUS"))
'''                    Case 2
'''                        sStat = CStr(rst.Fields("GSTATUS"))
'''                End Select
                If sENode <> "e" & rst.Fields("ELTID") Then
                    sENode = "e" & rst.Fields("ELTID")
                    Set nodX = tvwGraphics(2).Nodes(sENode)
                    nodX.Image = 12
                    nodX.Parent.Image = 13
                End If
'''                If sTNode <> "t" & rst.Fields("GTYPE") & rst.Fields("ELTID") Then
'''                    iType = rst.Fields("GTYPE")
'''                    sTNode = "t" & rst.Fields("ELTID") & "-" & rst.Fields("GTYPE")
'''                    sDesc = GfxType(rst.Fields("GTYPE"))
'''                    Set nodX = tvwGraphics(2).Nodes.Add(sENode, tvwChild, sTNode, sDesc, iType)
'''                    nodX.Parent.Image = 5
'''                    nodX.Parent.Parent.Image = 5
'''                End If
'''                sGNode = "i" & rst.Fields("GTYPE") & sStat & "-" & rst.Fields("ELTID")
'''                sDesc = sGStatus(rst.Fields("GSTATUS")) ''' & " Graphics"
'''                Set nodX = tvwGraphics(2).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, _
'''                            iGStatus(rst.Fields("GSTATUS")))
        End Select
'''        iType = rst.Fields("GTYPE")
'''        Set nodX = tvwGraphics(2).Nodes.Add(sENode, tvwChild, sGNode, sDesc, iType)
'''        nodX.Parent.Image = 5
'''        nodX.Parent.Parent.Image = 5
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    
    '///// NOW, GET KIT GRAPHICS \\\\\
    strNest = "SELECT K.KITID " & _
                "FROM " & IGLKit & " K " & _
                "WHERE K.AN8_CUNO = " & lCUNO & " " & _
                "AND K.KSTATUS > 0"
    If bPerm(29) Then
        If iView = 0 Then
            strSelect = "SELECT GE.ELTID, GE.ES_ID, GM.GDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GE.ELTID IN (" & strNest & ") " & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & defSIN & ")"
        Else
            strSelect = "SELECT DISTINCT GE.ELTID " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GE.ELTID IN (" & strNest & ") " & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & defSIN & ")"
            
        End If
    Else
        If iView = 0 Then
            strSelect = "SELECT GE.ELTID, GE.ES_ID, GM.GDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GID = GE.GID"
        Else
            strSelect = "SELECT DISTINCT GE.ELTID " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GID = GE.GID"
        End If
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case iView
            Case 1
                If sKNode <> "k" & rst.Fields("ELTID") Then
                    sKNode = "k" & rst.Fields("ELTID")
                    Set nodX = tvwGraphics(2).Nodes(sKNode)
                    nodX.Image = 21
                End If
        End Select
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing

    
    
    
    
    
    
    
    
    Set nodX = Nothing
    tvwGraphics(2).Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strUpdate As String
    Dim bCheck As Boolean
    Dim Resp As VbMsgBoxResult
    Dim i As Integer, iCancel As Integer
    
    If bRedded Then
        Select Case ShallWeSave
            Case 0 ''NO''
            Case 1 ''YES''
            Case 2 ''CANCEL''
                Cancel = 2
                Exit Sub
        End Select
'''        If Not ShallWeSave Then Exit Sub
    End If
    
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassGID = lGID
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedSaved = False
    bRedMode = False
    
    Call ClearUndo(0)
    
    If bAddMode Then Unload frmAssign
    
'''    If bAcro Then Me.Controls.Remove pdfGraphic.Name
'''    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
    
    bGfxOpen = False
    
    
    
    
    If lblRed(0).Container.Name <> picRed.Name Then
        Call ResetLBLREDContainer(picRed)
    End If
        
    For i = LBound(bPopped) To UBound(bPopped)
        bPopped(i) = False
    Next i
'''    '///// CHECK FOR UN-NOTIFIED STATUS RESETS \\\\\'
'''    bCheck = CheckForNotify
'''    If bCheck Then
'''        Resp = MsgBox("You have either Posted new Graphic Files in the Graphic Importer, " & _
'''                    "or reset the Status on Graphic Files in the Annotator, " & _
'''                    "without sending out a Notification.  " & _
'''                    "These files will not be available to you, or other Annotator Users, " & _
'''                    "until a Notification is sent out." & vbNewLine & vbNewLine & _
'''                    "Do you want to access the Notification Interface before exiting?", _
'''                    vbExclamation + vbYesNoCancel, "New Status Changes...")
'''        If Resp = vbYes Then
'''            frmNotify.Show 1
''''''            Cancel = 1
''''''            GoTo CancelIt
'''        ElseIf Resp = vbCancel Then
'''            Cancel = 1
'''            GoTo CancelIt
'''        End If
'''    End If


    '///// KILL LOCK FILE, IF ACTIVE \\\\\
    If lNewLockId <> 0 Then
        strUpdate = "UPDATE " & ANOLockLog & " " & _
                    "SET LOCKCLOSEDTTM = SYSDATE, " & _
                    "LOCKSTATUS = LOCKSTATUS * -1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, " & _
                    "UPDCNT = UPDCNT + 1 " & _
                    "WHERE LOCKID = " & lNewLockId
        Conn.Execute (strUpdate)
    End If

    tSHYR = 0
    tBCC = ""
    tSHCD = 0
    tFBCN = ""
    tSHNM = ""
    
    
CancelIt:
End Sub

Private Sub hsc1_Change(Index As Integer)
    picInner(Index).Left = CLng(hsc1(Index).Value) * (-100)
End Sub

Private Sub hsc1_Scroll(Index As Integer)
    picInner(Index).Left = CLng(hsc1(Index).Value) * (-100)
End Sub

Private Sub imgComm_Click()
    With frmComments
        .PassREFID = lRefID
        .PassTable = sTable
        .PassIType = 1
        .PassBCC = tBCC
        .PassFBCN = tFBCN
        .PassSHCD = tSHCD
        .PassMessPath = lblWelcome.Caption
        .PassMessSub = lblGraphic.Caption
        .PassForm = "frmGraphics"
        If picPDF.Visible Then
            .PassGPath = sGPath & "Acrobatid.bmp"
        Else
            .PassGPath = CurrFile
        End If

        .Show 1
    End With
End Sub

'''Private Sub imgGfxApprove_Click()
'''    picGfxApprove.Visible = False
'''End Sub

'''Private Sub imgMinMax_Click()
'''    picGfxApprove.Visible = False
'''    cmdGfxApproveHide.Visible = True
'''End Sub


'''Private Sub imgSettings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Me.PopupMenu mnuSettings
'''End Sub

'''Private Sub imgStat_Click(Index As Integer)
'''    MsgBox "Reset Status for Image Row " & CInt(Index + 1)
'''End Sub

Private Sub imgStat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mnuOpen.Visible = False
    Call imx4_MouseDown(Index, vbRightButton, Shift, _
                X + imgStat(Index).Left - imx4(Index).Left, _
                Y + imgStat(Index).Top - imx4(Index).Top)
    
End Sub


Public Sub imgStatus_Click(Index As Integer)
    iTabStatus = Index
    Call SetStatus(iTabStatus)
End Sub

Private Sub imgType_Click(Index As Integer)
    sInType(sst1.Tab) = CStr(Index)
    Call SetType(Index)
    iTabType(sst1.Tab) = Index
    
    On Error Resume Next
    Call tvwGraphics_NodeClick(sst1.Tab, tvwGraphics(sst1.Tab).SelectedItem)
    Screen.MousePointer = 0
End Sub

Private Sub imgV_Click(Index As Integer)
    frmVersions.PassHDR = flxApprove.TextMatrix(Index + 1, 3)
    frmVersions.PassGID = CLng(imgV(Index).Tag)
    frmVersions.PassVID = CLng(imx4(Index).Tag)
    frmVersions.PassIndex = Index
    frmVersions.Show 1, Me
End Sub

Private Sub imx0_Click(Index As Integer)
    If bRedded Then
        Select Case ShallWeSave
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedMode = False: If Not bHideRN Then Call RedNoteVis(bRedMode) ''shpRN.Visible = False: lblRN.Visible = False
    bRedSaved = False
    redBCC = "": redSHCD = 0

'    If chkClose(0).value = 1 Then
        picTabs.Visible = False ''sst1.visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'    End If
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    iImage(0) = Index
    Call LoadGraphic(10, imx0(Index).Tag, lbl0(Index).Caption, CurrParNode(0), CurrParText(0))
    
End Sub

Private Sub imx1_Click(Index As Integer)
    If bRedded Then
        Select Case ShallWeSave
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedMode = False: If Not bHideRN Then Call RedNoteVis(bRedMode)
    bRedSaved = False
    redBCC = "": redSHCD = 0

'    If chkClose(1).value = 1 Then
        picTabs.Visible = False ''sst1.visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'    End If
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    iImage(1) = Index
    Call LoadGraphic(11, imx1(Index).Tag, lbl1(Index).Caption, CurrParNode(1), CurrParText(1))
'''    Call LoadThePicture(imx1(index).FileName)
End Sub

Private Sub imx1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        CurrIndex = Index
        Me.PopupMenu mnuGfx
    End If
End Sub

Private Sub imx2_Click(Index As Integer)
    If bRedded Then
        Select Case ShallWeSave
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedMode = False: If Not bHideRN Then Call RedNoteVis(bRedMode)
    bRedSaved = False
    redBCC = "": redSHCD = 0

'    If chkClose(2).value = 1 Then
        picTabs.Visible = False ''sst1.visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'    End If
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    iImage(2) = Index
    Call LoadGraphic(12, imx2(Index).Tag, lbl2(Index).Caption, CurrParNode(2), CurrParText(2))
End Sub

Private Sub imx3_Click(Index As Integer)
    If bRedded Then
        Select Case ShallWeSave
            Case 2 ''CANCEL''
                Exit Sub
        End Select
    End If
    
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedMode = False: If Not bHideRN Then Call RedNoteVis(bRedMode)
    bRedSaved = False
    redBCC = "": redSHCD = 0

'    If chkClose(3).value = 1 Then
        picTabs.Visible = False ''sst1.visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'    End If
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    iImage(3) = Index
    Call LoadGraphic(13, imx3(Index).Tag, lbl3(Index).Caption, CurrParNode(3), CurrParText(3))
End Sub


'''Private Sub imx3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If Button = vbRightButton And bGFXReviewer Then
'''        If CheckGFXCUNO(CLng(tBCC)) Then
'''            Debug.Print "imx3(" & Index & ") selected --- " & imx3(Index).Tag
'''            StatusGID = CLng(imx3(Index).Tag)
'''            mnuName(0).Caption = tvwGraphics(3).SelectedItem.Text & ": " & lbl3(Index).Caption
'''            mnuName(1).Caption = ""
'''            mnuName(0).Tag = "ONE"
'''            Me.PopupMenu mnuResetStatus
'''        End If
'''    End If
'''End Sub

Private Sub imx4_Click(Index As Integer)
    If Not bApproveDown Then
        iImage(4) = Index
        If bRedded Then
            Select Case ShallWeSave
                Case 2 ''CANCEL''
                    Exit Sub
            End Select
        End If
        Call PrepareToOpen(Index)
    End If
'''    If bRedded = True And bTeam = True Then
'''        With frmRedAlert
'''            .PassBCC = CLng(redBCC)
'''            .PassSHCD = redSHCD
'''            .PassHDR = lblWelcome
'''            .PassType = 1
'''            .Show 1
'''        End With
'''    End If
'''    bRedded = False
'''    redBCC = "": redSHCD = 0
'''
'''    If chkClose(4).value = 1 Then
'''        pictabs.visible = False ''sst1.visible = False
'''        imgDirs.tooltiptext = "Click to Open File Index..."
'''    End If
'''    '///// TIME TO LOAD THE GRAPHIC \\\\\
'''    Call LoadGraphic(14, flxApprove.TextMatrix(Index + 1, 0), flxApprove.TextMatrix(Index + 1, 3), "", "Testing")
End Sub

Private Sub imx4_DblClick(Index As Integer)
    Call PrepareToOpen(Index)
End Sub


Private Sub imx4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If Button <> vbRightButton Then bApproveDown = False

    CurrIndex = Index
    Select Case UCase(lblStat(Index))
        Case "INTERNAL": i = 0
        Case "CLIENT DRAFT": i = 1
        Case "APPROVED": i = 2
        Case "RETURNED": i = 4
        Case Else: i = 3
    End Select
    StatusGID = flxApprove.TextMatrix(Index + 1, 0) ''' CLng(imx4(Index).Tag)
    
    If Button = vbRightButton And bGFXReviewer Then
        If CheckGFXCUNO(CLng(fBCC(4))) Then
            bApproveDown = True
            iApprovalRow = Index + 1
            
            frmGfxApprove.PassX = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + _
                        picTabs.Left + picOuter(4).Left + picInner(4).Left + _
                        imx4(Index).Left + X
            frmGfxApprove.PassY = Me.Top + (Me.Height - Me.ScaleHeight) + _
                        picTabs.Top + picOuter(4).Top + picInner(4).Top + _
                        imx4(Index).Top + Y
            frmGfxApprove.PassHDR = "Graphic Approval for '" & flxApprove.TextMatrix(Index + 1, 3) & "'"
            frmGfxApprove.PassVal = i
            frmGfxApprove.PassBCC = CLng(fBCC(4))
            frmGfxApprove.PassFBCN = fFBCN(4)
            frmGfxApprove.PassType = "ONE"
            frmGfxApprove.Show 1, Me
                        
        End If
    End If
End Sub


Private Sub lblCount_Click(Index As Integer)
    Dim i As Integer
    
    Me.MousePointer = 11
    picWait.Visible = True
    picWait.Refresh
    
    For i = (0 + (sst1.Tab * 10)) To (9 + (sst1.Tab * 10))
        If i = Index Then
            lblCount(i).ForeColor = vbRed
            lblCount(i).Refresh
'''            lblCount(i).FontBold = True
        Else
            If lblCount(i).Visible Then
                lblCount(i).ForeColor = vbBlack
'''                lblCount(i).FontBold = False
            End If
        End If
    Next i
    TNode = tvwGraphics(sst1.Tab).SelectedItem.Key
    Call GetGraphics(sst1.Tab, Index, CurrSelect(sst1.Tab), lblCount(Index).Caption, TNode)
    
    picWait.Visible = False
    Me.MousePointer = 0
End Sub

Private Sub lblDisplayAll_Click()
    sInType(sst1.Tab) = "1, 2, 3, 4"
    Call SetType(0)
    
    On Error Resume Next
    Call tvwGraphics_NodeClick(sst1.Tab, tvwGraphics(sst1.Tab).SelectedItem)
    Screen.MousePointer = 0
End Sub

Private Sub lblDownload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xval As Single, Yval As Single
    
    lblDownload.ForeColor = vbRed
    lblEmail.ForeColor = vbButtonText
    
    Xval = picTabs.Left + fraMulti.Left + lblDownload.Left
    Yval = picTabs.Top + fraMulti.Top + lblDownload.Top + lblDownload.Height
    Me.PopupMenu mnuDownloadMulti, , Xval, Yval
    
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xval As Single, Yval As Single
    
    lblDownload.ForeColor = vbButtonText
    lblEmail.ForeColor = vbRed
    
    Xval = picTabs.Left + fraMulti.Left + lblEmail.Left
    Yval = picTabs.Top + fraMulti.Top + lblEmail.Top + lblEmail.Height
    Me.PopupMenu mnuEmailMulti, , Xval, Yval
End Sub

Public Sub lblFilterAll_Click()
    iTabStatus = 0
    Call SetStatus(0)
End Sub

Private Sub lblFirst_Click(Index As Integer)
    Call SetBatch(Index, "FIRST")
End Sub

Private Sub lblLast_Click(Index As Integer)
    Call SetBatch(Index, "LAST")
End Sub

Private Sub lblList_Click(Index As Integer)
    frmGfxList.PassFrom = Me.Name
    frmGfxList.PassSQL = CurrSelect(Index)
    frmGfxList.PassSize = (iRows * iCols)
    frmGfxList.Show 1, Me

End Sub

Private Sub lblNext_Click(Index As Integer)
    Call SetBatch(Index, "NEXT")
End Sub

Private Sub lblPrevious_Click(Index As Integer)
    Call SetBatch(Index, "PREVIOUS")
End Sub

'''Private Sub lblSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblSettings.ForeColor = vbWhite
'''End Sub

'''Private Sub lblStat_Click(Index As Integer)
'''    MsgBox "Reset Status for Image Row " & CInt(Index + 1)
'''End Sub

Private Sub lblStat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'''    Call imx4_MouseDown(Index, vbRightButton, Shift, X, Y)
    Call imx4_MouseDown(Index, vbRightButton, Shift, _
                X + lblStat(Index).Left - imx4(Index).Left, _
                Y + lblStat(Index).Top - imx4(Index).Top)
End Sub

Private Sub lblViewAll_Click(Index As Integer)
    Dim i As Integer
    Me.MousePointer = 11
    lblViewAll(Index).ForeColor = vbRed
    lblViewAll(Index).Refresh
    For i = (0 + (Index * 10)) To (9 + (Index * 10))
        If lblCount(i).Visible Then
            lblCount(i).ForeColor = vbBlack
            lblCount(i).Refresh
        End If
    Next i
    TNode = tvwGraphics(Index).SelectedItem.Key
    Call GetGraphics(Index, 99, CurrSelect(Index), "1-1000", TNode)
    Me.MousePointer = 0
End Sub

Private Sub lblSettings_Click()
    frmSettings.PassFrom = "GH"
    frmSettings.PassBCC = Val(fBCC(sst1.Tab))
    frmSettings.PassFBCN = fFBCN(sst1.Tab)
    frmSettings.PassBCC_DEF = defCUNO ''' lBCC_Def
    frmSettings.PassFBCN_DEF = defFBCN ''' sFBCN_Def
    frmSettings.Show 1
End Sub

Private Sub lstUndo_Click()
    frmUndoTest.Show 1, Me
End Sub

Private Sub mnuAssign_Click()
    bAddMode = True
    With frmAssign
        .PassBCC = tBCC
        .PassFBCN = tFBCN
        .PassSHYR = tSHYR
        .PassSHCD = tSHCD
        .PassSHNM = tSHNM
        On Error Resume Next
        .Show 1, Me
'''        If Err Then
'''            MsgBox "At this time, you cannot access the Assignment Interface when coming in from " & _
'''                        "the Floorplan Viewer.  To access the Assignment Interface, close both the " & _
'''                        "Graphics Viewer and the Floorplan Viewer, and then come directly into " & _
'''                        "the Graphics Interface from the opening screen.", vbExclamation, "Sorry..."
'''            Err.Clear
'''        End If
    End With

End Sub

Private Sub mnuCheckUsage_Click()
    Call CheckShows(CurrNode, CurrNodeText, 0)
End Sub

Private Sub mnuCheckUse_Click()
    Call CheckShows("G" & imx1(CurrIndex).Tag, lbl1(CurrIndex).Caption, 1)
End Sub

Private Sub mnuCommThumb_Click()
'''    MsgBox "imx" & sst1.Tab & "(" & CurrIndex & ") - " & imx1(CurrIndex).Tag
        
    
    With frmComments
        .PassREFID = CLng(imx1(CurrIndex).Tag)
        .PassTable = sTable
        .PassIType = 1
        .PassBCC = fBCC(sst1.Tab)
        .PassFBCN = fFBCN(sst1.Tab)
        .PassSHCD = fSHCD(sst1.Tab)
        .PassMessPath = fFBCN(sst1.Tab) & " " & sst1.TabCaption(sst1.Tab) & ": " & lbl1(CurrIndex).Caption ''' lblWelcome.Caption
        .PassMessSub = lbl1(CurrIndex).Caption ''' lblGraphic.Caption
        .PassForm = "frmGraphics"
        .Show 1
    End With
End Sub

Private Sub mnuDownload_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, sChk As String, sPath As String
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
'''        Exit Sub
'''    Else

    If pDownloadPath = "" Then
        Exit Sub
    Else
        Screen.MousePointer = 11
        On Error GoTo BadFile
        sFolder = pDownloadPath '' shlFolder.Items.Item.Path
        
        If iAppConn = 1 And UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        ElseIf iAppConn = 2 And UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto this Thin-Client drive." & _
                        vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        Err = 0
        On Error GoTo ErrorTrap
        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                    "FROM " & GFXMas & " " & _
                    "WHERE GID = " & lGID
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
        
        Screen.MousePointer = 0
        MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
    End If
    
Exit Sub
ErrorTrap:
    Screen.MousePointer = 0
    rst.Close: Set rst = Nothing
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

Private Sub mnuDownloadMode_Click()
    Dim i As Integer
    
    Err.Clear
    
    ''MAKE CERTAIN NOT IN EMAIL MODE''
    mnuEmailMode.Checked = 0
    lblEmail.ForeColor = vbButtonText
    bEMode = False
    
    iModeTab = sst1.Tab
        
    Select Case mnuDownloadMode.Checked
        Case Is = True ''END DOWNLOAD MODE''
            mnuDownloadMode.Checked = False
            lblDownload.ForeColor = vbButtonText
            Select Case iModeTab
                Case 1
                    For i = 0 To imx1.Count - 1
                        imx1(i).Enabled = True: chk1(i).Visible = False: chk1(i).Value = 0
                    Next i
                Case 2
                    For i = 0 To imx2.Count - 1
                        imx2(i).Enabled = True: chk2(i).Visible = False: chk2(i).Value = 0
                    Next i
                Case 3
                    For i = 0 To imx3.Count - 1
                        imx3(i).Enabled = True: chk3(i).Visible = False: chk3(i).Value = 0
                    Next i
                Case 4
                    For i = 0 To imx4.Count - 1
                        imx4(i).Enabled = True: chk4(i).Visible = False: chk4(i).Value = 0
                    Next i
            End Select
            mnuDownloadSels.Enabled = False
            mnuDownloadSels2.Visible = False
            lblDownload.ForeColor = vbButtonText
'            picMulti.Visible = False
            bDMode = False
            
        Case Is = False ''START DOWNLOAD MODE''
            mnuDownloadMode.Checked = True
            Select Case iModeTab
                Case 1
                    For i = 0 To imx1.Count - 1
                        imx1(i).Enabled = False: chk1(i).Visible = True: chk1(i).Value = 0: chk1(i).ZOrder
                    Next i
                Case 2
                    For i = 0 To imx2.Count - 1
                        imx2(i).Enabled = False: chk2(i).Visible = True: chk2(i).Value = 0: chk2(i).ZOrder
                    Next i
                Case 3
                    For i = 0 To imx3.Count - 1
                        If imx3(i).Visible Then
                            imx3(i).Enabled = False: chk3(i).Visible = True: chk3(i).Value = 0: chk3(i).ZOrder
                        End If
                    Next i
                Case 4
                    For i = 0 To imx4.Count - 1
                        imx4(i).Enabled = False: chk4(i).Visible = True: chk4(i).Value = 0: chk4(i).ZOrder
                    Next i
            End Select
            mnuDownloadSels.Enabled = True
            mnuDownloadSels2.Visible = True
            mnuEmailSels2.Visible = False
            lblDownload.ForeColor = vbRed
            bDMode = True
            
'            lblMess.Caption = "Checkboxes have been " & vbNewLine & _
'                        "placed adjacent to each of the Thumbnail " & _
'                        "Images in this page." & vbNewLine & vbNewLine & _
'                        "Please, check the Images you would like to Download, " & _
'                        "then return to the Download Popup to execute."
'            picMulti.Height = (lblMess.Top * 2) + lblMess.Height
'            picMulti.Visible = True
'            picMulti.SetFocus
            
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
'''                "Select Folder to download Graphic File(s) into:", _
'''                BIF_RETURNONLYFSDIRS)
'''
'''    If shlFolder Is Nothing Then
    
    If pDownloadPath = "" Then
        Exit Sub
    Else
        Screen.MousePointer = 11
        
        On Error GoTo BadFile
        sFolder = pDownloadPath '' shlFolder.Items.Item.Path
        
'''        If UCase(Left(sFolder, 1)) = "U" Or UCase(Left(sFolder, 1)) = "V" Then
'''            Screen.MousePointer = 0
'''            MsgBox "You do not have rights to download files onto one of " & _
'''                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
'''                        "Please, select another location.", vbExclamation, "Invalid Location..."
'''            Exit Sub
'''        End If
        
        If iAppConn = 1 And UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        ElseIf iAppConn = 2 And UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto this Thin-Client drive." & _
                        vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        On Error GoTo ErrorTrap
        Select Case iModeTab
            Case 0
                For i = 0 To chk0.Count - 1
                    If chk0(i).Value = 1 Then
                        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                                    "FROM " & GFXMas & " " & _
                                    "WHERE GID = " & lbl0(i).Tag
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
            Case 1
                For i = 0 To chk1.Count - 1
                    If chk1(i).Value = 1 Then
                        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                                    "FROM " & GFXMas & " " & _
                                    "WHERE GID = " & lbl1(i).Tag
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
            Case 2
                For i = 0 To chk2.Count - 1
                    If chk2(i).Value = 1 Then
                        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                                    "FROM " & GFXMas & " " & _
                                    "WHERE GID = " & lbl2(i).Tag
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
            Case 3
                For i = 0 To chk3.Count - 1
                    If chk3(i).Value = 1 Then
                        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                                    "FROM " & GFXMas & " " & _
                                    "WHERE GID = " & lbl3(i).Tag
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
            Case 4
                For i = 0 To chk4.Count - 1
                    If chk4(i).Value = 1 Then
                        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
                                    "FROM " & GFXMas & " " & _
                                    "WHERE GID = " & flxApprove.TextMatrix(i + 1, 0)
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
        End Select
        
        Call Me.ClearModes(iModeTab)
        
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

Private Sub mnuEmailMode_Click()
    Dim i As Integer
    
    ''MAKE CERTAIN NOT IN DOWNLOAD MODE''
    mnuDownloadMode.Checked = 0
    lblDownload.ForeColor = vbButtonText
    bDMode = False
    
    Select Case mnuEmailMode.Checked
        Case Is = True ''END EMAIL MODE''
            mnuEmailMode.Checked = False
            Select Case iModeTab
                Case 0
                    For i = 0 To imx0.Count - 1
                        imx0(i).Enabled = True: chk0(i).Visible = False: chk0(i).Value = 0
                    Next i
                Case 1
                    For i = 0 To imx1.Count - 1
                        imx1(i).Enabled = True: chk1(i).Visible = False: chk1(i).Value = 0
                    Next i
                Case 2
                    For i = 0 To imx2.Count - 1
                        imx2(i).Enabled = True: chk2(i).Visible = False: chk2(i).Value = 0
                    Next i
                Case 3
                    For i = 0 To imx3.Count - 1
                        imx3(i).Enabled = True: chk3(i).Visible = False: chk3(i).Value = 0
                    Next i
                Case 4
                    For i = 0 To imx4.Count - 1
                        imx4(i).Enabled = True: chk4(i).Visible = False: chk4(i).Value = 0
                    Next i
            End Select
            mnuEmailSels.Enabled = False
            mnuEmailSels2.Visible = False
            lblEmail.ForeColor = vbButtonText
            bEMode = False
            
        Case Is = False ''START EMAIL MODE''
            mnuEmailMode.Checked = True
            Select Case iModeTab
                Case 0
                    For i = 0 To imx0.Count - 1
                        If imx0(i).Visible Then
                            imx0(i).Enabled = False: chk0(i).Visible = True: chk0(i).ZOrder
                        End If
                    Next i
                Case 1
                    For i = 0 To imx1.Count - 1
                        If imx1(i).Visible Then
                            imx1(i).Enabled = False: chk1(i).Visible = True: chk1(i).ZOrder
                        End If
                    Next i
                Case 2
                    For i = 0 To imx2.Count - 1
                        If imx2(i).Visible Then
                            imx2(i).Enabled = False: chk2(i).Visible = True: chk2(i).ZOrder
                        End If
                    Next i
                Case 3
                    For i = 0 To imx3.Count - 1
                        If imx3(i).Visible Then
                            imx3(i).Enabled = False: chk3(i).Visible = True: chk3(i).ZOrder
                        End If
                    Next i
                Case 4
                    For i = 0 To imx4.Count - 1
                        If imx4(i).Visible Then
                            imx4(i).Enabled = False: chk4(i).Visible = True: chk4(i).ZOrder
                        End If
                    Next i
            End Select
            mnuEmailSels.Enabled = True
            mnuEmailSels2.Visible = True
            mnuDownloadSels2.Visible = False
            lblEmail.ForeColor = vbRed
            bEMode = True
            
'            lblMess.Caption = "Checkboxes have been " & vbNewLine & _
'                        "placed adjacent to each of the Thumbnail " & _
'                        "Images in this page." & vbNewLine & vbNewLine & _
'                        "Please, check the Images you would like to Email Copies of, " & _
'                        "then return to the Email Copy Popup to execute."
'            picMulti.Height = (lblMess.Top * 2) + lblMess.Height
'            picMulti.Visible = True
'            picMulti.SetFocus
                        
'''            MsgBox "Checkboxes have been placed adjacent to each of the Thumbnail " & _
'''                        "Images in this page." & vbNewLine & vbNewLine & _
'''                        "Please, check the Images you would like to Email Copies of, " & _
'''                        "then return to the Email Copy Popup to execute.", _
'''                        vbInformation, "Entering Email Copy Mode..."
    End Select

End Sub


Private Sub mnuEmailSel_Click()
    frmEmailFile.PassFrom = Me.Name & "-single"
    frmEmailFile.PassTAB = iModeTab
    frmEmailFile.PassBCC = tBCC
    frmEmailFile.PassFBCN = tFBCN
    Select Case iModeTab
        Case 4
            frmEmailFile.PassHDR = ""
        Case Else
            frmEmailFile.PassHDR = Me.GetHeader(tvwGraphics(iModeTab).SelectedItem) '' cboCUNO(iModeTab).Text & " Images"
    End Select
    frmEmailFile.Show 1, Me
End Sub

Private Sub mnuEmailSels_Click()
    frmEmailFile.PassFrom = Me.Name & "-multi"
    frmEmailFile.PassTAB = iModeTab
    frmEmailFile.PassBCC = fBCC(iModeTab)
    frmEmailFile.PassFBCN = fFBCN(iModeTab)
    Select Case iModeTab
        Case 4
            frmEmailFile.PassHDR = UCase(cboCUNO(4).Text) & " Approval Interface Graphics"
        Case Else
            frmEmailFile.PassHDR = Me.GetHeader(tvwGraphics(iModeTab).SelectedItem) '' cboCUNO(iModeTab).Text & " Images"
    End Select
    frmEmailFile.Show 1, Me
    Call ClearModes(iModeTab)
End Sub

Private Sub mnuEmailSels2_Click()
    Call mnuEmailSels_Click
End Sub


Private Sub mnuGfxApproval_Click()
    Dim i As Integer
    Select Case UCase(lblStat(CurrIndex))
        Case "INTERNAL": i = 0
        Case "CLIENT DRAFT": i = 1
        Case "APPROVED": i = 2
        Case "RETURNED": i = 4
        Case Else: i = 3
    End Select
    StatusGID = lGID
    
    frmGfxApprove.PassX = 0
    frmGfxApprove.PassY = 0
    frmGfxApprove.PassHDR = "Graphic Approval for '" & flxApprove.TextMatrix(CurrIndex + 1, 3) & "'"
    frmGfxApprove.PassVal = i
    frmGfxApprove.PassBCC = CLng(fBCC(4))
    frmGfxApprove.PassFBCN = fFBCN(4)
    frmGfxApprove.PassType = "VWR"
    frmGfxApprove.Show 1, Me
End Sub

Private Sub mnuGFXData_Click()
    Dim strSelect As String
    
    strSelect = "SELECT GM.* " & _
                "FROM " & GFXMas & " GM, " & GFXShow & " GS " & _
                "WHERE GS.SHOW_ID = " & imx1(CurrIndex).Tag & " " & _
                "AND GS.GID = GM.GID"
    Call GetGFXData(strSelect, "msgbox")
End Sub

Private Sub mnuGFXData2_Click()
    Dim strSelect As String
    
    strSelect = "SELECT * " & _
                "FROM " & GFXMas & " " & _
                "WHERE GID = " & lGID
    Call GetGFXData(strSelect, "msgbox")
End Sub

Private Sub mnuGPrint_Click(Index As Integer)
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
    
    If picPDF.Visible Then
        xpdf1.printWithDialog
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    cdl1.Flags = cdlPDPrintSetup
    cdl1.ShowPrinter
    
    Select Case Index
        Case 0
'''            imgSize.Picture = LoadPicture(CurrFile)
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
        Case 1
            pAspect = picJPG.Width / picJPG.Height
            If pAspect >= 1 Then '''LANDSCAPE'''
'''                cdl1.Orientation = cdlLandscape
                Printer.Orientation = 2
                If pAspect > (10 / 6.75) Then
                    pScaleX = 10 * 1440
                    pScaleY = picJPG.Height / picJPG.Width * pScaleX
                Else
                    pScaleY = 6.75 * 1440
                    pScaleX = picJPG.Width / picJPG.Height * pScaleY
                End If
                lXStart = (Printer.ScaleWidth - pScaleX) / 2
                lYStart = ((Printer.ScaleHeight - pScaleY) / 2) + (0.5 * 1440)
            Else '''PORTRAIT'''
'''                cdl1.Orientation = cdlPortrait
                Printer.Orientation = 1
                If pAspect < (7 / 9.25) Then
                    pScaleY = 9.25 * 1440
                    pScaleX = picJPG.Width / picJPG.Height * pScaleY
                Else
                    pScaleX = 7 * 1440
                    pScaleY = picJPG.Height / picJPG.Width * pScaleX
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
            Printer.PaintPicture picJPG.Image, lXStart, lYStart, pScaleX, pScaleY
    End Select
        
    
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
        If Err Then
            Screen.MousePointer = 0
            MsgBox sMsg, vbExclamation, "Printer Error..."
            Err = 0
        Else
            Screen.MousePointer = 0
            MsgBox "Image printed to default printer:  " & Printer.DeviceName, vbInformation, "Print Sent..."
        End If
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

Private Sub mnuGRedClear_Click()
'    Call Me.ClearUndo(0)
    
    If picJPG.Visible Then
'''        If lblRedNote.BorderStyle = 1 Then
'''            lblRedNote.BorderStyle = 0
'''            lblRedNote.Visible = False
'''            txt1.Text = ""
'''            lblRedNote.Caption = ""
'''        End If
        If RedFile <> "" Then
            Call LoadThePicture(RedFile, True)
            Call AddToUndo("picjpg", -1)
        Else
            picJPG.PaintPicture imgSize.Picture, 0, 0, picJPG.Width, picJPG.Height
            Call AddToUndo("picjpg", -1)
            Call SetImageState
'''            Call LoadThePicture(CurrFile)
        End If
    ElseIf picRed.Visible Then
        bRedMode = False
        Call SetupPDFRed(iRedMode, RedFile)
    End If
    If SaveMess <> "" Then lblStatus = SaveMess
    SaveMess = ""
    bRedded = False
    If Dir(RedFile, vbNormal) <> "" And RedFile <> "" Then mnuGRedLoad.Enabled = True Else mnuGRedLoad.Enabled = False
    mnuGRedSave.Enabled = False
    mnuGRedClear.Enabled = False
    mnuGRedDelete.Enabled = False
    imgUtility(0).Enabled = False
    imgUtility(0).Picture = imlRedMode.ListImages(11).Picture
    Call UpdateSCD(0, 0, Abs(CInt(imgUtility(2).Enabled)))
'    ClearUndo (0)
    
End Sub

Private Sub mnuGRedDelete_Click()
    Dim strDelete As String
    Dim RetVal As VbMsgBoxResult
    RetVal = MsgBox("Are you certain you want to delete this Redline File?", _
                vbExclamation + vbYesNoCancel, "About to Delete Redline File...")
    If RetVal = vbYes Then
'''        strDelete = "DELETE FROM " & GFXRed & " " & _
'''                    "WHERE REF_ID = " & lRedID
        strDelete = "DELETE FROM " & GFXRed & " " & _
                    "WHERE REF_ID = " & lRedID & " " & _
                    "AND PAGE_ID = " & iPDFPage & " " & _
                    "AND RED_STATUS > 0"
        Conn.Execute (strDelete)
        
        bRedSaved = False
        bRedded = False
        TextMode = False: RedMode = False
        If Dir(RedFile, vbNormal) <> "" Then Kill RedFile
        RedFile = ""
        
        If picRed.Visible Then
            ''THIS IS A PDF''
            
            lblRedline.Caption = Me.GetPDFRedCount(lRedID)
'            Call AddToUndo("PICRED", -1)
            Call imgRedReload_Click
        Else
            ''THIS IS A JPG''
            
            
            lblRedline.Caption = ""
            Call LoadThePicture(CurrFile, False)
            
        End If
'        Call mnuGRedEnd_Click
        
        Call UpdateSCD(0, 0, 0)
'        mnuGRedLoad.Enabled = False
'        mnuGRedClear.Enabled = False
'        mnuGRedDelete.Enabled = False
'        mnuGRedSave.Enabled = True
        
        
        
''        Call LoadThePicture(CurrFile)
        If SaveMess <> "" Then lblStatus = SaveMess
        SaveMess = ""
'        lblGraphic = Trim(Left(lblGraphic, Len(lblGraphic) - 26))
'''''        Kill RedFile
        
        
'''        Call UpdateSCD(0, 0, 1)
        Call ClearUndo(0)
        
'''        If lblRedNote.BorderStyle = 1 Then
'''            lblRedNote.BorderStyle = 0
'''            lblRedNote.Visible = False
'''            txt1.Text = ""
'''            lblRedNote.Caption = ""
'''        End If
    End If
End Sub

Private Sub mnuGRedEnd_Click()
    
    RedMode = False
    TextMode = False
    bRedMode = False: If Not bHideRN Then Call RedNoteVis(bRedMode)
    mnuGRedSketch.Checked = False
    mnuGRedText.Checked = False
    If picJPG.Visible Then
        picJPG.MousePointer = 0
        If lblRedNote.BorderStyle = 1 Then
            lblRedNote.BorderStyle = 0
            lblRedNote.Visible = False
            txt1.Text = ""
            lblRedNote.Caption = ""
        End If
        picRedTools.Visible = False
        picJPGTools.Visible = True
        Call imgNav_Click(-1)
    ElseIf picPDF.Visible Then
        If bRedded Then
            Select Case ShallWeSave
                Case 0 ''NO''
                Case 1 ''YES''
                Case 2 ''CANCEL''
                    Exit Sub
            End Select
        End If
        picRed.Visible = False
        picRed.Picture = LoadPicture()
        xpdf1.Visible = True
        picRedTools.Visible = False
        picPDFTools.Visible = True
    
    End If
End Sub

Private Sub mnuGRedLoad_Click()
    If Dir(RedFile, vbNormal) = "" Or RedFile = "" Then
        mnuGRedLoad.Enabled = False
        MsgBox "No Redline File found", vbInformation, "Sorry..."
        Exit Sub
    End If
    
    mnuGRedClear.Enabled = True
    mnuGRedDelete.Enabled = True
    mnuGRedSave.Enabled = True
    Call UpdateSCD(0, 0, 1)
    Call ClearUndo(0)
    
    If picJPG.Visible Then
        Call LoadThePicture(RedFile, True)
        imgUtility(2).Enabled = True
'''        Call AddToUndo("picjpg", -1)
        bRedMode = True: If Not bHideRN Then Call RedNoteVis(bRedMode)
        SaveMess = lblStatus.Caption
        If RedMess <> "" Then lblStatus.Caption = RedMess
        
'''        Call AddToUndo("picjpg", iUndoIndex)
        
        Call mnuGRedSketch_Click
        
''''        Set picCurrentRed = picJPG
''''
''''        picRedTools.Visible = True
''''
''''        imgRedMode(0).Picture = imlRedMode.ListImages(1).Picture
''''        imgRedMode(1).Picture = imlRedMode.ListImages(4).Picture
        
'        If lblRedNote.BorderStyle = 1 Then
'            lblRedNote.BorderStyle = 0
'            lblRedNote.Visible = False
'            txt1.Text = ""
'            lblRedNote.Caption = ""
'        End If
    ElseIf picPDF.Visible Then
        bRedMode = False
        Call SetupPDFRed(1, RedFile)
        Call UpdateSCD(0, 0, 1)
        Call ClearUndo(0)
    End If
End Sub

Private Sub mnuGRedSave_Click()
    Dim strInsert As String, strUpdate As String
    
    Screen.MousePointer = 11
    On Error Resume Next
    Conn.BeginTrans
    If Not CheckForRed(lRefID, iPDFPage) Then
        '///// NEW REDLINE FILE \\\\\
        RedFile = sGPath & RedName
'''        strInsert = "INSERT INTO " & GFXRed & " " & _
'''                    "(REF_ID, REDPATH, " & _
'''                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
'''                    "VALUES " & _
'''                    "(" & lRefID & ", '" & RedFile & "', " & _
'''                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"

        strInsert = "INSERT INTO " & GFXRed & " " & _
                    "(REF_ID, PAGE_ID, REDPATH, RED_STATUS, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lRefID & ", " & iPDFPage & ", '" & RedFile & "', 1, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
    Else
'''        strUpdate = "UPDATE " & GFXRed & " " & _
'''                    "SET UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
'''                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
'''                    "WHERE REF_ID = " & lRefID
        RedFile = sGPath & RedName
        strUpdate = "UPDATE " & GFXRed & " " & _
                    "SET REDPATH = '" & RedFile & "', " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE REF_ID = " & lRefID & " " & _
                    "AND PAGE_ID = " & iPDFPage
        Conn.Execute (strUpdate)
    End If
    
    If picJPG.Visible Then
        Call BurnishIt(picJPG)
        SavePicture picJPG.Image, RedFile
    ElseIf picPDF.Visible Then
        Call BurnishIt(picRed)
        SavePicture picRed.Image, RedFile
    End If
    
    If Err = 0 Then
        Conn.CommitTrans
        mnuGRedDelete.Enabled = True
        
        If picRed.Visible Then
            Call UpdateSCD(0, 0, 1)
            lblRedline.Caption = Me.GetPDFRedCount(lRedID)
            Call AddToUndo("PICRED", -1)
        Else
            lblRedline.Caption = "[A Redline File Exists]"
            Call AddToUndo("PICJPG", -1)
        End If
        lblRedline.Visible = True
        Call UpdateSCD(0, 0, 1)
'        Call ClearUndo(0)
'''        If InStr(1, lblGraphic.Caption, "       [A Redline File Exists]") = 0 Then
'''            lblGraphic = lblGraphic & "       [A Redline File Exists]"
'''        End If
        bRedSaved = True
        bRedded = False
        redBCC = tBCC: redSHCD = tSHCD
    Else
        Conn.RollbackTrans
        MsgBox "Error:  " & Err.Description, vbCritical, "Error Encountered during Save..."
        Err.Clear
        bRedded = False
        redBCC = "": redSHCD = 0
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub mnuGRedSketch_Click()
    Dim Resp As VbMsgBoxResult
    
    ClearChecks
    iRedMode = 1
'''    If lblRedNote.BorderStyle = 1 Then
'''        lblRedNote.BorderStyle = 0
'''        lblRedNote.Visible = False
'''        txt1.Text = ""
'''        lblRedNote.Caption = ""
'''    End If
    mnuGRedSketch.Checked = True
'''    mnuVAnnotation.Checked = True
    mnuGRedSave.Enabled = True
    mnuGRedClear.Enabled = True
    
    RedMode = True
    TextMode = False
    If picJPG.Visible Then
'''        picJPG.MouseIcon = imgIcon(0).Picture
        If Not bRedMode Then
            ''CHECK FOR EXISTING REDLINE''
            If RedFile <> "" Then
                ''FOUND A REDLINE''
                Resp = MsgBox("There is a current Redline File.  Do you want to load it?", _
                            vbQuestion + vbYesNo, "Active Redline File...")
                If Resp = vbYes Then
                    Call LoadThePicture(RedFile, True)
                    Call UpdateSCD(0, 0, 1)
                Else
                    Call UpdateSCD(0, 0, 0)
                End If
'            ElseIf RedFile <> "" Then
'                Call UpdateSCD(0, 0, 1)
            Else
                Call UpdateSCD(0, 0, 0)
            End If
'            Call ClearUndo(0)
            Call AddToUndo("picjpg", -1)
        End If
        
        picJPG.MousePointer = 99
        Set picCurrentRed = picJPG
        picRedTools.Visible = True
        picJPGTools.Visible = False
        bRedMode = True: If Not bHideRN Then Call RedNoteVis(bRedMode)
        imgRedMode(0).Picture = imlRedMode.ListImages(1).Picture
        imgRedMode(1).Picture = imlRedMode.ListImages(4).Picture
''        bRedLine = True
''        bRedText = False
            
    ElseIf picPDF.Visible Then
        lblEsc.Visible = False
        Call SetupPDFRed(iRedMode, RedFile) '' RedName)
    End If
End Sub

Private Sub mnuGRedText_Click()
    Dim Resp As VbMsgBoxResult
    
    mnuGRedSketch.Checked = False
    mnuGRedText.Checked = True
    
    TextMode = True
    RedMode = False
    iRedMode = 2
    If picJPG.Visible Then
        If Not bRedMode Then
            ''CHECK FOR EXISTING REDLINE''
            If RedFile <> "" Then
                ''FOUND A REDLINE''
                Resp = MsgBox("There is a current Redline File.  Do you want to load it?", _
                            vbQuestion + vbYesNo, "Active Redline File...")
                If Resp = vbYes Then
                    Call LoadThePicture(RedFile, True)
                    Call UpdateSCD(0, 0, 1)
                Else
                    Call UpdateSCD(0, 0, 0)
                End If
'            ElseIf RedFile <> "" Then
'                Call UpdateSCD(0, 0, 1)
            Else
                Call UpdateSCD(0, 0, 0)
            End If
'            Call ClearUndo(0)
            Call AddToUndo("picjpg", -1)
        End If
        
        bRedMode = True: If Not bHideRN Then Call RedNoteVis(bRedMode)
        If lblRed(0).Container.Name <> picJPG.Name Then
            Call ResetLBLREDContainer(picJPG)
        End If
        picJPG.MousePointer = 3 '' 99
        picRedTools.Visible = True
        picJPGTools.Visible = False
        imgRedMode(1).Picture = imlRedMode.ListImages(3).Picture
        imgRedMode(0).Picture = imlRedMode.ListImages(2).Picture
            
    Else
        Call SetupPDFRed(iRedMode, RedFile)
    End If
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub mnuMaxGraphic_Click()
    Dim rFactor As Double
    
    If picJPG.Visible Then
        With picJPG
            .Visible = False
            .Width = rMX
            .Height = rMY
            .Top = dMTop
            .Left = dMLeft
            .PaintPicture imgSize.Picture, 0, 0, .Width, .Height
            .Visible = True
            rFactor = .Width / imgSize.Width
            lblSize.Caption = CLng(rFactor * 100) & "%"
            If imgSize.Width > picJPG.Width Then
                Call ResetJPGZoom(0, 2, 1)
            Else
                Call ResetJPGZoom(1, 2, 0)
            End If
        End With
    ElseIf picPDF.Visible Then
        SetZoomMode (1)
        cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
    End If
    mnuResizeGraphic.Enabled = True
    mnuMaxGraphic.Enabled = False
''''''''    lblResize.Caption = "Resize"
    
    iImageState = 1
End Sub

Private Sub mnuOpen_Click()
    Call PrepareToOpen(CurrIndex)
End Sub



Private Sub mnuOptClearAll_Click()
    Dim i As Integer
    Select Case iModeTab
        Case 0
            For i = 0 To chk0.Count - 1
                If chk0(i).Visible Then chk0(i).Value = 0
            Next i
        Case 1
            For i = 0 To chk1.Count - 1
                If chk1(i).Visible Then chk1(i).Value = 0
            Next i
        Case 2
            For i = 0 To chk2.Count - 1
                If chk2(i).Visible Then chk2(i).Value = 0
            Next i
        Case 3
            For i = 0 To chk3.Count - 1
                If chk3(i).Visible Then chk3(i).Value = 0
            Next i
        Case 4
            For i = 0 To chk4.Count - 1
                If chk4(i).Visible Then chk4(i).Value = 0
            Next i
    End Select
End Sub

Private Sub mnuOptSelAll_Click()
    Dim i As Integer
    Select Case iModeTab
        Case 0
            For i = 0 To chk0.Count - 1
                If chk0(i).Visible Then chk0(i).Value = 1
            Next i
        Case 1
            For i = 0 To chk1.Count - 1
                If chk1(i).Visible Then chk1(i).Value = 1
            Next i
        Case 2
            For i = 0 To chk2.Count - 1
                If chk2(i).Visible Then chk2(i).Value = 1
            Next i
        Case 3
            For i = 0 To chk3.Count - 1
                If chk3(i).Visible Then chk3(i).Value = 1
            Next i
        Case 4
            For i = 0 To chk4.Count - 1
                If chk4(i).Visible Then chk4(i).Value = 1
            Next i
    End Select
End Sub

Private Sub mnuPartData_Click()
    Dim strSelect As String
    Dim bNodes As Boolean
'''    frmLogistics.PassEID = Mid(tvwGraphics(2).SelectedItem.key, 2)
'''    frmLogistics.PassElem = tvwGraphics(2).SelectedItem.Text
'''    frmLogistics.PassFrom = "GFX"
'''    frmLogistics.Show 1, Me
    
    If tvwGraphics(2).SelectedItem.Children = 0 Then bNodes = True Else bNodes = False
    
    strSelect = "SELECT PARTID, PARTDESC, PKGTYPE, " & _
                "NVL(WIDTH, 0) AS WIDTH, NVL(HEIGHT, 0) AS HEIGHT, " & _
                "NVL(LENGTH, 0) AS LENGTH, SIZEUNIT, " & _
                "NVL(WEIGHT, 0) AS WEIGHT, WTUNIT, " & _
                "(FABLOC||TO_CHAR(YRBUILT, 'YY')||'-'||PNUMBER)PNUM " & _
                "FROM IGLPROD.IGL_PART " & _
                "WHERE ELTID = " & Mid(tvwGraphics(2).SelectedItem.Key, 2) & " " & _
                "AND TSTATUS > 0 " & _
                "ORDER BY PARTDESC, PNUM"
    Call GetPartNodes(CLng(Mid(tvwGraphics(2).SelectedItem.Key, 2)), _
                tvwGraphics(2).SelectedItem.Text, strSelect, bNodes)
    tvwGraphics(2).SelectedItem.Expanded = True
End Sub

Private Sub mnuResizeGraphic_Click()
    Dim rFactor As Double
    
    If picJPG.Visible Then
        With picJPG
            .Visible = False
            .Width = rSX
            .Height = rSY
            .Top = dSTop ''' dGTop + ((rY - rYO) / 2): .Left = dGLeft + ((rX - rXO) / 2)
            .Left = dSLeft
    '''        .Left = (maxX - rXO) / 2
            .PaintPicture imgSize.Picture, 0, 0, .Width, .Height
            .Visible = True
''''''''            rFactor = .Width / imgSize.Width
''''''''            lblSize.Caption = CLng(rFactor * 100) & "%"
            Call ResetJPGZoom(2, 1, 0)
        End With
        
    ElseIf picPDF.Visible Then
        xpdf1.Zoom = 100
        cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
        SetZoomMode (3)
        
'''        bZWindow = False
'''        bPan = True
'''        Xpdf1.enableMouseEvents = CBool(1)
'''        Xpdf1.enableSelect = False
'''        Xpdf1.mouseCursor = imgCur(1).Picture
        
        
        
    End If
    mnuResizeGraphic.Enabled = False
    mnuMaxGraphic.Enabled = True
''''''''    lblResize.Caption = "Maximize"
    
    iImageState = 0
End Sub

Private Sub mnuSelByImage_Click(Index As Integer)
    Dim i As Integer
    
    If iView = 0 Then
        Call ResetControls(1)
        '///// RESET SELECTIONS \\\\\
        For i = 0 To imx1.Count - 1
            imx1(i).Visible = False
            lbl1(i).Visible = False
        Next i
        iView = 1
        If cboSHCD <> "" Then cboSHCD.Text = cboSHCD.Text
        If cboCUNO(1) <> "" Then cboCUNO(1).Text = cboCUNO(1).Text
        If cboCUNO(2) <> "" Then cboCUNO(2).Text = cboCUNO(2).Text
        If cboCUNO(3) <> "" Then cboCUNO(3).Text = cboCUNO(3).Text
    End If
    iRes = Index
    
    mnuSelByName.Checked = False
    For i = 2 To 3
        If i = iRes Then mnuSelByImage(i).Checked = True Else mnuSelByImage(i).Checked = False
    Next i
    
End Sub

Private Sub mnuSelByName_Click()
    If iView = 1 Then
        Call ResetControls(0)
        iView = 0
        mnuSelByImage(2).Checked = False
        mnuSelByImage(3).Checked = False
        mnuSelByName.Checked = True
        If cboSHCD <> "" Then cboSHCD.Text = cboSHCD.Text
        If cboCUNO(1) <> "" Then cboCUNO(1).Text = cboCUNO(1).Text
        If cboCUNO(2) <> "" Then cboCUNO(2).Text = cboCUNO(2).Text
        If cboCUNO(3) <> "" Then cboCUNO(3).Text = cboCUNO(3).Text
    End If
End Sub

Private Sub mnuSendALink_Click()
    frmSendALink.PassBCC = CLng(tBCC)
    frmSendALink.PassFrom = "GH"
    frmSendALink.PassGID = lGID
    frmSendALink.PassSub = "AnnoLink:  " & _
            Trim(lblWelcome.Caption) & "  (" & Trim(lblGraphic.Caption) & ")"
    frmSendALink.Show 1, Me

End Sub

Private Sub mnuStatusCancel_Click()
    Dim i As Integer, iCnt As Integer
    Dim iNewStatus As Integer
    Dim strUpdate As String, strSelect As String, sNodeKey As String
    Dim rst As ADODB.Recordset
    Dim GIDList As String
    Dim Resp As VbMsgBoxResult
    
    Select Case iCurrStatus
        Case 10: iNewStatus = 2
        Case 20: iNewStatus = 3
        Case 30: iNewStatus = 4
    End Select
    
'''    Conn.BeginTrans
    On Error GoTo ErrorTrap
    Select Case mnuName(0).Tag
        Case "ONE"
            Resp = MsgBox("Are you certain you want to Cancel this file" & _
                        " ?", vbExclamation + vbYesNo, "Cancel Verification...")
            If Resp = vbYes Then
                strUpdate = "UPDATE " & GFXMas & " " & _
                            "SET GSTATUS = " & iNewStatus & ", " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                            "WHERE GID = " & StatusGID
                Conn.Execute (strUpdate)
                StatusGID = 0
            Else
                GoTo CancelIt
            End If
        
        Case "VIEW"
            GIDList = "": iCnt = 0
            For i = 0 To imx3.Count - 1
                If imx3(i).Visible = True Then
                    If GIDList = "" Then GIDList = imx3(i).Tag Else GIDList = GIDList & ", " & imx3(i).Tag
                    iCnt = iCnt + 1
                End If
            Next i
            Select Case iCnt
                Case 1
                    Resp = MsgBox("Are you certain you want to Cancel the one " & _
                            tvwGraphics(3).SelectedItem.Parent.Parent.Text & " File in the Current View", _
                            vbExclamation + vbYesNo, "Cancel Verification...")
                Case Else
                    Resp = MsgBox("Are you certain you want to Cancel the ( " & iCnt & " ) " & _
                            tvwGraphics(3).SelectedItem.Parent.Parent.Text & " Files in the Current View", _
                            vbExclamation + vbYesNo, "Cancel Verification...")
            End Select
            If Resp = vbYes Then
'''                GIDList = ""
'''                For i = 0 To imx3.Count - 1
'''                    If imx3(i).Visible = True Then
'''                        If GIDList = "" Then GIDList = imx3(i).Tag Else GIDList = GIDList & ", " & imx3(i).Tag
'''                    End If
'''                Next i
                
                strUpdate = "UPDATE " & GFXMas & " " & _
                            "SET GSTATUS = " & iNewStatus & ", " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                            "WHERE GID IN (" & GIDList & ")"
                Conn.Execute (strUpdate)
            Else
                GoTo CancelIt
            End If
            
        Case "ALL"
            Resp = MsgBox("Are you certain you want to Cancel all ( " & iGFXCount(3) & " ) files in the '" & _
                        tvwGraphics(3).SelectedItem.Parent.Text & " - " & tvwGraphics(3).SelectedItem.Text & "' Folder?", _
                        vbExclamation + vbYesNo, "Cancel Verification...")
            If Resp = vbYes Then
                GIDList = ""
                sNodeKey = tvwGraphics(3).SelectedItem.Key
                strSelect = "SELECT GID FROM " & GFXMas & " " & _
                            "WHERE AN8_CUNO = " & CLng(fBCC(4)) & " " & _
                            "AND GTYPE = " & Right(sNodeKey, 1) & " " & _
                            "AND GSTATUS = " & CInt(Mid(sNodeKey, 2, 2)) & " " & _
                            "AND TRUNC(ADDDTTM) BETWEEN TO_DATE('" & Mid(sNodeKey, 5, 2) & "/01/" & Mid(sNodeKey, 7, 4) & "', 'MM/DD/YYYY') " & _
                            "AND TO_DATE('" & Format(DateAdd("M", 1, DateValue(Mid(sNodeKey, 5, 2) & "/01/" & Mid(sNodeKey, 7, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY')"
                Set rst = Conn.Execute(strSelect)
                Do While Not rst.EOF
                    If GIDList = "" Then GIDList = rst.Fields("GID") Else GIDList = GIDList & ", " & rst.Fields("GID")
                    rst.MoveNext
                Loop
                rst.Close: Set rst = Nothing
                
                strUpdate = "UPDATE " & GFXMas & " " & _
                            "SET GSTATUS = " & iNewStatus & ", " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                            "WHERE GID IN (" & GIDList & ")"
                Conn.Execute (strUpdate)
            Else
                GoTo CancelIt
            End If
    End Select
'''    Conn.CommitTrans
    
    cboCUNO(4).Text = cboCUNO(4).Text
'''    Call tvwGraphics_NodeClick(3, tvwGraphics(3).SelectedItem)
CancelIt:
Exit Sub
ErrorTrap:
'''    Conn.RollbackTrans
    MsgBox "Error Encountered during Cancellation." & vbNewLine & vbNewLine & _
                "Error:  " & Err.Description, vbCritical, "Status Change Aborted..."
    Err.Clear
End Sub

Private Sub mnuStatusReset_Click(Index As Integer)
    Dim i As Integer
    Dim iNewStatus(0 To 2) As Integer
    Dim strUpdate As String, strSelect As String, strInsert As String
    Dim sNodeKey As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim GIDList As String
'''    Dim sComm As String
    Dim lCOMMID As Long
    Dim iErr As Integer
    
    iNewStatus(0) = 5
    iNewStatus(1) = 15
    iNewStatus(2) = 25
    
    Conn.BeginTrans
    On Error GoTo ErrorTrap
    Select Case mnuName(0).Tag
        Case "ONE"
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = " & iNewStatus(Index) & ", " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                        "WHERE GID = " & StatusGID
            Conn.Execute (strUpdate)
            
'''            ''NOTE STATUS CHANGE IN COMMENTS''
'''            Select Case Index
'''                Case 0: sComm = "Graphic Status reset to 'DRAFT' by " & LogName & "."
'''                Case 1: sComm = "Graphic Status reset to 'RELEASED' by " & LogName & "."
'''                Case 2: sComm = "Graphic 'APPROVED' by " & LogName & "."
'''            End Select
            
            iErr = InsertComment(StatusGID, Index, "")
            If iErr > 0 Then GoTo ErrorTrap
            
            StatusGID = 0
            
        Case "VIEW"
            GIDList = ""
            For i = 1 To flxApprove.Rows - 1
                If GIDList = "" Then
                    GIDList = flxApprove.TextMatrix(i, 0)
                Else
                    GIDList = GIDList & ", " & flxApprove.TextMatrix(i, 0)
                End If
                
                iErr = InsertComment(flxApprove.TextMatrix(i, 0), Index, "")
                If iErr > 0 Then GoTo ErrorTrap
            
            Next i
            
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = " & iNewStatus(Index) & ", " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                        "WHERE GID IN (" & GIDList & ")"
            Conn.Execute (strUpdate)
            
            
'''        Case "ALL"
'''            GIDList = ""
'''            sNodeKey = tvwGraphics(3).SelectedItem.key
'''            strSelect = "SELECT GID FROM " & GFXMas & " " & _
'''                        "WHERE AN8_CUNO = " & CLng(tBCC) & " " & _
'''                        "AND GTYPE = " & Right(sNodeKey, 1) & " " & _
'''                        "AND GSTATUS = " & CInt(Mid(sNodeKey, 2, 2)) & " " & _
'''                        "AND TRUNC(ADDDTTM) BETWEEN TO_DATE('" & Mid(sNodeKey, 5, 2) & "/01/" & Mid(sNodeKey, 7, 4) & "', 'MM/DD/YYYY') " & _
'''                        "AND TO_DATE('" & format(DateAdd("M", 1, DateValue(Mid(sNodeKey, 5, 2) & "/01/" & Mid(sNodeKey, 7, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY')"
'''            Set rst = Conn.Execute(strSelect)
'''            Do While Not rst.EOF
'''                If GIDList = "" Then GIDList = rst.Fields("GID") Else GIDList = GIDList & ", " & rst.Fields("GID")
'''                rst.MoveNext
'''            Loop
'''            rst.Close: Set rst = Nothing
'''
'''            strUpdate = "UPDATE " & GFXMas & " " & _
'''                        "SET GSTATUS = " & iNewStatus(Index) & ", " & _
'''                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
'''                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
'''                        "WHERE GID IN (" & GIDList & ")"
'''            Conn.Execute (strUpdate)
                    
    End Select
    Conn.CommitTrans
    
    cboCUNO(4).Text = cboCUNO(4).Text
    
'''    cmdStatusEdit_Notify.Enabled = CheckForNotify
    
'''    Call tvwGraphics_NodeClick(3, tvwGraphics(3).SelectedItem)
    
Exit Sub
ErrorTrap:
    Conn.RollbackTrans
    MsgBox "Error Encountered during Status Change." & vbNewLine & vbNewLine & _
                "Error:  " & Err.Description, vbCritical, "Status Change Aborted..."
    Err.Clear
End Sub



Private Sub mnuTextClear_Click()
    Dim i As Integer
    
    lblRed(iRed).Caption = ""
    lblRed(iRed).Visible = False
    lblEsc.Visible = False
    
    ''LOOP THRU UNDO STACK AND REMOVE IF FOUND''
    For i = lstUndo.ListCount - 1 To 0 Step -1
        If UCase(Left(lstUndo.List(i), 3)) = "LBL" And lstUndo.ItemData(i) = iRed Then
            lstUndo.RemoveItem (i)
            Exit For
        End If
    Next i
        
    If iRed = lblRed.Count - 1 And iRed > 0 Then
        Unload lblRed(iRed)
        iRed = iRed - 1
    End If
    
End Sub

Private Sub mnuTextColor_Click()
    lblRed(iRed).ForeColor = lAnnoColor
    lblRed(iRed).Tag = lAnnoColor
End Sub

Private Sub mnuTextEdit_Click()
    frmTextEditor.PassIndex = iRed
    frmTextEditor.PassText = lblRed(iRed).Caption
    frmTextEditor.Show 1, Me
End Sub

Private Sub optApproverView_Click(Index As Integer)
    Dim i As Integer
    
    If Not bResetting Then
        iApproverView = Index
'''        If cboCUNO(4).Text = "" Then
'''            MsgBox "No Client has been selected", vbCritical, "Hey..."
'''            Exit Sub
'''        End If
        
        If iApproverView < 2 Then
            lblClient.Caption = "Client:"
            cboCUNO(4).Enabled = True
            txtNoShows.Text = sNoShows
            If cboSHYR(4).ListCount > 0 Then txtNoShows.Visible = False
            
            Call cmdRefresh_Click
'''''            flxApprove.Visible = False: picOuter(4).Visible = False
'''''            Me.MousePointer = 11
'''''            If cboASHCD.Text = "" Then
'''''                Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, 0, 0)
'''''            Else
'''''                Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, fSHYR(4), fSHCD(4))
'''''            End If
'''''            flxApprove.Visible = True
'''''            picInner(4).Visible = True
'''''            picOuter(4).Visible = True: picOuter(4).Refresh
            
'            picInner(4).Refresh
'            Me.Refresh
            
'''''            For i = 0 To imx4.Count - 1
'''''                If imx4(i).Visible = True Then imx4(i).Refresh
'''''            Next i
'''''
'''''            Me.MousePointer = 0
'''            Call cboCUNO_Click(4)
        End If
    End If
End Sub

'''Private Sub optFilter_Click(Index As Integer)
'''    Dim i As Integer
'''    Dim bCheck As Boolean
'''
'''    For i = 0 To optFilter.Count - 1
'''        optFilter(i).Refresh
'''    Next i
'''
'''    Me.MousePointer = 11
'''    Select Case Index
'''        Case 0
'''            sIn = "10, 20"
'''            lblMess.Caption = "...Refreshing with All Files..."
'''            bCheck = True
'''        Case 1
'''            sIn = "10"
'''            lblMess.Caption = "...Refreshing with Draft Files Only..."
'''            cmdStatusEdit_View.Enabled = True
'''        Case 2
'''            sIn = "20"
'''            lblMess.Caption = "...Refreshing with Released Files Only..."
'''            cmdStatusEdit_View.Enabled = True
'''    End Select
'''
'''    picMess.Visible = True: picMess.Refresh
'''    flxApprove.Visible = False: picOuter(4).Visible = False
'''    Call GetApprovalGraphics(CLng(tBCC), sOrder)
'''    picMess.Visible = False
'''    flxApprove.Visible = True: picOuter(4).Visible = True
'''    If bCheck Then
'''        cmdStatusEdit_View.Enabled = True
'''        For i = 0 To flxApprove.Rows - 2
'''            If lblStat(i).Caption <> lblStat(0) Then
'''                cmdStatusEdit_View.Enabled = False
'''                Exit For
'''            End If
'''        Next i
'''    End If
'''
'''    Me.MousePointer = 0
'''End Sub

'''Private Sub optGfxApprove_Click(Index As Integer)
'''    If optGfxApprove(Index).value = True Then
'''        cmdGfxApprove.Enabled = True
'''    Else
'''        cmdGfxApprove.Enabled = False
'''    End If
'''End Sub

Private Sub optSort_Click(Index As Integer)
    Screen.MousePointer = 11
    iSSSort = Index
    Call ClearThumbnails1(0)
    Call LoadClientShows(CLng(fBCC(sst1.Tab)), fSHYR(sst1.Tab), iSSSort)
    Screen.MousePointer = 0
End Sub

'''Private Sub optUnits_Click(Index As Integer)
'''    Dim i As Integer
'''    Dim sToolTip As String
'''
'''    iUnit = Index
'''    Select Case Index
'''        Case 1
'''            optWgtUnit(1).value = True
'''            sToolTip = "Textbox accepts Fractional entry (4'-6 3/32" & Chr(34) & ")"
'''        Case 8
'''            optWgtUnit(2).value = True
'''            sToolTip = "Enter data as millimeters"
'''    End Select
'''    For i = 0 To txtDim.Count - 1
'''        txtDim(i).Enabled = True
'''        txtDim(i).ToolTipText = sToolTip
'''    Next i
'''
'''End Sub


'''Private Sub picColor_Click(Index As Integer)
'''    lAnnoColor = QBColor(Index)
'''    picCurrentRed.ForeColor = lAnnoColor
'''    shpHL.Left = picColor(Index).Left - 30
'''End Sub

'''Private Sub pdfGraphic_GotFocus()
'''    Debug.Print "PDF GotFocus"
'''End Sub

'''Private Sub pdfGraphic_Validate(Cancel As Boolean)
'''    Debug.Print "Val PDF Control"
'''End Sub





''''Private Sub picExpand_Click(Index As Integer)
''''    picExpand(Index).Visible = False
''''    picExpand(Abs(Index - 1)).Visible = True
''''    Select Case Index
''''        Case 0: picXD.Visible = True
''''        Case 1: picXD.Visible = False
''''    End Select
''''End Sub



Private Sub picJPG_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - x0, Y - y0
End Sub

'''Private Sub mnuSelRes_Click(Index As Integer)
'''    Dim i As Integer
'''    If iView = 0 Then Call ResetControls(1)
'''    iRes = Index
'''
'''    mnuSelByName.Checked = False
'''    For i = 1 To 3
'''        If i = iRes Then mnuSelRes(i).Checked = True Else mnuSelRes(i).Checked = False
'''    Next i
'''
'''    Select Case Index
'''        Case 1
'''            iRows = 3
'''        Case 2, 3
'''            iRows = 4
'''    End Select
'''    imageY = (picInner(1).Height - hsc1(1).Height - 240 - 900) / iRows
'''    spaceY = imageY + 300
'''    imageX = (imageY / 3) * 4
'''    spaceX = imageX + 240
'''    iView = 1
'''    For i = 0 To imx1.Count - 1
'''        imx1(i).Width = imageX
'''        imx1(i).Height = imageY
'''    Next i
'''End Sub

Private Sub picJPG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuRightClick
    Else
        If RedMode = True Then
            xStr = X: yStr = Y
            bRedding = True
        ElseIf TextMode = True Then
            
'''            bTexting = True
'''            xStr = X: yStr = Y
'''            lblRedNote.Left = xStr: lblRedNote.Top = yStr
'''            lblRedNote.BorderStyle = 1
'''            lblRedNote.Visible = True
'            bRedded = True
            If Trim(lblRed(iRed).Caption) <> "" Then
                iRed = GetNextLabel
                If iRed > 0 Then
    '''                lblRed(iRed - 1).ForeColor = vbGreen '' RGB(111, 175, 28) ''(255, 160, 0)
                    lblRed(iRed - 1).WordWrap = False
                End If
            End If
            Call AddToUndo("lbl", iRed)
            Call UpdateSCD(1, 1, Abs(CInt(imgUtility(2).Enabled)))
            
            lblRed(iRed).ForeColor = lAnnoColor '' vbRed '' RGB(111, 175, 28) ''(255, 160, 0) '' vbRed
            lblRed(iRed).Tag = lAnnoColor
            lblRed(iRed).Caption = ""
            txtRed.Text = ""
            lblRed(iRed).Top = Y - 120: lblRed(iRed).Left = X
            lblRed(iRed).Width = picRed.Width - X
            lblEsc.Top = lblRed(iRed).Top + 45
            lblEsc.Left = lblRed(iRed).Left - lblEsc.Width - 60
            If lblEsc.Left < 0 Then
                lblEsc.Top = lblRed(iRed).Top - lblEsc.Height
                lblEsc.Left = lblRed(iRed).Left
            End If
            If lblEsc.Top > 0 And lblEsc.Left > 0 Then lblEsc.Visible = True _
                        Else lblEsc.Visible = False
            lblRed(iRed).Visible = True
            txtRed.Visible = True
            txtRed.SetFocus
        End If
'        bRedded = True
    End If
End Sub

Private Sub picJPG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedding Then
        picJPG.Line (xStr, yStr)-(X, Y)
        xStr = X: yStr = Y
    ElseIf bTexting = True Then
        If X - xStr > 0 Then lblRedNote.Width = X - xStr
    End If
End Sub

Private Sub picJPG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedding Then
        bRedded = True
        picJPG.Line (xStr, yStr)-(X, Y)
        bRedding = False
        mnuGRedSave.Enabled = True
        mnuGRedClear.Enabled = True
        iCurrUndo = iCurrUndo + 1
        Call AddToUndo("picjpg", iCurrUndo)
        Call UpdateSCD(1, 1, 0)
    ElseIf bTexting Then
        bRedded = True
        picJPG.AutoRedraw = True
        bTexting = False
        txt1.SetFocus
    End If
End Sub


Private Sub picRed_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X - x0, Y - y0
End Sub

Private Sub picRed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedLine And Button = vbLeftButton Then
'        bRedded = True
        bRedding = True
        x0 = X: y0 = Y
        
    ElseIf bRedText And Button = vbLeftButton Then
'        bRedded = True
        If Trim(lblRed(iRed).Caption) <> "" Then
            iRed = GetNextLabel
            If iRed > 0 Then
'''                lblRed(iRed - 1).ForeColor = vbGreen '' RGB(111, 175, 28) ''(255, 160, 0)
                lblRed(iRed - 1).WordWrap = False
            End If
        End If
        Call AddToUndo("lbl", iRed)
        Call UpdateSCD(1, 1, Abs(CInt(imgUtility(2).Enabled)))
        
        lblRed(iRed).ForeColor = lAnnoColor '' vbRed '' RGB(111, 175, 28) ''(255, 160, 0) '' vbRed
        lblRed(iRed).Tag = lAnnoColor
        lblRed(iRed).Caption = ""
        txtRed.Text = ""
        lblRed(iRed).Top = Y - 120: lblRed(iRed).Left = X
        lblRed(iRed).Width = picRed.Width - X
        lblEsc.Top = lblRed(iRed).Top + 45
        lblEsc.Left = lblRed(iRed).Left - lblEsc.Width - 60
        If lblEsc.Left < 0 Then
            lblEsc.Top = lblRed(iRed).Top - lblEsc.Height
            lblEsc.Left = lblRed(iRed).Left
        End If
        If lblEsc.Top > 0 And lblEsc.Left > 0 Then lblEsc.Visible = True _
                    Else lblEsc.Visible = False
        lblRed(iRed).Visible = True
        txtRed.Visible = True
        txtRed.SetFocus
    ElseIf Button = vbRightButton Then
        Me.PopupMenu mnuRightClick
    End If
End Sub

Private Sub picRed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedding Then
        picRed.Line (x0, y0)-(X, Y)
        x0 = X: y0 = Y
    End If
End Sub

Private Sub picRed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedding Then
        picRed.Line (x0, y0)-(X, Y)
        bRedded = True
        bRedding = False
        iCurrUndo = iCurrUndo + 1
        Call AddToUndo("picred", iCurrUndo)
        Call UpdateSCD(1, 1, 0)
    End If
End Sub


Private Sub sst1_Click(PreviousTab As Integer)
    If bReSize Then Exit Sub
    
    sTabDesc = sst1.TabCaption(sst1.Tab)
    Call SetType(iTabType(sst1.Tab))
    
    fraMulti.Left = picType.Left - fraMulti.Width
    fraMulti.Top = picType.Top + picType.Height - fraMulti.Height '' 360
'''''    picIconSize.Visible = True
    
    On Error Resume Next
    Select Case sst1.Tab
        Case 0
            mnuGfxApproval.Visible = False
''''''''            imgSearch.Visible = False
            lblSearch.Enabled = False
            If cboSHYR(0).Text = "" Then
                Me.MousePointer = 11
                If fSHYR(0) <> 0 Then
                    cboSHYR(0).Text = fSHYR(0)
                Else
                    cboSHYR(0).Text = CurrSHYR
                End If
            End If
            If cboCUNO(0).Text = "" Then
                If fBCC(0) <> "" And fFBCN(0) <> "" Then
                    Me.MousePointer = 11
                    cboCUNO(0).Text = fFBCN(0)
                End If
            Else
                fBCC(0) = Right("00000000" & cboCUNO(sst1.Tab).ItemData(cboCUNO(sst1.Tab).ListIndex), 8)
                fFBCN(0) = cboCUNO(sst1.Tab).List(cboCUNO(sst1.Tab).ListIndex)
            End If
            If cboSHCD.Text = "" Then
                If fSHCD(0) <> 0 And fSHNM(0) <> "" Then
                    Me.MousePointer = 11
                    cboSHCD.Text = fSHNM(0)
                End If
            End If
        Case 1
            mnuGfxApproval.Visible = False
''''''''            imgSearch.Visible = False
            lblSearch.Enabled = False
            If cboSHYR(1).Text = "" Then
                Me.MousePointer = 11
                If fSHYR(1) <> 0 Then
                    cboSHYR(1).Text = fSHYR(1)
                Else
                    cboSHYR(1).Text = CurrSHYR
                End If
            End If
            If cboCUNO(1).Text = "" Then
                If fBCC(1) <> "" And fFBCN(1) <> "" Then
                    Me.MousePointer = 11
                    cboCUNO(1).Text = fFBCN(1)
                End If
            Else
                fBCC(1) = Right("00000000" & cboCUNO(sst1.Tab).ItemData(cboCUNO(sst1.Tab).ListIndex), 8)
                fFBCN(1) = cboCUNO(sst1.Tab).List(cboCUNO(sst1.Tab).ListIndex)
            End If
        Case 2, 3, 4
            If sst1.Tab = 2 Then
''''''''                imgSearch.Visible = False
                lblSearch.Enabled = False
                mnuGfxApproval.Visible = False
            ElseIf sst1.Tab = 3 Then
''''''''                imgSearch.Visible = True
                lblSearch.Enabled = True
                mnuGfxApproval.Visible = False
            Else
''''''''                imgSearch.Visible = True
                lblSearch.Enabled = True
                mnuGfxApproval.Visible = True
            End If
            If sst1.Tab = 4 Then
'''''                picIconSize.Visible = False
                If cboCUNO(4).Text = "" Then lblSearch.Enabled = False Else lblSearch.Enabled = True
                picType.Visible = False
                
                fraMulti.Top = picReview.Top + picReview.Height - fraMulti.Height
                fraMulti.Left = picReview.Left - fraMulti.Width
'''                fraMulti.Visible = True
'''                picReview.Left = sst1.Width - 180 - picReview.Width
            End If
            If cboCUNO(sst1.Tab).Text = "" Then
                If fBCC(sst1.Tab) <> "" And fFBCN(sst1.Tab) <> "" Then
                    Me.MousePointer = 11
                    cboCUNO(sst1.Tab).Text = fFBCN(sst1.Tab)
                End If
            Else
                fBCC(sst1.Tab) = Right("00000000" & cboCUNO(sst1.Tab).ItemData(cboCUNO(sst1.Tab).ListIndex), 8)
                fFBCN(sst1.Tab) = cboCUNO(sst1.Tab).List(cboCUNO(sst1.Tab).ListIndex)
            End If
            
            
    End Select
    
    Call ClearModes(iModeTab)
    iModeTab = sst1.Tab
    Select Case iModeTab
        Case 0
            fraMulti.Visible = imx0(0).Visible
'            picType.Visible = imx0(0).Visible
        Case 1
            fraMulti.Visible = imx1(0).Visible
'            picType.Visible = imx1(0).Visible
        Case 2
            fraMulti.Visible = imx2(0).Visible
'            picType.Visible = imx2(0).Visible
        Case 3
            fraMulti.Visible = imx3(0).Visible
'            picType.Visible = imx3(0).Visible
        Case 4: If imx4(0).Visible Then fraMulti.Visible = True Else fraMulti.Visible = False
    End Select
    
    Me.MousePointer = 0
    
'    If cboCUNO(sst1.Tab).Text <> "" Then
'        tBCC = Right("00000000" & cboCUNO(sst1.Tab).ItemData(cboCUNO(sst1.Tab).ListIndex), 8)
'        tFBCN = cboCUNO(sst1.Tab).List(cboCUNO(sst1.Tab).ListIndex)
''        Debug.Print CLng(tBCC) & " - " & tFBCN
''        If sst1.Tab = 4 And bGFXReviewer Then
''            If CheckGFXCUNO(CLng(tBCC)) Then
''                picReview.Width = 6735
''                bApprover = True
''            Else
''                picReview.Width = 2400
''                bApprover = False
''            End If
''        End If
''    End If
End Sub

Private Sub tvwGraphics_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Index = 1 And UCase(Left(CurrNode, 1)) = "G" Then
        Debug.Print "CurrNode = " & CurrNode
        PopupMenu mnuCheck
        
    End If
    
    If Button = vbRightButton And Index = 2 Then
        If tvwGraphics(2).SelectedItem.Children = 0 Then PopupMenu mnuPart
    End If
End Sub

Private Sub tvwGraphics_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Index = 2 Then
        PopupMenu mnuPart
    End If
End Sub

Private Sub tvwGraphics_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    Dim strSelect As String, strInsert As String, strUpdate As String, sList As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim i As Integer, iLock As Integer, iCol As Integer, iRow As Integer, _
                i1 As Integer, i2 As Integer
    Dim imxCon As ImagXpress
    Dim lblCon As Label
    Dim lEID As Long, lFID As Long, lKID As Long
    
    picWait.ZOrder 0
    
    Screen.MousePointer = 11
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
'''        Call RedAlert(1, lblWelcome, redBCC, redSHCD)
    End If
    bRedSaved = False
    redBCC = "": redSHCD = 0
    
    Debug.Print Node.Key
    
    TNode = Node.Key
    
    
    ''THIS GETS SINGLE GRAPHIC WHEN IN TEXT MODE''
    If UCase(Left(Node.Key, 1)) = "G" Then
        Debug.Print Node.Key & " - YOU GOT ONE!"
        Call LoadGraphic(Index, Node.Key, Node.Text, Node.Parent.Key, Node.Parent.Text)
    
    
    ''THIS POPULATES ELEMENT LIST NODES''
    ElseIf UCase(Left(Node.Key, 1)) = "A" And Index = 1 Then ''POP ELEMENTS''
        Call LoadNodes(Index, Node.Key, Node.Text, Node.Parent.Key, Node.Parent.Text)
    
    
    ''THIS POPULATES SHOW INFO FOR POPUP DISPLAY''
    ElseIf UCase(Left(Node.Key, 1)) = "S" And Index = 1 Then ''GET SHOW INFO''
        Screen.MousePointer = 0
        strSelect = "SELECT SM.SHY56NAMA, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DAY')BEG_DAY, " & _
                    "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DAY')END_DAY, " & _
                    "AD.ALCTY1, AD.ALADDS " & _
                    "FROM " & F5601 & " SM, " & F0116 & " AD " & _
                    "WHERE SM.SHY56SHCD = " & Mid(Node.Key, 2) & " " & _
                    "AND SM.SHY56SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                    "AND SM.SHY56FCCDT = AD.ALAN8 (+)"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            
            mnuShowName.Caption = UCase(Trim(rst.Fields("SHY56NAMA")))
            If IsNull(rst.Fields("ALCTY1")) Then
                mnuShowLoc.Visible = False
'''                mnuShowLoc.Caption = "Show Location:  N/A"
            Else
                mnuShowLoc.Visible = True
                mnuShowLoc.Caption = "Show Location:  " & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                            UCase(Trim(rst.Fields("ALADDS")))
            End If
            mnuShowOpen.Caption = "Show Open:  " & UCase(Trim(rst.Fields("BEG_DAY"))) & _
                        "  " & UCase(Trim(rst.Fields("BEG_DATE")))
            mnuShowClose.Caption = "Show Close:  " & UCase(Trim(rst.Fields("END_DAY"))) & _
                        "  " & UCase(Trim(rst.Fields("END_DATE")))
            rst.Close: Set rst = Nothing
            PopupMenu mnuShowData
        Else
            rst.Close: Set rst = Nothing
        End If
    
    
    ''THIS GETS ELEMENT GFX WHEN ELEMENT NODE IS SELECTED''
    ElseIf UCase(Left(Node.Key, 1)) = "E" And Index = 1 Then ''GET ELEMENT GRAPHICS''
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        i1 = InStr(1, Node.Key, "-")
        lEID = Mid(Node.Key, i1 + 1)
        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GE.ELTID = " & lEID & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND (GM.GTYPE IN (" & sInType(Index) & ") OR GM.GTYPE = 87) " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    
    ''THIS GETS ARCHIVED ELEMENT GFX WHEN ARCHIVED ELEMENT NODE IS SELECTED''
    ElseIf UCase(Left(Node.Key, 1)) = "R" And Index = 1 Then ''GET ARCHIVED ELEMENT GRAPHICS''
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        fSHCD(Index) = Mid(Node.Parent.Key, 3)
        fSHNM(Index) = Node.Parent.Parent.Text
        i1 = InStr(1, Node.Key, "-")
        lEID = Mid(Node.Key, i1 + 1)
        strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GS.SHYR = " & fSHYR(Index) & " " & _
                    "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                    "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                    "AND GS.ELTID = " & lEID & " " & _
                    "AND GS.GID = GM.GID " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
'''        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
'''                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
'''                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
'''                    "Where GE.ELTID = " & lEID & " " & _
'''                    "AND GE.GID = GM.GID " & _
'''                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
'''                    "AND GM.GSTATUS IN (20, 30) " & _
'''                    "ORDER BY GM.GDESC"
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
        
    
    ''THIS GETS SHOW-SPECIFIC GRAPHICS FOR SHOW SEASON TAB''
    ElseIf UCase(Left(Node.Key, 2)) = "HS" And Index = 1 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Key
        CurrParText(Index) = Node.Text
        fSHCD(Index) = Mid(Node.Key, 3)
        fSHNM(Index) = Node.Parent.Text
        If bGPJ Then
            strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, " & _
                        "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                        "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GS.SHYR = " & fSHYR(Index) & " " & _
                        "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "MINUS " & _
                        "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                        "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                        "Where GS.SHYR = " & fSHYR(Index) & " " & _
                        "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GID > 0 " & _
                        "AND GM.GTYPE IN (1, 2, 3, 4) " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.FLR_ID = GF.FLR_ID " & _
                        "AND GF.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "ORDER BY GDESC"
        Else
            strSelect = " SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                        "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GS.SHYR = " & fSHYR(Index) & " " & _
                        "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GID > 0 " & _
                        "AND GM.GTYPE IN (1, 2, 3, 4) " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "MINUS " & _
                        "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                        "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                        "Where GS.SHYR = " & fSHYR(Index) & " " & _
                        "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GID > 0 " & _
                        "AND GM.GTYPE IN (1, 2, 3, 4) " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.FLR_ID = GF.FLR_ID " & _
                        "AND GF.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                        "AND GF.CLIENTRESTRICT_FLAG = 1 " & _
                        "ORDER BY GDESC"
        End If
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    
    ''THIS POPULATES THE IMAGES IN A SHOW SEASON FOLDER''
    ElseIf UCase(Left(Node.Key, 1)) = "F" And Index = 1 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        i1 = InStr(1, Node.Key, "-")
        lFID = Mid(Node.Key, i1 + 1)
        fSHCD(Index) = Mid(Node.Key, 2, i1 - 2)
        fSHNM(Index) = Node.Parent.Parent.Text
        strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GS.SHYR = " & fSHYR(Index) & " " & _
                    "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                    "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                    "AND GS.ELTID IS NULL " & _
                    "AND GS.GID = GM.GID " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID = " & lFID & " " & _
                    "ORDER BY GM.GDESC"
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    
    ''THIS POPULATES THE IMAGES IN A KIT-BASED FOLDER''
    ElseIf UCase(Left(Node.Key, 1)) = "F" And Index = 2 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        i1 = InStr(1, Node.Key, "-")
        lKID = Mid(Node.Key, 2, i1 - 2)
        lFID = Mid(Node.Key, i1 + 1)
        
        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GE.ELTID = " & lKID & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID = " & lFID & " " & _
                    "ORDER BY GM.GDESC"
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    
    ''THIS POPULATES THE IMAGES IN A CLIENT-BASED FOLDER''
    ElseIf UCase(Left(Node.Key, 1)) = "F" And Index = 3 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        lFID = Mid(Node.Key, 2)
        strSelect = "SELECT GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_MASTER GM " & _
                    "WHERE GM.FLR_ID = " & lFID & " " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    
    ''USER SELECTED A HEADER TITLE THAT WOULD NOT HAVE ANY IMAGES''
    ElseIf UCase(Node.Key) = "H0" Then
        CurrParNode(Index) = ""
        CurrParText(Index) = ""
        Call ClearThumbnails3(0)
        CurrSelect(Index) = ""
    
    
    ElseIf UCase(Left(Node.Key, 1)) = "T" And Index = 1 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        
                    
        Select Case UCase(Left(Node.Parent.Key, 1))
            Case "H" ''SHOW-SPECIFIC''
                fSHCD(Index) = Mid(Node.Parent.Key, 3)
                fSHNM(Index) = Node.Parent.Parent.Text
                strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, " & _
                            "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                            "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                            "Where GS.SHYR = " & fSHYR(Index) & " " & _
                            "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                            "AND GS.AN8_SHCD = " & fSHCD(Index) & " " & _
                            "AND GS.ELTID IS NULL " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GTYPE = " & Right(Node.Key, 1) & " " & _
                            "AND GM.GSTATUS IN (" & defSIN & ") " & _
                            "ORDER BY GM.GDESC"
            Case "E" ''ELEMENT''
                i1 = InStr(1, Node.Key, "-")
                i2 = InStr(i1 + 1, Node.Key, "-")
                lEID = Mid(Node.Key, i1 + 1, i2 - i1 - 1)
                strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
                            "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                            "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
                            "Where GE.ELTID = " & lEID & " " & _
                            "AND GE.GID = GM.GID " & _
                            "AND GM.GTYPE = " & Right(Node.Key, 1) & " " & _
                            "AND GM.GSTATUS IN (" & defSIN & ") " & _
                            "ORDER BY GM.GDESC"
        End Select
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    ElseIf UCase(Left(Node.Key, 1)) = "T" And Index = 2 Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        lblInactive(Index).Visible = False
        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                    "WHERE GE.ELTID = " & Mid(Node.Key, 2, Len(Node.Key) - 3) & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND GM.GTYPE = " & Right(Node.Key, 1) & " " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
    
    ElseIf UCase(Left(Node.Key, 1)) = "K" And Index = 2 Then ''GET KIT GRAPHICS''
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = ""
        CurrParText(Index) = ""
        lKID = Mid(Node.Key, 2)
        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GE.ELTID = " & lKID & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
        
'''        Call ClearThumbnails2(0)
'''        CurrSelect(Index) = ""
        
    ElseIf UCase(Left(Node.Key, 1)) = "E" And Index = 2 Then ''GET ELEMENT GRAPHICS''
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        lEID = Mid(Node.Key, 2)
        strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, " & _
                    "GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                    "FROM ANNOTATOR.GFX_ELEMENT GE, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GE.ELTID = " & lEID & " " & _
                    "AND GE.GID = GM.GID " & _
                    "AND (GM.GTYPE IN (" & sInType(Index) & ") OR GM.GTYPE = 87) " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "ORDER BY GM.GDESC"
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
        
    ElseIf UCase(Left(Node.Key, 1)) = "D" And Index = 3 Then
        picWait.Visible = True
        picWait.Refresh
'''        MsgBox Node.Key
'        CurrParNode(Index) = Node.Parent.key
'        CurrParText(Index) = Node.Parent.Text
        lblInactive(Index).Visible = False
        
'        iCurrStatus = CInt(Mid(Node.key, 2, 2))
        
'''        strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
'''                    "FROM " & GFXMas & " GM " & _
'''                    "WHERE GM.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                    "AND GM.GTYPE = " & Right(Node.key, 1) & " " & _
'''                    "AND GM.GSTATUS = " & CInt(Mid(Node.key, 2, 2)) & " " & _
'''                    "AND TRUNC(GM.ADDDTTM) BETWEEN TO_DATE('" & Mid(Node.key, 5, 2) & "/01/" & Mid(Node.key, 7, 4) & "', 'MM/DD/YYYY') " & _
'''                    "AND TO_DATE('" & format(DateAdd("M", 1, DateValue(Mid(Node.key, 5, 2) & "/01/" & Mid(Node.key, 7, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY') " & _
'''                    "ORDER BY GM.GDESC"
'        strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
'                    "FROM " & GFXMas & " GM " & _
'                    "WHERE GM.AN8_CUNO = " & CLng(tBCC) & " " & _
'                    "AND GM.GTYPE = " & Right(Node.key, 1) & " " & _
'                    "AND GM.GSTATUS IN (20, 30) " & _
'                    "AND TRUNC(GM.ADDDTTM) BETWEEN TO_DATE('" & Mid(Node.key, 2, 2) & "/01/" & Mid(Node.key, 4, 4) & "', 'MM/DD/YYYY') " & _
'                    "AND TO_DATE('" & format(DateAdd("M", 1, DateValue(Mid(Node.key, 2, 2) & "/01/" & Mid(Node.key, 4, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY') " & _
'                    "ORDER BY GM.GDESC"
        strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS, GM.GTYPE " & _
                    "FROM " & GFXMas & " GM " & _
                    "WHERE GM.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                    "AND GM.GTYPE IN (" & sInType(Index) & ") " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND TRUNC(GM.ADDDTTM) BETWEEN TO_DATE('" & Mid(Node.Key, 2, 2) & "/01/" & Mid(Node.Key, 4, 4) & "', 'MM/DD/YYYY') " & _
                    "AND TO_DATE('" & Format(DateAdd("M", 1, DateValue(Mid(Node.Key, 2, 2) & "/01/" & Mid(Node.Key, 4, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY') " & _
                    "ORDER BY GM.GDESC"
                    
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
        
        If bGFXReviewer Then
'''            cmdStatusEdit_All.ToolTipText = "Use this control to Edit all '" & Node.Parent.Text & _
'''                        " - " & Node.Text & "'."
            cmdStatusEdit_View.ToolTipText = "Use this control to Edit all '" & Node.Text & _
                        "' in the window below."
            For i = 0 To 2
                If (CInt(Mid(Node.Key, 2, 2)) / 10) - 1 = i Then
                    mnuStatusReset(i).Visible = False
                Else
                    mnuStatusReset(i).Visible = True
                End If
            Next i
        End If
    
    ElseIf UCase(Left(Node.Key, 1)) = "I" Then
        picWait.Visible = True
        picWait.Refresh
        CurrParNode(Index) = Node.Parent.Key
        CurrParText(Index) = Node.Parent.Text
        lblInactive(Index).Visible = False
        Select Case Index
            Case 0
                If UCase(Mid(Node.Key, 2, 1)) = "S" Then
                    strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                                "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                                "WHERE GS.SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                                "AND GS.AN8_SHCD = " & Mid(Node.Key, 7) & " " & _
                                "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                                "AND GS.ELTID IS NULL " & _
                                "AND GS.GID = GM.GID " & _
                                "AND GM.GTYPE = " & CInt(Mid(Node.Key, 3, 1)) & " " & _
                                "AND GM.GSTATUS = " & CInt(Mid(Node.Key, 4, 2)) & " " & _
                                "ORDER BY GM.GDESC"
                ElseIf UCase(Mid(Node.Key, 2, 1)) = "E" Then
                    strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                                "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                                "WHERE GE.ELTID = " & Mid(Node.Key, 7) & " " & _
                                "AND GE.GID = GM.GID " & _
                                "AND GM.GTYPE = " & CInt(Mid(Node.Key, 3, 1)) & " " & _
                                "AND GM.GSTATUS = " & CInt(Mid(Node.Key, 4, 2)) & " " & _
                                "ORDER BY GM.GDESC"
                End If
                
'''                If UCase(Mid(Node.Key, 2, 1)) = "S" Then
'''                    strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
'''                                "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
'''                                "WHERE GS.SHYR = " & CInt(cboSHYR(1).Text) & " " & _
'''                                "AND GS.AN8_SHCD = " & Mid(Node.Key, 4) & " " & _
'''                                "AND GS.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                                "AND GS.GID = GM.GID " & _
'''                                "AND GM.GTYPE = " & CInt(Mid(Node.Key, 3, 1)) & " " & _
'''                                "ORDER BY GM.GDESC"
'''                ElseIf UCase(Mid(Node.Key, 2, 1)) = "E" Then
'''                    strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
'''                                "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
'''                                "WHERE GE.ELTID = " & Mid(Node.Key, 4) & " " & _
'''                                "AND GE.GID = GM.GID " & _
'''                                "AND GM.GTYPE = " & CInt(Mid(Node.Key, 3, 1)) & " " & _
'''                                "ORDER BY GM.GDESC"
'''                End If
''''''                CurrSelect(Index) = strSelect
''''''                Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, "1-20", TNode)
''''''                Screen.MousePointer = 0
''''''                Exit Sub
            Case 1
                strSelect = "SELECT GS.SHOW_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                            "WHERE GS.SHYR = " & CInt(cboSHYR(1).Text) & " " & _
                            "AND GS.AN8_SHCD = " & Mid(Node.Key, 6) & " " & _
                            "AND GS.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                            "AND GS.ELTID IS NULL " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GTYPE = " & CInt(Mid(Node.Key, 2, 1)) & " " & _
                            "AND GM.GSTATUS = " & CInt(Mid(Node.Key, 3, 2)) & " " & _
                            "ORDER BY GM.GDESC"
'''                CurrSelect(Index) = strSelect
'''                Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, "1-20", TNode)
'''                Screen.MousePointer = 0
'''                Exit Sub
            Case 2
                strSelect = "SELECT GE.ES_ID, GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                            "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                            "WHERE GE.ELTID = " & Mid(Node.Key, 6) & " " & _
                            "AND GE.GID = GM.GID " & _
                            "AND GM.GTYPE = " & CInt(Mid(Node.Key, 2, 1)) & " " & _
                            "AND GM.GSTATUS = " & CInt(Mid(Node.Key, 3, 2)) & " " & _
                            "ORDER BY GM.GDESC"
'''                CurrSelect(Index) = strSelect
'''                Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, "1-20", TNode)
'''                Screen.MousePointer = 0
'''                Exit Sub
            Case 3
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GSTATUS " & _
                            "FROM " & GFXMas & " GM " & _
                            "WHERE GM.AN8_CUNO = " & CLng(fBCC(Index)) & " " & _
                            "AND GM.GTYPE = " & Mid(Node.Key, 2, 1) & " " & _
                            "AND GM.GSTATUS = " & CInt(Right(Node.Key, 2)) & " " & _
                            "AND TRUNC(GM.ADDDTTM) BETWEEN TO_DATE('" & Mid(Node.Key, 3, 2) & "/01/" & Mid(Node.Key, 5, 4) & "', 'MM/DD/YYYY') " & _
                            "AND TO_DATE('" & Format(DateAdd("M", 1, DateValue(Mid(Node.Key, 3, 2) & "/01/" & Mid(Node.Key, 5, 4))) - 1, "MM/DD/YYYY") & "', 'MM/DD/YYYY') " & _
                            "ORDER BY GM.GDESC"
'''                CurrSelect(Index) = strSelect
'''                Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, "1-20", TNode)
'''                Screen.MousePointer = 0
'''                Exit Sub
        End Select
        
        CurrSelect(Index) = strSelect
        iListStart(Index) = 1
        Call GetGraphics(Index, (0 + (sst1.Tab * 10)), strSelect, iListStart(Index), TNode)
        
        
        
'''        Set rst = Conn.Execute(strSelect)
'''        i = 0
''''''''        picInner(Index).Visible = False
'''        Do While Not rst.EOF
'''            iCol = Int(i / iRows): iRow = i Mod iRows
'''            Select Case Index
'''                Case 0
'''                    If i >= imx0.Count Then Load imx0(i)
'''                    Set imxCon = imx0(i)
'''                Case 1
'''                    If i >= imx1.Count Then Load imx1(i)
'''                    Set imxCon = imx1(i)
'''                Case 2
'''                    If i >= imx2.Count Then Load imx2(i)
'''                    Set imxCon = imx2(i)
'''                Case 3
'''                    If i >= imx3.Count Then Load imx3(i)
'''                    Set imxCon = imx3(i)
'''            End Select
'''            With imxCon
'''                .Left = 120 + (iCol * spaceX)
'''                .Top = 120 + (iRow * spaceY)
'''                .Update = False
'''                .PICThumbnail = iRes
'''                If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
'''                    .FileName = sGPath & "Acrobatid.bmp"
'''                Else
'''                    .FileName = Trim(rst.Fields("GPATH"))
'''                End If
'''                .ToolTipText = Trim(rst.Fields("GDESC"))
'''                .Update = True
'''                .Buttonize 1, 1, 50
'''                .Visible = True
'''                .Refresh
'''                Select Case Index
'''                    Case 0
'''                        If UCase(Mid(Node.Key, 2, 1)) = "S" Then
'''                            .Tag = "S" & CStr(rst.Fields("SHOW_ID"))
'''                        ElseIf UCase(Mid(Node.Key, 2, 1)) = "E" Then
'''                            .Tag = "E" & CStr(rst.Fields("ES_ID"))
'''                        End If
'''                    Case 1: .Tag = CStr(rst.Fields("SHOW_ID"))
'''                    Case 2: .Tag = CStr(rst.Fields("ES_ID"))
'''                    Case 3: .Tag = CStr(rst.Fields("GID"))
'''                End Select
'''            End With
'''            Select Case Index
'''                Case 0
'''                    If i >= lbl0.Count Then Load lbl0(i)
'''                    Set lblCon = lbl0(i)
'''                Case 1
'''                    If i >= lbl1.Count Then Load lbl1(i)
'''                    Set lblCon = lbl1(i)
'''                Case 2
'''                    If i >= lbl2.Count Then Load lbl2(i)
'''                    Set lblCon = lbl2(i)
'''                Case 3
'''                    If i >= lbl3.Count Then Load lbl3(i)
'''                    Set lblCon = lbl3(i)
'''            End Select
'''            With lblCon
''''''                If Len(Trim(rst.FIELDS("GDESC"))) > 25 Then
''''''                    .Caption = Left(Trim(rst.FIELDS("GDESC")), 25) & "..."
''''''                Else
'''                    .Caption = Trim(rst.Fields("GDESC"))
''''''                End If
'''                .Left = 120 + (iCol * spaceX) + ((imageX - .Width) / 2)
'''                .Top = 120 + imageY + (iRow * spaceY)
'''                If rst.Fields("GSTATUS") > 0 Then
'''                    .BackColor = vbWindowBackground
'''                Else
'''                    .BackColor = vbRed
'''                    lblInactive(Index).Visible = True
'''                End If
'''                .Visible = True
'''            End With
'''            i = i + 1
'''            rst.MoveNext
'''        Loop
'''        rst.Close: Set rst = Nothing
'''
'''        Select Case Index
'''            Case 0
'''                For i = i To imx0.Count - 1
'''                    imx0(i).Visible = False
'''                    imx0(i).FileName = ""
'''                    lbl0(i).Visible = False
'''                Next i
'''            Case 1
'''                For i = i To imx1.Count - 1
'''                    imx1(i).Visible = False
'''                    imx1(i).FileName = ""
'''                    lbl1(i).Visible = False
'''                Next i
'''            Case 2
'''                For i = i To imx2.Count - 1
'''                    imx2(i).Visible = False
'''                    imx2(i).FileName = ""
'''                    lbl2(i).Visible = False
'''                Next i
'''            Case 3
'''                Call ResetCounts(i, 0)
'''                For i = i To imx3.Count - 1
'''                    imx3(i).Visible = False
'''                    imx3(i).FileName = ""
'''                    lbl3(i).Visible = False
'''                Next i
'''
'''        End Select
'''        picInner(Index).Width = ((iCol + 1) * spaceX) + 240
'''        If picInner(Index).Width < picOuter(Index).ScaleWidth Then
'''            hsc1(Index).Max = picInner(Index).Width / 100
'''            hsc1(Index).Visible = False
'''        Else
'''            hsc1(Index).Max = (picInner(Index).Width / 100) - (picOuter(Index).ScaleWidth / 100)
'''            hsc1(Index).Visible = True
'''        End If
'''        hsc1(Index).Value = 0 '''picOuter(1).ScaleWidth
'''        hsc1(Index).LargeChange = picOuter(Index).ScaleWidth / 100
'''
'''        picInner(Index).Visible = True
    End If
    
    If bDMode Or bEMode Then ClearModes (iModeTab)
    
    picWait.Visible = False
    Screen.MousePointer = 0
End Sub

Public Sub LoadThePicture(sPath As String, pRed As Boolean)
    On Error GoTo ErrorOpening
    
    picPDF.Visible = False
    picJPG.Visible = False
    picPDFTools.Visible = False
    Select Case pRed
        Case True
            picJPGTools.Visible = False
            picRedTools.Visible = True
        Case False
            picJPGTools.Visible = True
            picRedTools.Visible = False
            picJPG.MousePointer = 0
            TextMode = False
            RedMode = False
    End Select
    
    imgRedReload.Picture = imlRedMode.ListImages(5).Picture
    picPDF.MousePointer = 0
    
'''    MsgBox "picJPG.Visible = False"
'''    lblByGeorge(0).Visible = False
'''    MsgBox "lblByGeorge(0).Visible = False"
'''    lblByGeorge(1).Visible = False
'''    MsgBox "lblByGeorge(1).Visible = False"
    Set picJPG.Picture = LoadPicture()
    
    
'''    MsgBox "picJPG cleared"
    Set imgSize.Picture = LoadPicture(sPath)
'''    MsgBox "imgSize loaded"
    Debug.Print "X = " & imgSize.Width & ", y = " & imgSize.Height
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
    
'''    MsgBox "imgSize captured"
    
'''    MsgBox "picJPG.AutoRedraw = " & picJPG.AutoRedraw
'''    MsgBox "me.AutoRedraw = " & Me.AutoRedraw
'''    MsgBox "picJPG.Width = " & picJPG.Width & vbNewLine & _
'''                "picJPG.Height = " & picJPG.Height
    
    picJPG.PaintPicture imgSize.Picture, 0, 0, picJPG.Width, picJPG.Height
    
'    Call AddToUndo("picjpg", -1)
    
'''    MsgBox "picJPG.PaintPicture complete"
    
'''    picJPG.ScaleWidth = imgSize.Width
'''    picJPG.ScaleHeight = imgSize.Height
'''    picJPG.Picture = LoadPicture(sPath, vbLPCustom, 0, picJPG.Width, picJPG.Height)
    
    Call SetImageState
'''    If picJPG.Width = maxX Or picJPG.Height = maxY Then
'''        rMX = picJPG.Width: rMY = picJPG.Height
'''        rSX = imgSize.Width: rSY = imgSize.Height
'''        dMTop = picJPG.Top: dMLeft = picJPG.Left
'''        dSTop = picJPG.Top: dSLeft = picJPG.Left
'''        iImageState = 1
'''    Else
'''        rSX = picJPG.Width: rSY = picJPG.Height
'''        dSTop = picJPG.Top: dSLeft = picJPG.Left
'''        Select Case rAsp
'''            Case Is = rFAsp
'''                rMX = maxX: rMY = maxY
'''                dMTop = dTop: dMLeft = dLeft
'''            Case Is > rFAsp 'X IS DETERMINING FACTOR'
'''                rMX = maxX: rMY = picJPG.Height * (maxX / picJPG.Width)
'''                dMLeft = dLeft: dMTop = dTop + (maxY - rMY) / 2
'''            Case Is < rFAsp 'Y IS DETERMINING FACTOR'
'''                rMY = maxY: rMX = picJPG.Width * (maxY / picJPG.Height)
'''                dMTop = dTop: dMLeft = dLeft + (maxX - rMX) / 2
'''        End Select
'''    End If
                
    
'''    '///// CHECK IF IMAGE IS STRETCHED \\\\\
'''    If picJPG.Width > imgSize.Width Then
'''        mnuResizeGraphic.Visible = True
'''        mnuMaxGraphic.Visible = True
'''        mnuResizeGraphic.Enabled = True
'''        mnuMaxGraphic.Enabled = False
'''        lblresize.Caption = "Resize"
'''        lblresize.Visible = True
'''    Else
'''        mnuResizeGraphic.Visible = False
'''        mnuMaxGraphic.Visible = False
'''        lblresize.Visible = False
'''    End If
    
    '///// CHECK IF IMAGE COULD BE STRETCHED \\\\\
    If picJPG.Width < maxX _
                And picJPG.Height < maxY Then
        mnuResizeGraphic.Visible = True
        mnuMaxGraphic.Visible = True
        mnuResizeGraphic.Enabled = False
        mnuMaxGraphic.Enabled = True
''''''''        lblResize.Caption = "Maximize"
''''''''        lblResize.Enabled = True
''''''''        lblFullSize.Enabled = False
        Call ResetJPGZoom(2, 1, 0)
    Else
        mnuResizeGraphic.Visible = False
        mnuMaxGraphic.Visible = False
''''''''        lblResize.Enabled = False
''''''''        lblFullSize.Enabled = True
        Call ResetJPGZoom(0, 2, 1)
    End If
    
'''    Set imgSize.Picture = LoadPicture()
    
    picJPG.Visible = True
'''    cmdMenu.Visible = True
    imgMenu.Visible = True: lblMenu.Visible = True
    picMenu2.Visible = True
    bPicLoaded = True
    iPDFPage = 0
    RedName = lRedID & "-" & iPDFPage & "RED.bmp"
Exit Sub

ErrorOpening:
    MsgBox "Error encountered while attempting to open file." & vbNewLine & _
                "Error:  " & Err.Description, vbCritical, "Cannot Open File..."
    picJPG.Visible = False
    bPicLoaded = False
'''    cmdMenu.Visible = False
'''    lblresize.Visible = False
    imgMenu.Visible = False: lblMenu.Visible = False
    picMenu2.Visible = False
'''    Set imgSize.Picture = LoadPicture()
End Sub

'''Public Sub LoadThePDF(sPath As String)
'''    picJPG.Visible = False
'''
'''    mnuRedlining.Enabled = False
'''    mnuResizeGraphic.Visible = False
'''    mnuMaxGraphic.Visible = False
'''
'''    imgMenu.Visible = True: lblMenu.Visible = True
'''    picMenu2.Visible = False
''''''    lblByGeorge(0).Visible = False
''''''    lblByGeorge(1).Visible = False
''''    pdfGraphic.src = sPath
'''
'''    On Error Resume Next: Err = 0
'''    pdfGraphic.LoadFile (sPath)
'''    If Err = 0 Then
'''        pdfGraphic.setShowToolbar (False)
''''    pdfGraphic.setView ("Fit")
''''    pdfGraphic.setShowToolbar (False)
''''    Me.Refresh
'''
''''    pdfGraphic.setZoom -1
'''        pdfGraphic.Visible = True
'''        bPicLoaded = False
'''    Else
'''        pdfGraphic.Visible = False
'''        MsgBox "Unable to open image." & vbNewLine & _
'''                    "Error(" & Err.Number & "): " & Err.Description, _
'''                    vbExclamation, "Sorry..."
'''    End If
'''
''''''    hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
'''End Sub

Public Sub LoadPDF(pFile As String)

    Call ClearLabels
    picJPG.Visible = False
    
    
    On Error Resume Next
    xpdf1.loadFile pFile
    
    If Err = -2147220988 Then
        MsgBox "Unable to open PDF file.  File might be password protected.  " & _
                    "The Annotator does not allow the import of password protected images.", _
                    vbCritical, "Sorry..."
        xpdf1.Visible = False
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Unable to open PDF file", vbCritical, "Sorry..."
        xpdf1.Visible = False
        Exit Sub
    End If
    
    iPDFPage = xpdf1.currentPage
    RedName = lRedID & "-" & iPDFPage & "RED.bmp"
    
    imgRedReload.Picture = imlRedMode.ListImages(6).Picture
    xpdf1.matteColor = Me.BackColor
    
    ''SET ZOOM MODE''
    iZoomMode = 0
    Call SetZoomMode(iZoomMode)
    dZF = xpdf1.zoomPercent
    cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
    
    ''SET PAGE MODE''
    iPageMode = 0
    Call SetPageMode(iPageMode)
    
    ''BE CERTAIN PDF CONTAINERS ARE VISIBLE''
    picPDF.Visible = True
    picJPGTools.Visible = False
    picPDFTools.Visible = True
    picRedTools.Visible = False
    
    xpdf1.Visible = True
    xpdf1.enableMouseEvents = True
    xpdf1.enableSelect = False
    
    mnuResizeGraphic.Enabled = True: mnuResizeGraphic.Visible = True
    mnuMaxGraphic.Enabled = True: mnuMaxGraphic.Visible = True
    
    imgMenu.Visible = True: lblMenu.Visible = True
    picMenu2.Visible = True
    bPicLoaded = True
    
    bRedMode = False: shpRN.Visible = False: lblRN.Visible = False
    bRedLine = False
    bRedText = False
    picRed.Visible = False
    picRed.MousePointer = 0

End Sub
Public Sub PopShowGraphics(tmpCUNO As String, tmpSHYR As Integer, tmpSHCD As Long, bFutureShow As Boolean)
    Dim strSelect As String, sList As String, sStat As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sPNode As String, sGNode As String, sSNode As String, sTNode As String, _
                sDesc As String, sDescPar As String
    Dim sKNode As String, sENode As String
    Dim lParent As Long, lElem As Long
    Dim iType As Integer
    Dim sGStatus(0 To 30) As String
    Dim iGStatus(0 To 30) As Integer
    Dim tblElement As String, strAnd As String, strDisc As String, strField As String
    
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(0) = "DE-ACTIVED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    iGStatus(0) = 10
    iGStatus(10) = 7
    iGStatus(20) = 8
    iGStatus(30) = 9
    
    If bPerm(56) Then ''ABLE TO VIEW DRAFTS''
        sList = "0, 10, 20, 30"
    Else ''NOT ABLE TO VIEW DRAFTS''
        sList = "0, 20, 30"
    End If
    
    tvwGraphics(0).Visible = False
    tvwGraphics(0).Nodes.Clear
    tvwGraphics(0).ImageList = ImageList1
    If bPerm(29) Then
        If iView = 0 Then
            strSelect = "SELECT GM.GDESC, GM.GTYPE, GS.SHOW_ID, GM.GSTATUS " & _
                        "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                        "WHERE GS.SHYR = " & tmpSHYR & " " & _
                        "AND GS.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & sList & ") " & _
                        "ORDER BY GM.GTYPE, GM.GSTATUS, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                        "WHERE GS.SHYR = " & tmpSHYR & " " & _
                        "AND GS.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & sList & ") " & _
                        "ORDER BY GM.GTYPE, GM.GSTATUS"
        End If
    Else
        If iView = 0 Then
            strSelect = "SELECT GM.GDESC, GM.GTYPE, GS.SHOW_ID, GM.GSTATUS " & _
                        "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                        "WHERE GS.SHYR = " & tmpSHYR & " " & _
                        "AND GS.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & sList & ") " & _
                        "AND GM.GTYPE <> 3 " & _
                        "ORDER BY GM.GTYPE, GM.GSTATUS, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                        "WHERE GS.SHYR = " & tmpSHYR & " " & _
                        "AND GS.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID " & _
                        "AND GM.GSTATUS IN (" & sList & ") " & _
                        "AND GM.GTYPE <> 3 " & _
                        "ORDER BY GM.GTYPE, GM.GSTATUS"
        End If
    End If
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sGNode = "": sSNode = "": sTNode = ""
        sPNode = "ShowSpec"
        sDesc = "Show-Specific Graphics"
        Set nodX = tvwGraphics(0).Nodes.Add(, , sPNode, sDesc, 5)
        Do While Not rst.EOF
            If iView = 0 Then
                '///// ADD GRAPHIC TO TREEVIEW \\\\\
                sGNode = "gs" & rst.Fields("SHOW_ID")
                sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
                iType = rst.Fields("GTYPE")
                Set nodX = tvwGraphics(0).Nodes.Add(sPNode, tvwChild, sGNode, sDesc, iType)
            Else
                If sTNode <> "t" & rst.Fields("GTYPE") & tmpSHCD Then
                    sTNode = "t" & rst.Fields("GTYPE") & tmpSHCD
                    sDesc = GfxType(rst.Fields("GTYPE"))
                    iType = rst.Fields("GTYPE")
                    Set nodX = tvwGraphics(0).Nodes.Add(sPNode, tvwChild, sTNode, sDesc, iType)
                End If
                Select Case Len(rst.Fields("GSTATUS"))
                    Case 1
                        sStat = "0" & CStr(rst.Fields("GSTATUS"))
                    Case 2
                        sStat = CStr(rst.Fields("GSTATUS"))
                End Select
                sSNode = "is" & rst.Fields("GTYPE") & sStat & "-" & tmpSHCD
                sDesc = sGStatus(rst.Fields("GSTATUS")) ''' & " Graphics"
                Set nodX = tvwGraphics(0).Nodes.Add(sTNode, tvwChild, sSNode, sDesc, _
                            iGStatus(rst.Fields("GSTATUS")))
            End If
            
            
'''            '///// ADD TO ASSIGN LIST W/NO ITEMDATA \\\\\
'''            lstGraphics.AddItem sDesc
            
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    If bFutureShow Then
        tblElement = "ANNOTATOR.GFX_ELEMENT"
        strField = "ES_ID"
        strAnd = ""
        strDisc = ""
    Else
        tblElement = "ANNOTATOR.GFX_SHOW"
        strField = "SHOW_ID"
        strAnd = "AND EU.SHYR = GE.SHYR " & _
                    "AND EU.AN8_CUNO = GE.AN8_CUNO " & _
                    "AND EU.AN8_SHCD = GE.AN8_SHCD "
        strDisc = "  (Archived Graphics)"
    End If
    
    If bPerm(29) Then
        If iView = 0 Then
            strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
                        "GE." & strField & ", GM.GTYPE, GM.GDESC, GM.GSTATUS " & _
                        "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & _
                        "" & IGLKit & " K, " & tblElement & " GE, " & GFXMas & " GM " & _
                        "WHERE KU.SHYR = " & tmpSHYR & " " & _
                        "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND KU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND KU.SHYR = EU.SHYR " & _
                        "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                        "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                        "AND KU.KITUSEID = EU.KITUSEID " & _
                        "AND EU.SHSTATUS <> 3 " & _
                        "AND EU.KITID = K.KITID " & _
                        "AND EU.ELTID = GE.ELTID " & _
                        strAnd & _
                        "AND GE.GID = GM.GID " & _
                        "ORDER BY K.KITREF, EU.ELTFNAME, GM.GTYPE, GM.GSTATUS, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT K.KITREF, K.KITID, K.KITFNAME, EU.ELTID, " & _
                        "EU.ELTFNAME, EU.ELTDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & _
                        "" & IGLKit & " K, " & tblElement & " GE, " & GFXMas & " GM " & _
                        "WHERE KU.SHYR = " & tmpSHYR & " " & _
                        "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND KU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND KU.SHYR = EU.SHYR " & _
                        "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                        "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                        "AND KU.KITUSEID = EU.KITUSEID " & _
                        "AND EU.SHSTATUS <> 3 " & _
                        "AND EU.KITID = K.KITID " & _
                        "AND EU.ELTID = GE.ELTID " & _
                        strAnd & _
                        "AND GE.GID = GM.GID " & _
                        "ORDER BY K.KITREF, EU.ELTFNAME, GM.GTYPE, GM.GSTATUS"
'''                        "AND GM.GSTATUS IN (" & sList & ") " & _
'''                        "ORDER BY K.KITREF, EU.ELTFNAME, GM.GTYPE, GM.GSTATUS"
        End If
    Else
        If iView = 0 Then
            strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
                        "GE." & strField & ", GM.GTYPE, GM.GDESC, GM.GSTATUS " & _
                        "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & _
                        "" & IGLKit & " K, " & tblElement & " GE, " & GFXMas & " GM " & _
                        "WHERE KU.SHYR = " & tmpSHYR & " " & _
                        "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND KU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND KU.SHYR = EU.SHYR " & _
                        "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                        "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                        "AND KU.KITUSEID = EU.KITUSEID " & _
                        "AND EU.SHSTATUS <> 3 " & _
                        "AND EU.KITID = K.KITID " & _
                        "AND EU.ELTID = GE.ELTID " & _
                        strAnd & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GTYPE <> 3 " & _
                        "ORDER BY K.KITREF, EU.ELTFNAME, GM.GTYPE, GM.GSTATUS, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT K.KITREF, K.KITID, K.KITFNAME, EU.ELTID, " & _
                        "EU.ELTFNAME, EU.ELTDESC, GM.GTYPE, GM.GSTATUS " & _
                        "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & _
                        "" & IGLKit & " K, " & tblElement & " GE, " & GFXMas & " GM " & _
                        "WHERE KU.SHYR = " & tmpSHYR & " " & _
                        "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                        "AND KU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND KU.SHYR = EU.SHYR " & _
                        "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                        "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                        "AND KU.KITUSEID = EU.KITUSEID " & _
                        "AND EU.SHSTATUS <> 3 " & _
                        "AND EU.KITID = K.KITID " & _
                        "AND EU.ELTID = GE.ELTID " & _
                        strAnd & _
                        "AND GE.GID = GM.GID " & _
                        "AND GM.GTYPE <> 3 " & _
                        "ORDER BY K.KITREF, EU.ELTFNAME, GM.GTYPE, GM.GSTATUS"
        End If
    End If
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sDesc = UCase(Trim(rst.Fields("KITFNAME"))) & " Kit" & strDisc
        sDescPar = UCase(Trim(rst.Fields("KITFNAME")))
        lParent = rst.Fields("KITID")
        sKNode = "k" & rst.Fields("KITID")
        Set nodX = tvwGraphics(0).Nodes.Add(, , sKNode, sDesc, 5)
        Do While Not rst.EOF
            If rst.Fields("KITID") = lParent Then
                '///// THIS IS A CHILD \\\\\
                lElem = rst.Fields("ELTID")
                sENode = "e" & lElem
                sDesc = sDescPar & "-" & UCase(Trim(rst.Fields("ELTFNAME"))) & "  " & _
                            UCase(Trim(rst.Fields("ELTDESC")))
                Set nodX = tvwGraphics(0).Nodes.Add(sKNode, tvwChild, sENode, sDesc, 5)
                Do While rst.Fields("ELTID") = lElem
                    If InStr(1, sList, rst.Fields("GSTATUS")) <> 0 Then
                        If iView = 0 Then
                            If bFutureShow Then
                                sGNode = "ge" & rst.Fields("ES_ID")
                            Else
                                sGNode = "gs" & rst.Fields("SHOW_ID")
                            End If
                            sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
                            iType = rst.Fields("GTYPE")
                            Set nodX = tvwGraphics(0).Nodes.Add(sENode, tvwChild, sGNode, sDesc, iType)
                        Else
                            If sTNode <> "ie" & rst.Fields("GTYPE") & rst.Fields("ELTID") Then
                                sTNode = "ie" & rst.Fields("GTYPE") & rst.Fields("ELTID")
                                sDesc = GfxType(rst.Fields("GTYPE"))
                                iType = rst.Fields("GTYPE")
                                Set nodX = tvwGraphics(0).Nodes.Add(sENode, tvwChild, sTNode, sDesc, iType)
                            End If
                            Select Case Len(rst.Fields("GSTATUS"))
                                Case 1
                                    sStat = "0" & CStr(rst.Fields("GSTATUS"))
                                Case 2
                                    sStat = CStr(rst.Fields("GSTATUS"))
                            End Select
                            sSNode = "ie" & rst.Fields("GTYPE") & sStat & "-" & rst.Fields("ELTID")
                            sDesc = sGStatus(rst.Fields("GSTATUS")) ''' & " Graphics"
                            Set nodX = tvwGraphics(0).Nodes.Add(sTNode, tvwChild, sSNode, sDesc, _
                                        iGStatus(rst.Fields("GSTATUS")))
                        End If
                    End If
'''                    iType = rst.Fields("GTYPE")
'''                    Set nodX = tvwGraphics(0).Nodes.Add(sENode, tvwChild, sGNode, sDesc, iType) '''', rst.FIELDS("GTYPE"))
                    rst.MoveNext
                    If rst.EOF Then GoTo DoneLoopin
                Loop
            Else
                sDesc = UCase(Trim(rst.Fields("KITFNAME"))) & " Kit" & strDisc
                lParent = rst.Fields("KITID")
                sKNode = "k" & rst.Fields("KITID")
                Set nodX = tvwGraphics(0).Nodes.Add(, , sKNode, sDesc, 5)
                rst.MoveNext
            End If
            
        Loop
    End If
DoneLoopin:
    rst.Close
    Set rst = Nothing
    
    Set nodX = Nothing
    tvwGraphics(0).Visible = True
End Sub

Public Sub ResizeThePicture()
    Select Case rAsp
        Case Is = rFAsp
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
End Sub

Public Sub GetGraphicList(lCUNO As Long)
    Dim strSelect As String, sStat As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sDesc As String, sCNode As String, sDNode As String, sTNode As String, _
                sGNode As String, sSNode As String, sFNode As String, SHNode As String
    Dim iType As Integer
'''    Dim sGStatus(0 To 30) As String   ''TAKEN OUT TO ELIMINATE SEEING DE-ACTIVATED FILES''
    Dim sGStatus(10 To 30) As String
    Dim iGStatus(0 To 30) As Integer
    
    '///// FILE STATUS VARIABLES \\\\\
'''    sGStatus(0) = "DE-ACTIVATED"   ''TAKEN OUT TO ELIMINATE SEEING DE-ACTIVATED FILES''
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    iGStatus(0) = 10
    iGStatus(10) = 7
    iGStatus(20) = 8
    iGStatus(30) = 9
    
    
    tvwGraphics(3).Visible = False
    tvwGraphics(3).Nodes.Clear
    tvwGraphics(3).ImageList = ImageList1
    sCNode = "": sDNode = "": sTNode = "": sGNode = "": sSNode = ""
    
    
'''    ''ADD TEST 'CLIENT FOLDERS'''
'''    sCNode = "c" & CStr(lCUNO)
'''    sDesc = "Client Folders"
'''    Set nodX = tvwGraphics(3).Nodes.Add(, , sCNode, sDesc, 14)
'''    sFNode = "f1": sDesc = "New A-Kit Proposal"
'''        Set nodX = tvwGraphics(3).Nodes.Add(sCNode, tvwChild, sFNode, sDesc, 14)
'''    sFNode = "f2": sDesc = "Client Picnic Photos"
'''        Set nodX = tvwGraphics(3).Nodes.Add(sCNode, tvwChild, sFNode, sDesc, 14)
'''    sFNode = "f3": sDesc = "C-Kit Damage Photos"
'''        Set nodX = tvwGraphics(3).Nodes.Add(sCNode, tvwChild, sFNode, sDesc, 14)
    
    ''FIRST CHECK FOR CLIENT FOLDERS''
    If bGPJ Then
        strSelect = "SELECT DISTINCT GM.AN8_CUNO, " & _
                    "GM.FLR_ID, GF.FLRDESC, GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                    "FROM ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                    "Where GM.AN8_CUNO = " & lCUNO & " " & _
                    "AND GM.FLR_ID > 0 " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID  = GF.FLR_ID " & _
                    "ORDER BY GF.FLRDESC"
    Else
        strSelect = "SELECT DISTINCT GM.AN8_CUNO, " & _
                    "GM.FLR_ID, GF.FLRDESC, GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                    "FROM ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                    "Where GM.GID > 0 " & _
                    "AND GM.AN8_CUNO = " & lCUNO & " " & _
                    "AND GM.FLR_ID > 0 " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID = GF.FLR_ID " & _
                    "AND GF.CLIENTRESTRICT_FLAG = 0 " & _
                    "ORDER BY GF.FLRDESC"
    End If
    
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        iType = 14
        SHNode = "h0"
        sDesc = "Client Folders"
        Set nodX = tvwGraphics(3).Nodes.Add(, , SHNode, sDesc, iType)
        Do While Not rst.EOF
            iType = rst.Fields("FLAG") + 14
            sFNode = "f" & rst.Fields("FLR_ID")
            sDesc = Trim(rst.Fields("FLRDESC"))
            Set nodX = tvwGraphics(3).Nodes.Add(SHNode, tvwChild, sFNode, sDesc, iType)
            
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    
    
    If bPerm(29) Then
        If iView = 0 Then
            strSelect = "SELECT GM.AN8_CUNO, C.ABALPH, GM.GID, GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
'''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT " & _
                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1 " & _
                        "FROM ANNOTATOR.GFX_MASTER GM " & _
                        "Where GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "ORDER BY Y4, M2"
                        
'''            strSelect = "SELECT DISTINCT GM.AN8_CUNO, C.ABALPH, GM.GTYPE, GM.GSTATUS, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
'''                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
'''                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
'''                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
'''                        "AND GM.GSTATUS IN (" & defsin & ") " & _
'''                        "AND GM.AN8_CUNO = C.ABAN8 " & _
'''                        "ORDER BY Y4, M2, GM.GTYPE"
''''''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
        End If
    Else
        If iView = 0 Then
            strSelect = "SELECT GM.AN8_CUNO, C.ABALPH, GM.GID, GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
'''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE, UPPER(GM.GDESC)"
        Else
            strSelect = "SELECT DISTINCT GM.AN8_CUNO, C.ABALPH, GM.GTYPE, GM.GSTATUS, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MONTH') AS M1, " & _
                        "TO_CHAR(GM.ADDDTTM, 'MM') AS M2, " & _
                        "TO_CHAR(GM.ADDDTTM, 'YYYY') AS Y4 " & _
                        "FROM " & GFXMas & " GM, " & F0101 & " C " & _
                        "WHERE GM.AN8_CUNO = " & lCUNO & " " & _
                        "AND GM.GSTATUS IN (" & defSIN & ") " & _
                        "AND GM.GTYPE <> 3 " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
'''                        "ORDER BY GM.GSTATUS, Y4, M2, GM.GTYPE"
        End If
    End If
        
    
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
'''        Select Case Len(rst.Fields("GSTATUS"))
'''            Case 1
'''                sStat = "0" & CStr(rst.Fields("GSTATUS"))
'''            Case 2
'''                sStat = CStr(rst.Fields("GSTATUS"))
'''        End Select
'''        If sSNode <> "S" & sStat Then
'''            sDesc = sGStatus(rst.Fields("GSTATUS")) & " Files"
'''            sSNode = "S" & sStat
'''            Set nodX = tvwGraphics(3).Nodes.Add(, , sSNode, sDesc, iGStatus(rst.Fields("GSTATUS")))
'''        End If
        
'''        If sDNode <> "D" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") Then
        If sDNode <> "D" & rst.Fields("M2") & rst.Fields("Y4") Then
            sDesc = "Posted:  " & UCase(Trim(rst.Fields("M1"))) & " " & Trim(rst.Fields("Y4"))
            sDNode = "D" & rst.Fields("M2") & rst.Fields("Y4")
'''            sDNode = "D" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4")
            Set nodX = tvwGraphics(3).Nodes.Add(, , sDNode, sDesc, 5)
        End If
        
'''        iType = rst.Fields("GTYPE")
''''''        If sTNode <> "T" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") & _
''''''                        "-" & rst.Fields("GTYPE") Then
'''        If sTNode <> "T" & rst.Fields("M2") & rst.Fields("Y4") & _
'''                        "-" & rst.Fields("GTYPE") Then
'''            sDesc = GfxType(rst.Fields("GTYPE"))
''''''            sTNode = "T" & sStat & "-" & rst.Fields("M2") & rst.Fields("Y4") & _
''''''                        "-" & rst.Fields("GTYPE")
'''            sTNode = "T" & rst.Fields("M2") & rst.Fields("Y4") & _
'''                        "-" & rst.Fields("GTYPE")
'''            Set nodX = tvwGraphics(3).Nodes.Add(sDNode, tvwChild, sTNode, sDesc, iType)
'''        End If
        
'''        If sCNode <> "C" & rst.Fields("AN8_CUNO") Then
'''            sDesc = UCase(Trim(rst.Fields("ABALPH")))
'''            sCNode = "C" & rst.Fields("AN8_CUNO")
'''            Set nodX = tvwGraphics(3).Nodes.Add(, , sCNode, sDesc, 5)
'''        End If
'''
'''        If sDNode <> "D" & rst.Fields("M2") & rst.Fields("Y4") Then
'''            sDesc = UCase(Trim(rst.Fields("M1"))) & " " & Trim(rst.Fields("Y4"))
'''            sDNode = "D" & rst.Fields("M2") & rst.Fields("Y4")
'''            Set nodX = tvwGraphics(3).Nodes.Add(sCNode, tvwChild, sDNode, sDesc, 5)
'''        End If
        
'''        iType = rst.Fields("GTYPE")
'''        If sTNode <> "T" & rst.Fields("GTYPE") & rst.Fields("M2") & rst.Fields("Y4") Then
'''            sDesc = GfxType(rst.Fields("GTYPE"))
'''            sTNode = "T" & rst.Fields("GTYPE") & rst.Fields("M2") & rst.Fields("Y4")
'''            Set nodX = tvwGraphics(3).Nodes.Add(sDNode, tvwChild, sTNode, sDesc, iType)
'''        End If
        
'''        If iView = 0 Then
'''            sGNode = "g" & rst.Fields("GID")
'''            sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
'''            Set nodX = tvwGraphics(3).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, iType) ''' iGStatus(rst.Fields("GSTATUS")))
'''        End If
'''        iType = rst.Fields("GTYPE")
'''        Set nodX = tvwGraphics(3).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, iGStatus(rst.Fields("GSTATUS"))) '' iType)
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    tvwGraphics(3).Visible = True
End Sub

Public Sub LoadClientShows(tmpCUNO As Long, tmpSHYR As Integer, iSort As Integer)
    Dim strSelect As String, strNest As String, sStat As String, sChk As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sDesc As String, sCNode As String, SMNode As String, SHNode As String, _
                sSNode As String, sGNode As String, sCUNO As String, sTNode As String, _
                sANode As String
    Dim lSHCD As Long, lCUNO As Long
    Dim iType As Integer
    Dim sGStatus(0 To 30) As String
    Dim iGStatus(0 To 30) As Integer
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(0) = "DE-ACTIVATED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    iGStatus(0) = 10
    iGStatus(10) = 7
    iGStatus(20) = 8
    iGStatus(30) = 9
    
'''    If bPerm(56) Then ''ABLE TO VIEW DRAFTS''
'''        If bGPJ Then sList = "10, 20, 27, 30" Else sList = "20, 27, 30"
'''    Else ''NOT ABLE TO VIEW DRAFTS''
'''        sList = "20, 27, 30"
'''    End If
    
    tvwGraphics(1).Visible = False
    tvwGraphics(1).Nodes.Clear
    tvwGraphics(1).ImageList = ImageList1
    lCUNO = 0: lSHCD = -1
    
    strNest = "(SELECT GID, GTYPE, GSTATUS, GDESC FROM " & GFXMas & " " & _
                "WHERE GSTATUS IN (" & defSIN & "))"
    Select Case iView
        Case 0
            strSelect = "SELECT CS.CSY56CUNO, C.ABALPH AS CLNM, CS.CSY56SHYR, CS.CSY56SHCD, " & _
                        "SM.SHY56NAMA AS SHNM, GS.SHOW_ID, GM.GDESC, GM.GTYPE, GM.GSTATUS, " & _
                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
                        "FROM " & F0101 & " C, " & F5611 & " CS, " & _
                        "" & F5601 & " SM, " & GFXShow & " GS, " & strNest & " GM " & _
                        "WHERE CS.CSY56SHCD > 0 " & _
                        "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                        "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                        "AND CS.CSY56CUNO = C.ABAN8 " & _
                        "AND C.ABAT1 = 'C' " & _
                        "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                        "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                        "AND CS.CSY56CUNO = GS.AN8_CUNO (+) " & _
                        "AND CS.CSY56SHCD = GS.AN8_SHCD (+) " & _
                        "AND CS.CSY56SHYR = GS.SHYR (+) " & _
                        "AND GS.GID > 0 " & _
                        "AND GS.SHOW_ID > 0 " & _
                        "AND GS.ELTID IS NULL " & _
                        "AND GS.GID = GM.GID (+) " & _
                        "ORDER BY SM.SHY56BEGDT, SHNM, GM.GSTATUS, GDESC"
        Case 1
'''            strSelect = "SELECT DISTINCT CS.CSY56CUNO, C.ABALPH AS CLNM, CS.CSY56SHYR, CS.CSY56SHCD, " & _
'''                        "SM.SHY56NAMA AS SHNM, GM.GTYPE, GM.GSTATUS, SM.SHY56BEGDT, " & _
'''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
'''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
'''                        "FROM " & F0101 & " C, " & F5611 & " CS, " & _
'''                        "" & F5601 & " SM, " & GFXShow & " GS, " & strNest & " GM " & _
'''                        "WHERE CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
'''                        "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
'''                        "AND CS.CSY56CUNO = C.ABAN8 " & _
'''                        "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
'''                        "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
'''                        "AND CS.CSY56CUNO = GS.AN8_CUNO (+) " & _
'''                        "AND CS.CSY56SHCD = GS.AN8_SHCD (+) " & _
'''                        "AND CS.CSY56SHYR = GS.SHYR (+) " & _
'''                        "AND GS.ELTID IS NULL " & _
'''                        "AND GS.GID = GM.GID (+) " & _
'''                        "ORDER BY SM.SHY56BEGDT, SHNM, GTYPE, GM.GSTATUS"
            Select Case iSort
                Case 1
                    strSelect = "SELECT DISTINCT " & _
                                "CS.CSY56SHCD, SM.SHY56NAMA AS SHNM, SM.SHY56BEGDT, " & _
                                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
                                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
                                "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM " & _
                                "Where CS.CSY56SHCD > 0 " & _
                                "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                                "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                                "AND CS.CSY56CUNO = C.ABAN8 " & _
                                "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                                "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                                "ORDER BY SM.SHY56BEGDT, SHNM"
                Case 0
                    strSelect = "SELECT DISTINCT " & _
                                "CS.CSY56SHCD, SM.SHY56NAMA AS SHNM " & _
                                "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM " & _
                                "Where CS.CSY56SHCD > 0 " & _
                                "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                                "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                                "AND CS.CSY56CUNO = C.ABAN8 " & _
                                "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                                "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                                "ORDER BY SHNM"
            End Select
                        
    End Select
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        ''ADD TEST FOR PASSING IN VARS''
        Select Case iSort
            Case 0
                If lSHCD <> rst.Fields("CSY56SHCD") Then
                    sSNode = "s" & rst.Fields("CSY56SHCD")
                    sDesc = UCase(Trim(rst.Fields("SHNM")))
                    Set nodX = tvwGraphics(1).Nodes.Add(, , sSNode, sDesc, 6)
                    lSHCD = CLng(rst.Fields("CSY56SHCD"))
        '            nodX.Bold = True
                End If
            Case 1
                If SMNode <> "m" & Trim(rst.Fields("BEGMON")) Then
                    SMNode = "m" & Trim(rst.Fields("BEGMON"))
                    sDesc = UCase(Trim(rst.Fields("BEGMON"))) & " " & Trim(rst.Fields("BEGYEAR"))
                    Set nodX = tvwGraphics(1).Nodes.Add(, , SMNode, sDesc, 6)
        '            nodX.Bold = True
                End If
        
                If lSHCD <> rst.Fields("CSY56SHCD") Then
                    sSNode = "s" & rst.Fields("CSY56SHCD")
                    sDesc = UCase(Trim(rst.Fields("SHNM")))
                    Set nodX = tvwGraphics(1).Nodes.Add(SMNode, tvwChild, sSNode, sDesc, 6)
                    lSHCD = CLng(rst.Fields("CSY56SHCD"))
        '            nodX.Bold = True
                End If
        End Select
        rst.MoveNext
    Loop
    rst.Close
    
    
    ''NOW, GET THE SHOW-SPECIFIC GRAPHIC NODES''
    sSNode = ""
    Select Case iSort
        Case 0
'            If bGPJ Then
'''                strSelect = "SELECT DISTINCT CS.CSY56SHCD " & _
'''                            "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, ANNOTATOR.GFX_SHOW GS, " & _
'''                            "(SELECT GID, GTYPE, GSTATUS, GDESC " & _
'''                            "From GFX_MASTER " & _
'''                            "WHERE GID > 0 " & _
'''                            "AND GSTATUS IN (20, 27, 30)) GM " & _
'''                            "Where CS.CSY56SHCD > 0 " & _
'''                            "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
'''                            "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
'''                            "AND CS.CSY56CUNO = C.ABAN8 " & _
'''                            "AND C.ABAT1 = 'C' " & _
'''                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
'''                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
'''                            "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
'''                            "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
'''                            "AND CS.CSY56SHYR = GS.SHYR " & _
'''                            "AND GS.GID > 0 " & _
'''                            "AND GS.SHOW_ID > 0 " & _
'''                            "AND GS.ELTID IS NULL " & _
'''                            "AND GS.GID = GM.GID"
                            
                strSelect = "SELECT DISTINCT CS.CSY56SHCD " & _
                            "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, ANNOTATOR.GFX_SHOW GS, " & _
                            "(SELECT GID " & _
                            "From ANNOTATOR.GFX_SHOW " & _
                            "WHERE GID > 0 AND SHOW_ID > 0 " & _
                            "AND AN8_CUNO = " & CLng(tmpCUNO) & ") GM " & _
                            "Where CS.CSY56SHCD > 0 " & _
                            "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                            "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                            "AND CS.CSY56CUNO = C.ABAN8 " & _
                            "AND C.ABAT1 = 'C' " & _
                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                            "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
                            "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
                            "AND CS.CSY56SHYR = GS.SHYR " & _
                            "AND GS.GID > 0 " & _
                            "AND GS.SHOW_ID > 0 " & _
                            "AND GS.ELTID IS NULL " & _
                            "AND GS.GID = GM.GID"
'            Else
'
'            End If
                        
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sSNode <> "s" & rst.Fields("CSY56SHCD") Then
                    sSNode = "s" & rst.Fields("CSY56SHCD")
                    SHNode = "hs" & rst.Fields("CSY56SHCD")
                    sDesc = "Show-Specific Graphics"
                    Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, SHNode, sDesc, 20)
        '            nodX.Bold = True
                    tvwGraphics(1).Nodes(sSNode).Image = 5
'''                    tvwGraphics(1).Nodes(SMNode).Image = 5
                    sTNode = ""
                End If
'''                If sTNode <> "t" & rst.Fields("CSY56SHCD") & "-" & rst.Fields("GTYPE") Then
'''                    sTNode = "t" & rst.Fields("CSY56SHCD") & "-" & rst.Fields("GTYPE")
'''                    sDesc = GfxType(rst.Fields("GTYPE"))
'''                    iType = rst.Fields("GTYPE")
'''                    Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sTNode, sDesc, iType)
'''                End If
                rst.MoveNext
            Loop
            rst.Close
        Case 1
'''            strSelect = "SELECT DISTINCT CS.CSY56SHCD, SM.SHY56BEGDT, GM.GTYPE, " & _
'''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
'''                        "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
'''                        "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, GFX_SHOW GS, " & _
'''                        "(SELECT GID, GTYPE, GSTATUS, GDESC " & _
'''                        "From GFX_MASTER " & _
'''                        "WHERE GSTATUS IN (20, 30)) GM " & _
'''                        "Where CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
'''                        "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
'''                        "AND CS.CSY56CUNO = C.ABAN8 " & _
'''                        "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
'''                        "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
'''                        "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
'''                        "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
'''                        "AND CS.CSY56SHYR = GS.SHYR " & _
'''                        "AND GS.ELTID IS NULL " & _
'''                        "AND GS.GID = GM.GID " & _
'''                        "ORDER BY SM.SHY56BEGDT, GM.GTYPE"
            If bGPJ Then
'''                strSelect = "SELECT DISTINCT CS.CSY56SHCD, SM.SHY56BEGDT, " & _
'''                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
'''                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
'''                            "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, GFX_SHOW GS, " & _
'''                            "(SELECT GID From GFX_MASTER " & _
'''                            "WHERE GID > 0 " & _
'''                            "AND AN8_CUNO = " & CLng(tmpCUNO) & " " & _
'''                            "AND GSTATUS IN (20, 27, 30)) GM " & _
'''                            "Where CS.CSY56SHCD > 0 " & _
'''                            "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
'''                            "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
'''                            "AND CS.CSY56CUNO = C.ABAN8 " & _
'''                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
'''                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
'''                            "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
'''                            "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
'''                            "AND CS.CSY56SHYR = GS.SHYR " & _
'''                            "AND GS.GID > 0 " & _
'''                            "AND GS.SHOW_ID > 0 " & _
'''                            "AND GS.ELTID IS NULL " & _
'''                            "AND GS.GID = GM.GID " & _
'''                            "ORDER BY SM.SHY56BEGDT"
                            
                strSelect = "SELECT DISTINCT CS.CSY56SHCD, SM.SHY56BEGDT, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
                            "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, ANNOTATOR.GFX_SHOW GS, " & _
                            "(SELECT GID From ANNOTATOR.GFX_SHOW " & _
                            "WHERE GID > 0 " & _
                            "AND AN8_CUNO = " & CLng(tmpCUNO) & ") GM " & _
                            "Where CS.CSY56SHCD > 0 " & _
                            "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                            "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                            "AND CS.CSY56CUNO = C.ABAN8 " & _
                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                            "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
                            "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
                            "AND CS.CSY56SHYR = GS.SHYR " & _
                            "AND GS.GID > 0 " & _
                            "AND GS.SHOW_ID > 0 " & _
                            "AND GS.ELTID IS NULL " & _
                            "AND GS.GID = GM.GID " & _
                            "ORDER BY SM.SHY56BEGDT"
                            
            Else
                strSelect = "SELECT DISTINCT CS.CSY56SHCD, SM.SHY56BEGDT, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'YYYY') AS BEGYEAR " & _
                            "FROM " & F0101 & " C, " & F5611 & " CS, " & F5601 & " SM, ANNOTATOR.GFX_SHOW GS, " & _
                            "(SELECT GM.GID From ANNOTATOR.GFX_MASTER GM " & _
                            "Where GM.GID > 0 " & _
                            "AND GM.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                            "AND GM.GSTATUS IN (20, 27, 30) MINUS " & _
                            "SELECT GM.GID From ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                            "Where GM.GID > 0 AND GM.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                            "AND GM.GSTATUS IN (20, 27, 30) " & _
                            "AND GM.FLR_ID = GF.FLR_ID " & _
                            "AND GF.CLIENTRESTRICT_FLAG = 1) GM " & _
                            "Where CS.CSY56SHCD > 0 " & _
                            "AND CS.CSY56CUNO = " & CLng(tmpCUNO) & " " & _
                            "AND CS.CSY56SHYR = " & tmpSHYR & " " & _
                            "AND CS.CSY56CUNO = C.ABAN8 " & _
                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                            "AND CS.CSY56CUNO = GS.AN8_CUNO " & _
                            "AND CS.CSY56SHCD = GS.AN8_SHCD " & _
                            "AND CS.CSY56SHYR = GS.SHYR " & _
                            "AND GS.GID > 0 AND GS.SHOW_ID > 0 " & _
                            "AND GS.ELTID IS NULL AND GS.GID = GM.GID " & _
                            "ORDER BY SM.SHY56BEGDT"
            End If
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If SHNode <> "hs" & rst.Fields("CSY56SHCD") Then
                    SMNode = "m" & Trim(rst.Fields("BEGMON"))
                    sSNode = "s" & rst.Fields("CSY56SHCD")
                    SHNode = "hs" & rst.Fields("CSY56SHCD")
                    sDesc = "Show-Specific Graphics"
                    Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, SHNode, sDesc, 20)
        '            nodX.Bold = True
                    tvwGraphics(1).Nodes(sSNode).Image = 5
                    tvwGraphics(1).Nodes(SMNode).Image = 5
                    
'''                    sANode = "ag" & rst.Fields("CSY56SHCD")
'''                    sDesc = "All Graphics"
'''                    Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sANode, sDesc, 14)
                    
                    sTNode = ""
                End If
'''                If sTNode <> "t" & rst.Fields("CSY56SHCD") & "-" & rst.Fields("GTYPE") Then
'''                    sTNode = "t" & rst.Fields("CSY56SHCD") & "-" & rst.Fields("GTYPE")
'''                    sDesc = GfxType(rst.Fields("GTYPE"))
'''                    iType = rst.Fields("GTYPE")
'''                    Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sTNode, sDesc, iType)
'''                End If
                rst.MoveNext
            Loop
            rst.Close
    End Select
    
    ''CHECK FOR SHOW-SPECIFIC FOLDERS''
    If bGPJ Then
        strSelect = "SELECT DISTINCT GS.AN8_SHCD, GF.FLR_ID, GF.FLRDESC, " & _
                    "GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                    "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                    "Where GS.GID > 0 " & _
                    "AND GS.SHOW_ID > 0 " & _
                    "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                    "AND GS.SHYR = " & tmpSHYR & " " & _
                    "AND GS.GID = GM.GID " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID > 0 " & _
                    "AND GM.FLR_ID = GF.FLR_ID " & _
                    "AND GF.AN8_CUNO = " & CLng(tmpCUNO)
    Else
        strSelect = "SELECT DISTINCT GS.AN8_SHCD, GF.FLR_ID, GF.FLRDESC, " & _
                    "GF.CLIENTRESTRICT_FLAG AS FLAG " & _
                    "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_FOLDER GF " & _
                    "Where GS.GID > 0 " & _
                    "AND GS.SHOW_ID > 0 " & _
                    "AND GS.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                    "AND GS.SHYR = " & tmpSHYR & " " & _
                    "AND GS.GID = GM.GID " & _
                    "AND GM.GSTATUS IN (" & defSIN & ") " & _
                    "AND GM.FLR_ID > 0 " & _
                    "AND GM.FLR_ID = GF.FLR_ID " & _
                    "AND GF.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                    "AND GF.CLIENTRESTRICT_FLAG = 0"
    End If
    Set rst = Conn.Execute(strSelect)
    On Error Resume Next
    Do While Not rst.EOF
'        If iSort = 0 Then
'            SHNode = "s" & rst.Fields("AN8_SHCD")
'        Else
            SHNode = "hs" & rst.Fields("AN8_SHCD")
'        End If
        
        Err = 0
        sChk = tvwGraphics(1).Nodes(SHNode).Text
        If Err = 0 Then
            sANode = "f" & rst.Fields("AN8_SHCD") & "-" & rst.Fields("FLR_ID")
            sDesc = Trim(rst.Fields("FLRDESC"))
            iType = rst.Fields("FLAG") + 14
            Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sANode, sDesc, iType)
        End If
        rst.MoveNext
    Loop
    rst.Close
    On Error GoTo 0
    
    ''CHECK FOR ELEMENT GRAPHICS''
    ''FUTURE SHOWS FIRST''
    strSelect = "SELECT DISTINCT EU.AN8_SHCD " & _
                "FROM " & AQUAEltU & " EU, ANNOTATOR.GFX_ELEMENT GE, " & F5601 & " SM " & _
                "Where EU.AN8_SHCD > 0 " & _
                "AND EU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                "AND EU.SHYR = " & tmpSHYR & " " & _
                "AND EU.KITUSEID > 0 " & _
                "AND EU.ELTUSEID > 0 " & _
                "AND EU.SHSTATUS IN (1, 4) " & _
                "AND EU.ELTID = GE.ELTID " & _
                "AND GE.GID > 0 " & _
                "AND EU.SHYR = SM.SHY56SHYR " & _
                "AND EU.AN8_SHCD = SM.SHY56SHCD " & _
                "AND SM.SHY56ENDDT > " & IGLToJDEDate(Now)

    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        sSNode = "s" & rst.Fields("AN8_SHCD")
        SHNode = "he" & rst.Fields("AN8_SHCD")
        sANode = "ae" & rst.Fields("AN8_SHCD")
        sDesc = "Assigned Element Graphics"
        Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, SHNode, sDesc, 19)
'        nodX.Bold = True
        tvwGraphics(1).Nodes(sSNode).Image = 5
        If iSSSort = 1 Then tvwGraphics(1).Nodes(sSNode).Parent.Image = 5
        sDesc = "<Click to display list>"
        Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sANode, sDesc, 12)
        
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    ''NOW, CLOSED SHOWS''
    strSelect = "SELECT DISTINCT EU.AN8_SHCD " & _
                "FROM " & AQUAEltU & " EU, ANNOTATOR.GFX_SHOW GS, " & F5601 & " SM " & _
                "Where EU.AN8_SHCD > 0 " & _
                "AND EU.AN8_CUNO = " & CLng(tmpCUNO) & " " & _
                "AND EU.SHYR = " & tmpSHYR & " " & _
                "AND EU.KITUSEID > 0 " & _
                "AND EU.ELTUSEID > 0 " & _
                "AND EU.SHSTATUS IN (1, 4) " & _
                "AND EU.ELTID = GS.ELTID " & _
                "AND EU.AN8_SHCD = GS.AN8_SHCD " & _
                "AND EU.SHYR = GS.SHYR " & _
                "AND EU.AN8_CUNO = GS.AN8_CUNO " & _
                "AND GS.GID > 0 " & _
                "AND GS.SHOW_ID > 0 " & _
                "AND EU.SHYR = SM.SHY56SHYR " & _
                "AND EU.AN8_SHCD = SM.SHY56SHCD " & _
                "AND SM.SHY56ENDDT <= " & IGLToJDEDate(Now)

    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        sSNode = "s" & rst.Fields("AN8_SHCD")
        SHNode = "he" & rst.Fields("AN8_SHCD")
        sANode = "ar" & rst.Fields("AN8_SHCD")
        sDesc = "Archived Element Graphics"
        Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, SHNode, sDesc, 19)
'        nodX.Bold = True
        tvwGraphics(1).Nodes(sSNode).Image = 5
        If iSSSort = 1 Then tvwGraphics(1).Nodes(sSNode).Parent.Image = 5
        sDesc = "<Click to display list>"
        Set nodX = tvwGraphics(1).Nodes.Add(SHNode, tvwChild, sANode, sDesc, 12)
        
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    
    
    
'''        Select Case iView
'''            Case 0
'''                If rst.Fields("GDESC") <> "" Then
'''                    Select Case rst.Fields("GSTATUS")
'''                        Case Is > 0: sGNode = "ga" & rst.Fields("SHOW_ID")
'''                        Case Else: sGNode = "gi" & rst.Fields("SHOW_ID")
'''                    End Select
'''                    sDesc = Trim(rst.Fields("GDESC")) & "  [" & sGStatus(rst.Fields("GSTATUS")) & "]"
'''                    iType = rst.Fields("GTYPE")
'''                    If iType = 3 And bPerm(29) = False Then GoTo SkipThisOne
'''                    Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, sGNode, sDesc, iType)
'''                    tvwGraphics(1).Nodes(sSNode).Image = 5
'''                    tvwGraphics(1).Nodes(SMNode).Image = 5
'''                    tvwGraphics(1).Nodes(sCNode).Image = 5
'''SkipThisOne:
'''                End If
'''            Case 1
'''                If Not IsNull(rst.Fields("GTYPE")) Then
'''                    iType = rst.Fields("GTYPE")
'''                    If sTNode <> "i" & iType & lSHCD Then
'''                        If iType = 3 And bPerm(29) = False Then GoTo SkipThisOne2
'''                        sTNode = "i" & iType & lSHCD
'''                        sDesc = GfxType(iType)
'''                        Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, sTNode, sDesc, iType)
'''                        tvwGraphics(1).Nodes(sSNode).Image = 5
'''                        tvwGraphics(1).Nodes(SMNode).Image = 5
''''                        tvwGraphics(1).Nodes(sCNode).Image = 5
'''''                        tvwGraphics(1).Nodes(sSNode).Bold = True
'''''                        tvwGraphics(1).Nodes(SMNode).Bold = True
'''                    End If
'''                    Select Case Len(rst.Fields("GSTATUS"))
'''                        Case 1
'''                            sStat = "0" & CStr(rst.Fields("GSTATUS"))
'''                        Case 2
'''                            sStat = CStr(rst.Fields("GSTATUS"))
'''                    End Select
'''                    sGNode = "i" & iType & sStat & "-" & lSHCD
'''                    sDesc = sGStatus(rst.Fields("GSTATUS")) '''& " Graphics"
'''                    Set nodX = tvwGraphics(1).Nodes.Add(sTNode, tvwChild, sGNode, sDesc, _
'''                                iGStatus(rst.Fields("GSTATUS")))
'''
''''''                If Not IsNull(rst.Fields("GTYPE")) Then
''''''                    iType = rst.Fields("GTYPE")
''''''                    sGNode = "i" & iType & lSHCD
''''''                    sDesc = GfxType(iType)
''''''                    If iType = 3 And bPerm(29) = False Then GoTo SkipThisOne2
''''''                    Set nodX = tvwGraphics(1).Nodes.Add(sSNode, tvwChild, sGNode, sDesc, iType)
''''''                    tvwGraphics(1).Nodes(sSNode).Image = 5
''''''                    tvwGraphics(1).Nodes(SMNode).Image = 5
''''''                    tvwGraphics(1).Nodes(sCNode).Image = 5
'''SkipThisOne2:
'''                End If
'''        End Select
'''
'''        rst.MoveNext
'''    Loop
'''    rst.Close
'''    Set rst = Nothing
    tvwGraphics(1).Visible = True
End Sub

Public Sub PopClientsWithGraphics(combo As ComboBox, Tree As TreeView)
    Dim strSelect As String, sClient As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    If combo.Text <> "" Then sClient = combo.Text
    combo.Clear
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT G.AN8_CUNO, C.ABALPH " & _
                    "FROM " & GFXMas & " G, " & F0101 & " C " & _
                    "WHERE G.GID > 0 " & _
                    "AND G.AN8_CUNO <> 40579 " & _
                    "AND G.AN8_CUNO = C.ABAN8 " & _
                    "AND G.GSTATUS > 0 " & _
                    "ORDER BY UPPER(C.ABALPH)"
    Else
        strSelect = "SELECT DISTINCT G.AN8_CUNO, C.ABALPH " & _
                "FROM " & GFXMas & " G, " & F0101 & " C " & _
                "WHERE G.GID > 0 " & _
                "AND G.AN8_CUNO IN (" & strCunoList & ") " & _
                "AND G.GSTATUS > 0 " & _
                "AND G.AN8_CUNO = C.ABAN8 " & _
                "ORDER BY UPPER(C.ABALPH)"
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        combo.AddItem UCase(Trim(rst.Fields("ABALPH")))
        combo.ItemData(combo.NewIndex) = rst.Fields("AN8_CUNO")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    Err.Clear
    On Error Resume Next
    For i = 0 To combo.ListCount - 1
        If combo.List(i) = sClient Then
            combo.Text = combo.List(i)
            GoTo FoundClient
        End If
    Next i
    Tree.Nodes.Clear
FoundClient:
End Sub

Public Function GetBCN(tmpBCC As String)
    Dim rstCN As ADODB.Recordset
    Dim strSelect As String
    
    strSelect = "SELECT ABALPH FROM " & F0101 & " " & _
                "WHERE ABAN8 = " & CLng(tmpBCC)
    Set rstCN = Conn.Execute(strSelect)
    If Not rstCN.EOF Then
        GetBCN = UCase(Trim(rstCN.Fields("ABALPH")))
    Else
        GetBCN = ""
    End If
    rstCN.Close
    Set rstCN = Nothing
End Function

Public Function GetSHNM(tmpSHCD As Long, tmpSHYR As Integer)
    Dim rstSN As ADODB.Recordset
    Dim strSelect As String
    
    '****'Conn.Open ALREADY****
    strSelect = "SELECT SHY56NAMA FROM " & F5601 & " " & _
                "WHERE SHY56SHCD = " & tmpSHCD & " " & _
                "AND TSHY56SHYR = " & tmpSHYR
    Set rstSN = Conn.Execute(strSelect)
    If Not rstSN.EOF Then
        GetSHNM = UCase(Trim(rstSN.Fields("SHY56NAMA")))
    Else
        GetSHNM = ""
    End If
    rstSN.Close
    Set rstSN = Nothing
End Function

Public Sub CheckShows(cNode As String, CNodeTxt As String, iFrom As Integer)
    Dim strSelect As String, sCUNO As String, sMess As String, strGID As String
    Dim rst As ADODB.Recordset 'IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MONTH') AS BEGMON
    Dim lCUNO As Long
    
    sMess = ""
    lCUNO = 0
    If UCase(Left(cNode, 1)) = "G" Then
        Select Case iFrom
            Case 0
                strGID = "SELECT GID FROM " & GFXShow & " " & _
                            "WHERE SHOW_ID = " & CLng(Mid(cNode, 3))
                strSelect = "SELECT GS.AN8_CUNO, C.ABALPH AS CLNM, GS.SHYR, " & _
                            "GS.AN8_SHCD, SM.SHY56NAMA AS SHNM, SM.SHY56BEGDT, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MON DD, YYYY') AS BEGD " & _
                            "FROM " & GFXShow & " GS, " & F0101 & " C, " & F5601 & " SM " & _
                            "WHERE GS.GID IN (" & strGID & ") " & _
                            "AND GS.AN8_CUNO = C.ABAN8 " & _
                            "AND GS.SHYR = SM.SHY56SHYR " & _
                            "AND GS.AN8_SHCD = SM.SHY56SHCD " & _
                            "ORDER BY SM.SHY56BEGDT"
            Case 1
                strSelect = "SELECT GS.AN8_CUNO, C.ABALPH AS CLNM, GS.SHYR, " & _
                            "GS.AN8_SHCD, SM.SHY56NAMA AS SHNM, SM.SHY56BEGDT, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MON DD, YYYY') AS BEGD " & _
                            "FROM " & GFXShow & " GS, " & F0101 & " C, " & F5601 & " SM " & _
                            "WHERE GS.GID = " & CLng(Mid(cNode, 2)) & " " & _
                            "AND GS.AN8_CUNO = C.ABAN8 " & _
                            "AND GS.SHYR = SM.SHY56SHYR " & _
                            "AND GS.AN8_SHCD = SM.SHY56SHCD " & _
                            "ORDER BY SM.SHY56BEGDT"
                
                strSelect = "SELECT GS.AN8_CUNO, C.ABALPH AS CLNM, GS.SHYR, " & _
                            "GS.AN8_SHCD, SM.SHY56NAMA AS SHNM, SM.SHY56BEGDT, " & _
                            "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MON DD, YYYY') AS BEGD " & _
                            "FROM " & GFXShow & " GS, " & F0101 & " C, " & F5601 & " SM " & _
                            "WHERE GS.GID IN (" & _
                            "SELECT GID FROM " & GFXShow & " WHERE SHOW_ID = " & CLng(Mid(cNode, 2)) & ") " & _
                            "AND GS.AN8_CUNO = C.ABAN8 " & _
                            "AND GS.SHYR = SM.SHY56SHYR " & _
                            "AND GS.AN8_SHCD = SM.SHY56SHCD " & _
                            "ORDER BY SM.SHY56BEGDT"
        End Select
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            If lCUNO <> rst.Fields("AN8_CUNO") Then
                sMess = sMess & vbNewLine & Space(4) & UCase(Trim(rst.Fields("CLNM"))) & vbNewLine
'''                sCUNO = Right("00000000" & rst.FIELDS("AN8_CUNO"), 8)
                lCUNO = rst.Fields("AN8_CUNO")
            End If
            sMess = sMess & Space(8) & rst.Fields("SHYR") & " - " & _
                        UCase(Trim(rst.Fields("SHNM"))) & " [" & Trim(rst.Fields("BEGD")) & "]" & vbNewLine
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
        If Len(sMess) > 0 Then
            With frmUsage
                .PassMess = sMess
                .PassTitle = "Show Assignments:  " & Trim(CNodeTxt)
                .Show 1
            End With
        Else
            MsgBox "The selected graphic has not been assigned to any Shows.", vbInformation, _
                        "Show Assignments:  " & Trim(CNodeTxt)
        End If
    End If
End Sub

Public Sub ClearChecks()
    mnuGRedSketch.Checked = False
    mnuGRedText.Checked = False
End Sub

Private Sub txt1_Change()
    lblRedNote.Caption = txt1.Text
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt1.Text <> "" Then
'''''            lblRedNote.BorderStyle = 0
'''''            picJPG.FontSize = 14
'''''            picJPG.ForeColor = vbRed
'''''            picJPG.CurrentX = xStr
'''''            picJPG.CurrentY = lblRedNote.Top
            mnuGRedSave.Enabled = True
            mnuGRedClear.Enabled = True
            Call UpdateSCD(1, 1, 0)
'''            lblRedNote.Visible = False
'''            picJPG.Print lblRedNote.Caption
            txt1.Text = ""
            lblRedNote.Caption = ""
        Else
            lblRedNote.BorderStyle = 0
            lblRedNote.Visible = False
            txt1.Text = ""
            lblRedNote.Caption = ""
        End If
    End If
End Sub

Private Sub txt1_LostFocus()
    Dim Resp As VbMsgBoxResult
    
    Debug.Print "Lost focus"
'''    If lblRedNote.BorderStyle = 1 Then
        If txt1.Text <> "" Then
            Resp = MsgBox("Do you wish to write the Text comment to the screen?", _
                        vbYesNo, "Comment in Process...")
            If Resp = vbYes Then
                lblRedNote.BorderStyle = 0
                picJPG.FontSize = 14
                picJPG.ForeColor = vbRed
                picJPG.CurrentX = xStr
                picJPG.CurrentY = lblRedNote.Top
                mnuGRedSave.Enabled = True
                mnuGRedClear.Enabled = True
                lblRedNote.Visible = False
                picJPG.Print lblRedNote.Caption
                txt1.Text = ""
                lblRedNote.Caption = ""
            Else
                lblRedNote.BorderStyle = 0
                lblRedNote.Visible = False
                txt1.Text = ""
                lblRedNote.Caption = ""
            End If
'''        Else
'''            lblRedNote.BorderStyle = 0
'''            lblRedNote.Visible = False
'''            txt1.Text = ""
'''            lblRedNote.Caption = ""
'''        End If
    
    End If
    
End Sub

Public Sub ResetControls(iNew As Integer)
    Select Case iNew
        Case 0
            sst1.Width = 6090
            picTabs.Width = sst1.Width
            '/// TAB 0 \\\
            tvwGraphics(0).Width = 5715: tvwGraphics(0).Height = 4995: tvwGraphics(0).Top = 1200
            fra0.Width = 5970
            lblShow.Left = 120: lblShow.Top = 540
'''''            If bPerm(24) Or bPerm(25) Then
'''''                cmdAssign.Left = 4560: cmdAssign.Top = 210: cmdAssign.Height = 855: cmdAssign.Width = 1275
'''''                cmdAssign.Visible = True
'''''                cboCUNO(0).Width = 3375
'''''                cboSHCD.Width = 4335
'''''            Else
'''''                cmdAssign.Visible = False
'''''                cboCUNO(0).Width = 4755
'''''                cboSHCD.Width = 5715
'''''            End If
            cboSHCD.Left = 120: cboSHCD.Top = 750
'''            cmdAssign.Left = 4560: cmdAssign.Top = 210: cmdAssign.Height = 855: cmdAssign.Width = 1275
            picOuter(0).Visible = False
            hsc1(0).Visible = False
            lblInactive(0).Visible = False
            '/// TAB 1 \\\
            tvwGraphics(1).Width = 5715
            cboCUNO(1).Width = 4755
            picOuter(1).Visible = False
            hsc1(1).Visible = False
            lblInactive(1).Visible = False
            '/// TAB 2 \\\
            tvwGraphics(2).Width = 5715
            cboCUNO(2).Width = 5715
            picOuter(2).Visible = False
            hsc1(2).Visible = False
            lblInactive(2).Visible = False
            '/// TAB 3 \\\
            tvwGraphics(3).Width = 5715
            cboCUNO(3).Width = 5715
            picOuter(3).Visible = False
            hsc1(3).Visible = False
            lblInactive(3).Visible = False
'''            picReview.Visible = False
'''            sst1.TabCaption(3) = "Client Graphics"
        Case 1
            sst1.Width = Me.ScaleWidth - 240 ''' 11550
            picTabs.Width = sst1.Width
            '/// TAB 0 \\\
            tvwGraphics(0).Width = 4335: tvwGraphics(0).Height = 5535: tvwGraphics(0).Top = 660
            fra0.Width = 11415
            lblShow.Left = 4560: lblShow.Top = 0
            cboCUNO(0).Width = 3375
            cboSHCD.Width = 4335: cboSHCD.Left = 4560: cboSHCD.Top = 210
'''''            If bPerm(24) Or bPerm(25) Then
'''''                cmdAssign.Left = 9000: cmdAssign.Top = 120: cmdAssign.Height = 435: cmdAssign.Width = 2295
'''''                cmdAssign.Visible = True
'''''            Else
'''''                cmdAssign.Visible = False
'''''            End If
            picOuter(0).Visible = True
            '/// TAB 1 \\\
            tvwGraphics(1).Width = 4335
            cboCUNO(1).Width = 3375
            picOuter(1).Visible = True
            '/// TAB 2 \\\
            tvwGraphics(2).Width = 4335
            cboCUNO(2).Width = 4335
            picOuter(2).Visible = True
            '/// TAB 3 \\\
            tvwGraphics(3).Width = 4335
            cboCUNO(3).Width = 4335
            picOuter(3).Visible = True
'''''            If bGFXReviewer Then
'''''                picReview.Visible = True
''''''                sst1.TabCaption(3) = "Client Graphics / Approval Interface"
'''''            End If
    End Select
End Sub

Public Sub LoadGraphic(Index As Integer, NodeKey As String, NodeText As String, _
            NodeParKey As String, NodeParText As String)
    Dim strSelect As String, strInsert As String, strUpdate As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim i As Integer, iLock As Integer, iCol As Integer, iRow As Integer, tIndex As Integer
    Dim sGStatus(0 To 30) As String
    Dim bCheckVersion As Boolean
    
    sGStatus(0) = "DE-ACTIVATED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED" '' FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
'''    picGfxApprove.Visible = False
'''    cmdGfxApproveHide.Visible = False
    
    Debug.Print NodeKey & " - YOU GOT ONE!"
    CurrNode = NodeKey
    CurrNodeText = NodeText
    Select Case Index
        Case 0 '/// CURRENT SHOW GRAPHICS \\\
            If UCase(Left(NodeKey, 2)) = "GS" Then
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                            "WHERE GS.SHOW_ID = " & CLng(Mid(NodeKey, 3)) & " " & _
                            "AND GS.GID = GM.GID"
            Else
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                            "WHERE GE.ES_ID = " & CLng(Mid(NodeKey, 3)) & " " & _
                            "AND GE.GID = GM.GID"
            End If
            sTable = "ANNOTATOR.GFX_SHOW"
            lRedID = Mid(NodeKey, 3)
            RedName = Mid(NodeKey, 3) & "-1RED.bmp"
        Case 1 '/// ENTIRE SHOW SEASON GRAPHICS \\\
            Debug.Print "Parent Node Key = " & NodeParKey
            If UCase(Mid(NodeParKey, 2, 1)) = "S" Then
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                            "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                            "WHERE GS.SHOW_ID = " & CLng(Mid(NodeKey, 3)) & " " & _
                            "AND GS.GID = GM.GID"
                sTable = "ANNOTATOR.GFX_SHOW"
            Else
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                            "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                            "WHERE GE.ES_ID = " & CLng(NodeKey) & " " & _
                            "AND GE.GID = GM.GID"
                sTable = "ANNOTATOR.GFX_ELEMENT"
            End If
            lRedID = Mid(NodeKey, 3)
            RedName = Mid(NodeKey, 3) & "-1RED.bmp"
        Case 2 '/// KIT/ELEMENT BASED GRAPHICS \\\
            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                        "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                        "WHERE GE.ES_ID = " & CLng(Mid(NodeKey, 2)) & " " & _
                        "AND GE.GID = GM.GID"
            sTable = "ANNOTATOR.GFX_ELEMENT"
            lRedID = Mid(NodeKey, 2)
            RedName = Mid(NodeKey, 2) & "-1RED.bmp"
        Case 3 '/// CLIENT DATABASE OF GRAPHICS \\\
            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                        "FROM " & GFXMas & " GM " & _
                        "WHERE GM.GID = " & CLng(Mid(NodeKey, 2))
            sTable = "ANNOTATOR.GFX_MASTER"
            lRedID = Mid(NodeKey, 2)
            RedName = Mid(NodeKey, 2) & "-1RED.bmp"
        Case 10
            Debug.Print "Parent Node Key = " & NodeParKey
            If UCase(Left(NodeKey, 1)) = "S" Then
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                            "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXMas & " GM, " & GFXShow & " GS " & _
                            "WHERE GS.SHOW_ID = " & CLng(Mid(NodeKey, 2)) & " " & _
                            "AND GS.GID = GM.GID"
                sTable = "ANNOTATOR.GFX_SHOW"
            ElseIf UCase(Left(NodeKey, 1)) = "E" Then
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                            "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXMas & " GM, " & GFXElt & " GE " & _
                            "WHERE GE.ES_ID = " & CLng(Mid(NodeKey, 2)) & " " & _
                            "AND GE.GID = GM.GID"
                sTable = "ANNOTATOR.GFX_ELEMENT"
            End If
            lRedID = CLng(Mid(NodeKey, 2))
            RedName = CStr(Mid(NodeKey, 2)) & "-1RED.bmp"
        Case 11
            Debug.Print "Parent Node Key = " & NodeParKey
            If UCase(Mid(NodeParKey, 2, 1)) = "S" Then
                strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                            "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                            "WHERE GS.SHOW_ID = " & CLng(NodeKey) & " " & _
                            "AND GS.GID = GM.GID"
                sTable = "ANNOTATOR.GFX_SHOW"
            Else
                Select Case UCase(Left(NodeParText, 3))
                    Case "ASS"
                        strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                                    "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                                    "FROM " & GFXElt & " GE, " & GFXMas & " GM " & _
                                    "WHERE GE.ES_ID = " & CLng(NodeKey) & " " & _
                                    "AND GE.GID = GM.GID"
                        sTable = "ANNOTATOR.GFX_ELEMENT"
                    Case "ARC"
                        strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                                    "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                                    "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                                    "WHERE GS.SHOW_ID = " & CLng(NodeKey) & " " & _
                                    "AND GS.GID = GM.GID"
                        sTable = "ANNOTATOR.GFX_SHOW"
                End Select
            End If
'''            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
'''                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
'''                        "FROM " & GFXMas & " GM, " & GFXShow & " GS " & _
'''                        "WHERE GS.SHOW_ID = " & CLng(NodeKey) & " " & _
'''                        "AND GS.GID = GM.GID"
'''            sTable = "GFX_SHOW"
            lRedID = CLng(NodeKey)
            RedName = CStr(NodeKey) & "-1RED.bmp"
        Case 12
            Debug.Print "Parent Node Key = " & NodeParKey
            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
                        "FROM " & GFXMas & " GM, " & GFXElt & " GE " & _
                        "WHERE GE.ES_ID = " & CLng(NodeKey) & " " & _
                        "AND GE.GID = GM.GID"
            sTable = "ANNOTATOR.GFX_ELEMENT"
            lRedID = CLng(NodeKey)
            RedName = CStr(NodeKey) & "-1RED.bmp"
        Case 13
'''            Debug.Print "Parent Node Key = " & NodeParKey
'''            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
'''                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
'''                        "FROM " & GFXMas & " GM " & _
'''                        "WHERE GM.GID = " & CLng(NodeKey)
                        
            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.GTYPE, GM.GSTATUS, GM.SUPDOC_ID, " & _
                        "TO_CHAR(GM.UPDDTTM, 'DD-MON-YYYY') AS STATDATE " & _
                        "FROM " & GFXMas & " GM " & _
                        "WHERE GM.GID = " & CLng(NodeKey)
            sTable = "ANNOTATOR.GFX_MASTER"
            lRedID = CLng(NodeKey)
            RedName = CStr(NodeKey) & "-1RED.bmp"
        Case 14
'''            Debug.Print "Parent Node Key = " & NodeParKey
'            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.SUPDOC_ID, " & _
'                        "GM.GTYPE, GM.GSTATUS, GM.VERSION_ID, GV.V_FORMAT, " & _
'                        "(TRIM(TO_CHAR(GM.UPDDTTM, 'MONTH')) || ' ' || TO_CHAR(GM.UPDDTTM, 'DD, YYYY')) STATDATE " & _
'                        "FROM " & GFXMas & " GM, GFX_VERSION GV " & _
'                        "WHERE GM.GID = " & CLng(NodeKey) & " " & _
'                        "AND GM.VERSION_ID = GV.VERSION_ID (+)"
                        
            strSelect = "SELECT GM.GID, GM.GDESC, GM.GPATH, GM.GFORMAT, GM.SUPDOC_ID, " & _
                        "GM.GTYPE, GM.GSTATUS, GM.VERSION_ID, GV.V_FORMAT, " & _
                        "TO_CHAR(GM.UPDDTTM, 'DD-MON-YYYY') AS STATDATE " & _
                        "FROM " & GFXMas & " GM, ANNOTATOR.GFX_VERSION GV " & _
                        "WHERE GM.GID = " & CLng(NodeKey) & " " & _
                        "AND GM.VERSION_ID = GV.VERSION_ID (+)"
            sTable = "ANNOTATOR.GFX_MASTER"
            lRedID = CLng(NodeKey)
            RedName = CStr(NodeKey) & "-1RED.bmp"
            bCheckVersion = True
    End Select
    lRefID = lRedID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If Dir(Trim(rst.Fields("GPATH")), vbNormal) <> "" Then
            If Index >= 10 Then tIndex = Index - 10 Else tIndex = Index
            If chkClose(tIndex).Value = 1 Then
                picTabs.Visible = False ''.Visible = False
                bDirsOpen = False
                imgDirs.ToolTipText = "Click to Open File Index..."
'''                Set imgDirs.Picture = imlDirs.ListImages(1).Picture
            End If
            
            tBCC = fBCC(sst1.Tab)
            tFBCN = fFBCN(sst1.Tab)
            tSHYR = fSHYR(sst1.Tab)
            tSHCD = fSHCD(sst1.Tab)
            tSHNM = fSHNM(sst1.Tab)
            
            lGID = rst.Fields("GID")
            
            sCGDesc = Trim(rst.Fields("GDESC"))
            iCurrGType = rst.Fields("GTYPE")
            
'''            If iCurrGType = 3 Or iCurrGType = 2 Then mnuRedlining.Enabled = True Else mnuRedlining.Enabled = False
            mnuRedlining.Enabled = True
            
'''            If iCurrGType = 3 And bPerm(39) Then mnuGPrint(1).Visible = True Else mnuGPrint(1).Visible = False
'''            If bPerm(39) Then mnuGPrint(1).Visible = True Else
            mnuGPrint(1).Visible = False
            
            If bCheckVersion Then
                If rst.Fields("VERSION_ID") > 0 Then
                    CurrFile = sVPath & rst.Fields("VERSION_ID") & "." & Trim(rst.Fields("V_FORMAT"))
                Else
                    CurrFile = rst.Fields("GPATH")
                End If
            Else
                CurrFile = rst.Fields("GPATH")
            End If
            
            If rst.Fields("SUPDOC_ID") > 0 Then
                imgSupDoc.Visible = True
                imgSupDoc.Tag = CStr(rst.Fields("SUPDOC_ID"))
            Else
                imgSupDoc.Visible = False
                imgSupDoc.Tag = ""
            End If
            
            lblStatus.Caption = "STATUS:  " & sGStatus(rst.Fields("GSTATUS")) & " (Last Status Update " & _
                        Trim(rst.Fields("STATDATE")) & ")"
            lblStatus.Visible = True
            
            '///// CLOSE ENTRY IN LOCKLOG \\\\\
            If lNewLockId <> 0 Then
                strUpdate = "UPDATE " & ANOLockLog & " " & _
                            "SET LOCKCLOSEDTTM = SYSDATE, " & _
                            "LOCKSTATUS = LOCKSTATUS * -1, " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, " & _
                            "UPDCNT = UPDCNT + 1 " & _
                            "WHERE LOCKID = " & lNewLockId
                Conn.Execute (strUpdate)
            End If
            
            Call ClearUndo(0)
            lUndoGID = lGID
            
            '///// TIME TO LOAD THE GRAPHIC \\\\\
            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
                imgRedReload.Picture = imlRedMode.ListImages(6).Picture
                Call LoadPDF(CurrFile) ''Call LoadThePDF(CurrFile)
'                Call AddToUndo
            Else
                imgRedReload.Picture = imlRedMode.ListImages(5).Picture
                Call LoadThePicture(CurrFile, False)
'                Call AddToUndo
            End If
            picTools.Visible = True
            
            
            
            '///// ADD ENTRY TO LOCKLOG \\\\\
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lNewLockId = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            
            If bPerm(39) Then iLock = 2 Else iLock = 1
            
            strInsert = "INSERT INTO " & ANOLockLog & " " & _
                        "(LOCKID, LOCKREFID, LOCKREFSOURCE, " & _
                        "USER_SEQ_ID, LOCKOPENDTTM, LOCKSTATUS, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lNewLockId & ", " & lGID & ", 'GFX_MASTER', " & _
                        UserID & ", SYSDATE, " & iLock & ", " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            
            Select Case Index
                Case Is >= 10
                    If Index = 10 Then
                        lblWelcome = tFBCN & " " & sTabDesc & ":   " & NodeText
'''                        lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    ElseIf Index = 11 Then
                        lblWelcome = tFBCN & " " & sTabDesc & ":   " & NodeText
'''                        lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    ElseIf Index = 12 Then
                        lblWelcome = tFBCN & " " & sTabDesc & ":   " & NodeText
'''                        lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    ElseIf Index = 13 Then
                        lblWelcome = tFBCN & " Client Graphics:   " & NodeText
'''                        lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    ElseIf Index = 14 Then
                        lblWelcome = tFBCN & " Approval Graphics:   " & NodeText
'''                        picGfxApprove.Visible = True
'''                        lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    End If
                Case Else
                    lblWelcome = tFBCN & " " & sTabDesc & ":   " & NodeText
'''                    lblWelcome = tFBCN & " " & GfxType(iCurrGType) & ":   " & NodeText
                    
            End Select
            lblWelcome.Visible = True
            
            If bAddMode Then frmAssign.cmdAdd.FontBold = True
            
            If bPerm(59) And tBCC = 40579 Then
                lblKeyEdit.Enabled = False
            ElseIf bPerm(59) Then
                lblKeyEdit.Enabled = True
            End If
        End If
    Else
        rst.Close
        Set rst = Nothing
        MsgBox "Graphic not found", vbExclamation, "Sorry..."
        GoTo NoGraphic
    End If
    rst.Close
    Set rst = Nothing
    
    '///// NOW CHECK FOR REDLINE \\\\\
    lblGraphic.Visible = False
'    strSelect = "SELECT REDPATH, UPDUSER, UPDDTTM " & _
'                "FROM " & GFXRed & " " & _
'                "WHERE REF_ID = " & lRedID
    strSelect = "SELECT REDPATH, UPDUSER, UPDDTTM " & _
                "FROM " & GFXRed & " " & _
                "WHERE REF_ID = " & lRedID & " " & _
                "AND PAGE_ID = " & iPDFPage & " " & _
                "AND RED_STATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If bPerm(39) And Not rst.EOF Then
        lblGraphic.Caption = sTabDesc & ":  " & NodeParText ''''''''& "       [A Redline File Exists]"
        If iPDFPage > 0 Then
            ''CHECK FOR MULTIPLE PAGES''
            lblRedline.Caption = GetPDFRedCount(lRedID)
        Else
            lblRedline.Caption = "[A Redline File Exists]"
        End If
        mnuGRedLoad.Enabled = True
        RedFile = Trim(rst.Fields("REDPATH"))
        RedMess = "REDLINE STATUS:  Last Edited on " & Format(rst.Fields("UPDDTTM"), "mmm d, yyyy") & _
                    " by " & Trim(rst.Fields("UPDUSER")) & "."
    Else
        lblRedline.Caption = ""
        RedMess = ""
        On Error Resume Next
        Select Case Index
            Case 0
                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & cboSHYR(0).Text & " " & cboSHCD.Text
            Case 1
                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & cboSHYR(1).Text & " " & NodeParText
            Case 2
                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & NodeParText
            Case 3
                lblGraphic.Caption = GfxType(iCurrGType) & ":  " & tvwGraphics(3).Nodes(NodeParKey).Parent.Text
            Case 10
                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text & " " & NodeParText & ":  " & _
                            cboSHYR(0).Text & " " & cboSHCD.Text
            Case 11
                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text & " " & GfxType(iCurrGType) & ":  " & _
                            cboSHYR(1).Text '''& " " & tvwGraphics(1).Nodes(NodeParKey).Parent.Text
            Case 12
'''                    lblGraphic.Caption = sTabDesc & ":  " & tvwGraphics(2).Nodes(NodeParKey).Parent.Text & " " & _
'''                                NodeParText & "  [" & tvwGraphics(Index - 10).SelectedItem.Text & "]"
                lblGraphic.Caption = tvwGraphics(Index - 10).SelectedItem.Text '''& " " & NodeParText & ":  " & _
                            tvwGraphics(2).Nodes(NodeParKey).Text
            Case 13
                
'''                lblGraphic.Caption = Mid(tvwGraphics(3).SelectedItem.Parent.Text, 10) & " " & _
'''                            tvwGraphics(3).SelectedItem.Text
                    lblGraphic.Caption = tvwGraphics(3).SelectedItem.Text
            Case 14
                lblGraphic.Caption = NodeParText
                
            Case Else
                lblGraphic.Caption = sTabDesc & ":  " & NodeParText '''''& "  [" & nodetext & "]"
                
                
'                    lblGraphic.Caption = sTabDesc & ":  " & NodeParText & "  [" & tvwGraphics(Index - 10).SelectedItem.Text & "]"
        End Select
        If Err Then lblGraphic.Caption = ""
        
        mnuGRedLoad.Enabled = False
        RedFile = ""
    End If
    rst.Close
    Set rst = Nothing
    
    Call SetNavCnt(sst1.Tab)
    
    '///// CHECK FOR TEAM \\\\\
    If bPerm(26) Then
        bTeam = CheckForTeam(tBCC, tSHCD, frmGraphics)
        imgComm.Enabled = True
        If bTeam Then
             '///// CHECK FOR COMMENT \\\\\
             strSelect = "SELECT COMMID FROM " & ANOComment & " " & _
                        "WHERE REFID = " & lRefID & " " & _
                        "AND COMMSTATUS > 0"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                Set imgComm.Picture = imgMail(1).Picture
                imgComm.ToolTipText = "There are saved Comments! Click to access."
            Else
                Set imgComm.Picture = imgMail(0).Picture
                imgComm.ToolTipText = "There are no saved Comments."
            End If
            rst.Close
            Set rst = Nothing
        Else
            Set imgComm.Picture = imgMail(2).Picture
            imgComm.Enabled = False
        End If
        imgComm.Visible = True
    Else
        imgComm.Visible = False
    End If
    
'''    strSelect = "SELECT * " & _
'''                "FROM " & GFXMas & " " & _
'''                "WHERE GID = " & lGID
'''    Call GetGFXData(strSelect, "control")
    
NoGraphic:
    lblGraphic.Visible = True
    mnuGRedDelete.Enabled = False
    mnuGRedClear.Enabled = False
    mnuGRedSave.Enabled = False
    RedMode = False
    TextMode = False
    mnuGRedSketch.Checked = False
    mnuGRedText.Checked = False
End Sub

Public Sub ResetCounts(iCnt As Integer, CntIndex As Integer)
    Dim i As Integer, i1 As Integer, i2 As Integer, iInt As Integer
    For i = 0 To 9
        lblCount(i + (sst1.Tab * 10)).Visible = False
        lblCount(i + (sst1.Tab * 10)).Caption = (i * 20) + 1 & "-" & (i + 1) * 20
        lblCount(i + (sst1.Tab * 10)).ForeColor = vbBlack
        lblCount(i + (sst1.Tab * 10)).FontBold = False
    Next i
    If CntIndex = 99 Then
        lblViewAll(sst1.Tab).ForeColor = vbRed
        lblViewAll(sst1.Tab).Visible = True
    Else
        lblViewAll(sst1.Tab).ForeColor = vbBlack
    End If
    iInt = Int((iCnt - 1) / 20)
    If iInt > 0 Then
        If CntIndex <> 99 Then lblCount(CntIndex).ForeColor = vbRed
        If iInt >= 10 Then iInt = 9
        For i = 0 To iInt
            If i = iInt Then lblCount(i + (sst1.Tab * 10)).Caption = (i * 20) + 1 & "-" & iCnt
            lblCount(i + (sst1.Tab * 10)).Visible = True
        Next i
        lblViewAll(sst1.Tab).Left = lblCount(iInt + (sst1.Tab * 10)).Left + lblCount(iInt + (sst1.Tab * 10)).Width + 300
        lblViewAll(sst1.Tab).Top = lblCount(iInt + (sst1.Tab * 10)).Top
    End If
    If iInt > 1 Then lblViewAll(sst1.Tab).Visible = True Else lblViewAll(sst1.Tab).Visible = False
    
    '///// SET EDIT CONTROLS \\\\\'
'''    Select Case iCnt
'''        Case Is < 1
'''            cmdStatusEdit_View.Enabled = False
'''            cmdStatusEdit_All.Enabled = False
'''        Case Is <= 20
'''            cmdStatusEdit_View.Enabled = True
'''            cmdStatusEdit_All.Enabled = False
'''        Case Is > 20
'''            cmdStatusEdit_View.Enabled = True
'''            cmdStatusEdit_All.Enabled = True
'''
'''    End Select
        
End Sub

Public Sub GetGraphics(Index As Integer, CntIndex As Integer, strSelect As String, iStart As Integer, sNode As String)
    Dim rst As ADODB.Recordset
    Dim i1 As Integer, i2 As Integer, iDash As Integer, iCnt As Integer
    Dim i As Integer, iLock As Integer, iCol As Integer, iRow As Integer
    Dim imxCon As ImagXpress
    Dim shpCon As Shape
    Dim lblCon As Label
    Dim chkCon As CheckBox
    Dim sGStatus(0 To 30) As String
    Dim sFile As String
    
    If Not bResizing Then
        ''CLEAR MEMORY''
        Select Case Index
            Case 0
                For i = 1 To imx0.Count - 1
                    Unload imx0(i): Unload lbl0(i): Unload chk0(i)
                Next i
            Case 1
                For i = 1 To imx1.Count - 1
                    Unload imx1(i): Unload lbl1(i): Unload shp1(i): Unload chk1(i)
                Next i
            Case 2
                For i = 1 To imx2.Count - 1
                    Unload imx2(i): Unload lbl2(i): Unload shp2(i): Unload chk2(i)
                Next i
            Case 3
                For i = 1 To imx3.Count - 1
                    Unload imx3(i): Unload lbl3(i): Unload shp3(i): Unload chk3(i)
                Next i
        End Select
    End If
    
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
'''    iDash = InStr(1, sSpan, "-")
'''    i1 = CInt(Left(sSpan, iDash - 1)) - 1
'''    i2 = CInt(Mid(sSpan, iDash + 1)) - 1
    i1 = iStart - 1 '''CInt(Left(sSpan, iDash - 1)) - 1
    i2 = i1 + (iCols * iRows) - 1 '' 19 ''CInt(Mid(sSpan, iDash + 1)) - 1
    
    Set rst = Conn.Execute(strSelect)
    i = 0: iCnt = 0
'''''        picInner(Index).Visible = False
    lstFiles(Index).Clear
    
    Do While Not rst.EOF
        Do While iCnt < i1
            lstFiles(Index).AddItem Trim(rst.Fields("GDESC"))
            Select Case Index
                Case 0, 1, 10, 11
                    lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("SHOW_ID")
                Case 2, 12
                    lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("ES_ID")
                Case 3, 4, 13, 14
                    lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("GID")
            End Select
            rst.MoveNext
            iCnt = iCnt + 1
        Loop
        
        lstFiles(Index).AddItem Trim(rst.Fields("GDESC"))
        Select Case Index
            Case 0, 1, 10, 11
                If UCase(Left(tvwGraphics(Index).SelectedItem.Key, 1)) = "E" Then
                    lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("ES_ID") '' rst.Fields("SHOW_ID")
                Else
                    lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("SHOW_ID")
                End If
            Case 2, 12
                lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("ES_ID")
            Case 3, 4, 13, 14
                lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("GID")
        End Select
        
        iCol = Int(i / iRows): iRow = i Mod iRows
        Select Case Index
            Case 0
                If i >= imx0.Count Then Load imx0(i)
                Set imxCon = imx0(i)
            Case 1
                If i >= imx1.Count Then Load imx1(i)
                Set imxCon = imx1(i)
            Case 2
                If i >= imx2.Count Then Load imx2(i)
                Set imxCon = imx2(i)
            Case 3
                If i >= imx3.Count Then Load imx3(i)
                Set imxCon = imx3(i)
        End Select
        With imxCon
            .Left = 360 + (iCol * spaceX)
            .Top = 120 + (iRow * spaceY)
            .Update = False
            .CompressInMemory = CMEM_ALWAYS
            
'''            .PICThumbnail = THUMB_64 '' THUMB_None
            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
                ''LOOK FOR BMP HERE''
                sFile = sGPath & "pdf_" & rst.Fields("GID") & ".bmp"
                If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.BMP''
                    ''PDF.BMP NOT FOUND''
                    sFile = sGPath & "pdf_" & rst.Fields("GID") & ".jpg"
                    If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.JPG''
                        ''NO THUMBNAIL AT ALL''
                        .PICThumbnail = THUMB_64
                        .FileName = sGPath & "pdf.bmp"
                    Else
                        ''DISPLAY PDF.JPG''
                        Select Case FileLen(sFile)
                            Case Is < 10000: .PICThumbnail = THUMB_None
                            Case Is < 25000: .PICThumbnail = THUMB_4
                            Case Is < 50000: .PICThumbnail = THUMB_16
                            Case Else: .PICThumbnail = THUMB_64
                        End Select
                        .FileName = sFile
                    End If
                Else
                    ''DISPLAY PDF.BMP''
                    Select Case FileLen(sFile)
                        Case Is < 10000: .PICThumbnail = THUMB_None
                        Case Is < 25000: .PICThumbnail = THUMB_4
                        Case Is < 50000: .PICThumbnail = THUMB_16
                        Case Else: .PICThumbnail = THUMB_64
                    End Select
                    .FileName = sFile
                End If
            Else
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
                
'''                Select Case FileLen(Trim(rst.Fields("GPATH")))
'''                    Case Is < 10000: .PICThumbnail = THUMB_None
'''                    Case Is < 25000: .PICThumbnail = THUMB_4
'''                    Case Is < 50000: .PICThumbnail = THUMB_16
'''                    Case Else: .PICThumbnail = THUMB_64
'''                End Select
'''                .FileName = Trim(rst.Fields("GPATH"))
            End If
'''            If bApprover Then
'''                .ToolTipText = "Right-Click to Reset the Status of this Graphic File"
'''            Else
'''                .ToolTipText = Trim(rst.Fields("GDESC"))
                .ToolTipText = sGStatus(rst.Fields("GSTATUS")) & _
                        " - " & Trim(rst.Fields("GDESC")) & "  [ID " & rst.Fields("GID") & "]"
'''            End If
            .Update = True
'''            .Buttonize 1, 1, 50
            .Visible = True
            .Refresh
            Select Case Index
                Case 0
                    If UCase(Mid(sNode, 2, 1)) = "S" Then
                        .Tag = "S" & CStr(rst.Fields("SHOW_ID"))
                    ElseIf UCase(Mid(sNode, 2, 1)) = "E" Then
                        .Tag = "E" & CStr(rst.Fields("ES_ID"))
                    End If
                Case 1
                    If UCase(Mid(tvwGraphics(1).Nodes(sNode).Key, 2, 1)) = "S" Then
                        .Tag = CStr(rst.Fields("SHOW_ID"))
'''                    ElseIf UCase(Left(tvwGraphics(1).Nodes(sNode).key, 1)) = "E" Then
                    ElseIf UCase(Left(tvwGraphics(1).Nodes(sNode).Key, 1)) = "E" Then
                        .Tag = CStr(rst.Fields("ES_ID"))
                    ElseIf UCase(Left(tvwGraphics(1).Nodes(sNode).Key, 1)) = "R" Then
                        .Tag = CStr(rst.Fields("SHOW_ID"))
                    ElseIf UCase(Left(tvwGraphics(1).Nodes(sNode).Key, 1)) = "F" Then
                        .Tag = CStr(rst.Fields("SHOW_ID"))
                    End If
                Case 2: .Tag = CStr(rst.Fields("ES_ID"))
                Case 3: .Tag = CStr(rst.Fields("GID"))
            End Select
        End With
        
        Select Case Index
            Case 0
                If i >= lbl0.Count Then Load lbl0(i)
                Set lblCon = lbl0(i)
            Case 1
                If i >= lbl1.Count Then Load lbl1(i)
                Set lblCon = lbl1(i)
            Case 2
                If i >= lbl2.Count Then Load lbl2(i)
                Set lblCon = lbl2(i)
            Case 3
                If i >= lbl3.Count Then Load lbl3(i)
                Set lblCon = lbl3(i)
        End Select
        With lblCon
                If Len(Trim(rst.Fields("GDESC"))) > 30 Then
                    .Caption = Left(Trim(rst.Fields("GDESC")), 30) & "..."
                Else
                    .Caption = Trim(rst.Fields("GDESC"))
                End If
            .Left = imxCon.Left + (imxCon.Width / 2) - (lblCon.Width / 2)
'''            .Left = 120 + (iCol * spaceX) + ((imageX - .Width) / 2)
            .Top = imxCon.Top + imxCon.Height + 60 ''120 + imxCon.Height + (iRow * spaceY) + 60
            .ToolTipText = sGStatus(rst.Fields("GSTATUS")) & _
                        " - " & Trim(rst.Fields("GDESC"))
            lblCon.Tag = CStr(rst.Fields("GID"))
'            .BackColor = lBColor(rst.Fields("GSTATUS"))
'''            If rst.Fields("GSTATUS") > 0 Then
'''                .BackColor = vbWindowBackground
'''            Else
'''                .BackColor = vbRed
'''                lblInactive(Index).Visible = True
'''            End If
            .Visible = True
        End With
        
        Select Case Index
            Case 1
                If i >= shp1.Count Then Load shp1(i)
                Set shpCon = shp1(i)
            Case 2
                If i >= shp2.Count Then Load shp2(i)
                Set shpCon = shp2(i)
            Case 3
                If i >= shp3.Count Then Load shp3(i)
                Set shpCon = shp3(i)
        End Select
        shpCon.BackColor = lBColor(rst.Fields("GSTATUS"))
        shpCon.Width = lblCon.Width
        shpCon.Height = lblCon.Height
        shpCon.Left = lblCon.Left + 60
        shpCon.Top = lblCon.Top + 45
        shpCon.Visible = True
        
        Select Case Index
            Case 0
                If i >= chk0.Count Then Load chk0(i)
                Set chkCon = chk0(i)
            Case 1
                If i >= chk1.Count Then Load chk1(i)
                Set chkCon = chk1(i)
            Case 2
                If i >= chk2.Count Then Load chk2(i)
                Set chkCon = chk2(i)
            Case 3
                If i >= chk3.Count Then Load chk3(i)
                Set chkCon = chk3(i)
        End Select
        chkCon.Left = imxCon.Left
        chkCon.Top = imxCon.Top
        chkCon.ZOrder
        
        i = i + 1
        iCnt = iCnt + 1
        If iCnt > i2 Then
            Do While Not rst.EOF
                rst.MoveNext
                If Not rst.EOF Then
                    iCnt = iCnt + 1
                    lstFiles(Index).AddItem Trim(rst.Fields("GDESC"))
                    Select Case Index
                        Case 0, 1, 10, 11
                            lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("SHOW_ID")
                        Case 2, 12
                            lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("ES_ID")
                        Case 3, 4, 13, 14
                            lstFiles(Index).ItemData(lstFiles(Index).NewIndex) = rst.Fields("GID")
                    End Select
                End If
            Loop
            iCnt = iCnt '' - 1
            GoTo CountDone
        End If
        rst.MoveNext
    Loop
CountDone:
    rst.Close: Set rst = Nothing
    
    iGFXCount(Index) = iCnt
    If iGFXCount(Index) > 0 Then
        fraMulti.Visible = True
        iModeTab = Index
    Else
        fraMulti.Visible = False
        iModeTab = 0
    End If
'''    Call ResetCounts(iCnt, CntIndex)
    Call ResetBatch(iListStart(Index), iGFXCount(Index), Index)
    
    Select Case Index
        Case 0
            Call ClearThumbnails0(i)
'            picType.Visible = imx0(0).Visible
'''            For i = i To imx0.Count - 1
'''                imx0(i).Visible = False
'''                imx0(i).FileName = ""
'''                lbl0(i).Visible = False
'''            Next i
        Case 1
            Call ClearThumbnails1(i)
'            picType.Visible = imx1(0).Visible
'''            For i = i To imx1.Count - 1
'''                imx1(i).Visible = False
'''                imx1(i).FileName = ""
'''                lbl1(i).Visible = False
'''            Next i
        Case 2
            Call ClearThumbnails2(i)
'            picType.Visible = imx2(0).Visible
        Case 3
            Call ClearThumbnails3(i)
'            picType.Visible = imx3(0).Visible
            
    End Select
'    picInner(Index).Width = ((iCol + 1) * spaceX) + 240
'    If picInner(Index).Width < picOuter(Index).ScaleWidth Then
'        hsc1(Index).Max = picInner(Index).Width / 100
'        hsc1(Index).Visible = False
'    Else
'        hsc1(Index).Max = (picInner(Index).Width / 100) - (picOuter(Index).ScaleWidth / 100)
'        hsc1(Index).Visible = True
'    End If
'    hsc1(Index).value = 0 '''picOuter(1).ScaleWidth
'    hsc1(Index).LargeChange = picOuter(Index).ScaleWidth / 100

    picInner(Index).Visible = True
    bPopped(Index) = True
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
'''    ElseIf picJPG.Width > maxX Or picJPG.Height > maxY Then
'''        ''FORM HAS BEEN RESIZED SMALLER THAN GRAPHIC''
'''        Select Case iImageState
'''            Case 0
'''                Select Case rAsp
'''                    Case Is = rFAsp
        
                
    End If
    
End Sub

Public Sub SetGFXCUNOArray(lID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iAdd As Integer, i  As Integer
    
    iAdd = 1
'''    strSelect = "SELECT T.AN8_CUNO AS CUNO " & _
'''                "FROM " & ANOETeamUR & " R, " & ANOETeam & " T " & _
'''                "WHERE R.USER_SEQ_ID = " & lID & " " & _
'''                "AND R.TEAM_ID = T.TEAM_ID " & _
'''                "AND R.RECIPIENT_FLAG1 = 1"
    If bPerm(71) Then
        strSelect = "SELECT DISTINCT T.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
                    "FROM ANNOTATOR.ANO_EMAIL_TEAM_USER_R R, ANNOTATOR.ANO_EMAIL_TEAM T, " & F0101 & " AB, ANNOTATOR.GFX_MASTER GM " & _
                    "Where R.USER_SEQ_ID = " & lID & " " & _
                    "AND R.TEAM_ID = T.TEAM_ID " & _
                    "AND T.AN8_CUNO = AB.ABAN8 " & _
                    "AND T.AN8_CUNO = GM.AN8_CUNO " & _
                    "AND GM.GSTATUS IN (10, 20, 27) " & _
                    "ORDER BY CLIENT"
                    
    ElseIf bGPJ Then
        strSelect = "SELECT DISTINCT T.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
                    "FROM " & ANOETeamUR & " R, " & ANOETeam & " T, " & _
                    F0101 & " AB, " & GFXMas & " GM " & _
                    "WHERE R.USER_SEQ_ID = " & lID & " " & _
                    "AND R.TEAM_ID = T.TEAM_ID " & _
                    "AND R.RECIPIENT_FLAG1 = 1 " & _
                    "AND T.AN8_CUNO = AB.ABAN8 " & _
                    "AND T.AN8_CUNO = GM.AN8_CUNO " & _
                    "AND GM.GSTATUS IN (10, 20, 27) " & _
                    "Union " & _
                    "SELECT DISTINCT GM.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
                    "FROM " & F0101 & " AB, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GM.GID > 0 " & _
                    "AND GM.GAPPROVER_ID = " & lID & " " & _
                    "AND GM.GSTATUS IN (10, 20, 27) " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "ORDER BY CLIENT"
    Else
'''        strSelect = "SELECT DISTINCT T.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
'''                    "FROM " & ANOETeamUR & " R, " & ANOETeam & " T, " & _
'''                    F0101 & " AB, " & GFXMas & " GM " & _
'''                    "WHERE R.USER_SEQ_ID = " & lID & " " & _
'''                    "AND R.TEAM_ID = T.TEAM_ID " & _
'''                    "AND R.RECIPIENT_FLAG1 = 1 " & _
'''                    "AND T.AN8_CUNO = AB.ABAN8 " & _
'''                    "AND T.AN8_CUNO = GM.AN8_CUNO " & _
'''                    "AND GM.GSTATUS IN (20, 27) " & _
'''                    "ORDER BY CLIENT"
        
'        strSelect = "SELECT DISTINCT T.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
'                    "FROM ANO_EMAIL_TEAM_USER_R R, ANO_EMAIL_TEAM T, " & _
'                    "" & F0101 & " AB, GFX_MASTER GM " & _
'                    "Where R.USER_SEQ_ID = " & lID & " " & _
'                    "AND R.TEAM_ID = T.TEAM_ID " & _
'                    "AND R.RECIPIENT_FLAG1 = 1 " & _
'                    "AND T.AN8_CUNO = AB.ABAN8 " & _
'                    "AND T.AN8_CUNO = GM.AN8_CUNO " & _
'                    "AND GM.GSTATUS IN (20, 27) " & _
'                    "Union " & _
'                    "SELECT DISTINCT GM.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
'                    "FROM " & F0101 & " AB, GFX_MASTER GM " & _
'                    "Where GM.GID > 0 " & _
'                    "AND GM.GAPPROVER_ID = " & lID & " " & _
'                    "AND GM.GSTATUS IN (20, 27) " & _
'                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
'                    "ORDER BY CLIENT"
        
        strSelect = "SELECT DISTINCT T.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
                    "FROM ANNOTATOR.ANO_EMAIL_TEAM_USER_R R, ANNOTATOR.ANO_EMAIL_TEAM T, " & _
                    "" & F0101 & " AB, ANNOTATOR.GFX_MASTER GM " & _
                    "Where R.USER_SEQ_ID = " & lID & " " & _
                    "AND R.TEAM_ID = T.TEAM_ID " & _
                    "AND (R.RECIPIENT_FLAG1 = 1 OR R.EXTCLIENTAPPROVER_FLAG = 1) " & _
                    "AND T.AN8_CUNO = AB.ABAN8 " & _
                    "AND T.AN8_CUNO = GM.AN8_CUNO " & _
                    "AND GM.GSTATUS IN (20, 27) " & _
                    "Union " & _
                    "SELECT DISTINCT GM.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT " & _
                    "FROM " & F0101 & " AB, ANNOTATOR.GFX_MASTER GM " & _
                    "Where GM.GID > 0 " & _
                    "AND GM.GAPPROVER_ID = " & lID & " " & _
                    "AND GM.GSTATUS IN (20, 27) " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "ORDER BY CLIENT"
    End If
                
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        ReDim Preserve GFXCUNO(iAdd)
        GFXCUNO(iAdd - 1) = rst.Fields("CUNO")
        iAdd = iAdd + 1
        
        cboCUNO(4).AddItem Trim(rst.Fields("CLIENT"))
        cboCUNO(4).ItemData(cboCUNO(4).NewIndex) = rst.Fields("CUNO")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    ReDim Preserve GFXCUNO(iAdd - 2)
    
    
    
    If Not bFromFloorplan Then
'''        For i = LBound(GFXCUNO) To UBound(GFXCUNO)
            
        If fBCC(4) <> "" And fFBCN(4) <> "" Then
            On Error Resume Next
            cboCUNO(4).Text = fFBCN(4)

            If Err = 0 Then sst1.Tab = 4 Else sst1.Tab = 3
        Else
            sst1.Tab = 3
        End If
    Else
        sst1.Tab = 1
    End If
End Sub

Public Function CheckGFXCUNO(lID As Long) As Boolean
    Dim i As Integer
    For i = LBound(GFXCUNO) To UBound(GFXCUNO)
        If GFXCUNO(i) = lID Then
            CheckGFXCUNO = True
            Exit Function
        End If
    Next i
    CheckGFXCUNO = False
End Function

Public Sub PopReviewClients(lID As Long)
    Dim strSelect As String, sClient As String
    Dim rst As ADODB.Recordset
    
    If cboCUNO(3).Text <> "" Then sClient = cboCUNO(3).Text Else sClient = ""
    
    cboCUNO(3).Clear: tvwGraphics(3).Nodes.Clear
    ''CLEAR IMAGES''
    strSelect = "SELECT T.AN8_CUNO AS CUNO, C.ABALPH AS CLIENT " & _
                "FROM " & ANOETeamUR & " R, " & ANOETeam & " T, " & F0101 & " C " & _
                "WHERE R.USER_SEQ_ID = " & lID & " " & _
                "AND R.RECIPIENT_FLAG1 = 1 " & _
                "AND R.TEAM_ID = T.TEAM_ID " & _
                "AND T.AN8_CUNO = C.ABAN8 " & _
                "ORDER BY UPPER(C.ABALPH)"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboCUNO(3).AddItem UCase(Trim(rst.Fields("CLIENT")))
        cboCUNO(3).ItemData(cboCUNO(3).NewIndex) = rst.Fields("CUNO")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    On Error Resume Next
    If sClient <> "" Then cboCUNO(3).Text = sClient
    If Err Then tvwGraphics(3).Nodes.Clear
    
End Sub

Public Sub ClearThumbnails0(iStart As Integer)
    Dim i As Integer
    For i = iStart To imx0.Count - 1
        imx0(i).Visible = False
        imx0(i).FileName = ""
        lbl0(i).Visible = False
        chk0(i).Visible = False
    Next i
End Sub

Public Sub ClearThumbnails1(iStart As Integer)
    Dim i As Integer
    For i = iStart To imx1.Count - 1
        imx1(i).Visible = False
        imx1(i).FileName = ""
        lbl1(i).Visible = False
        shp1(i).Visible = False
        chk1(i).Visible = False
    Next i
End Sub

Public Sub ClearThumbnails2(iStart As Integer)
    Dim i As Integer
    For i = iStart To imx2.Count - 1
        imx2(i).Visible = False
        imx2(i).FileName = ""
        lbl2(i).Visible = False
        shp2(i).Visible = False
        chk2(i).Visible = False
    Next i
End Sub

Public Sub ClearThumbnails3(iStart As Integer)
    Dim i As Integer
    For i = iStart To imx3.Count - 1
        imx3(i).Visible = False
        imx3(i).FileName = ""
        lbl3(i).Visible = False
        shp3(i).Visible = False
        chk3(i).Visible = False
    Next i
End Sub

'''Public Function CheckForNotify() As Boolean
'''    Dim strSelect As String
'''    Dim rst As ADODB.Recordset
'''
'''    strSelect = "SELECT GID FROM " & GFXMas & " " & _
'''                "WHERE GSTATUS IN (5, 15, 25) " & _
'''                "AND UPDUSER = '" & LogName & "'"
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        rst.Close: Set rst = Nothing
'''        CheckForNotify = True
'''    Else
'''        rst.Close: Set rst = Nothing
'''        CheckForNotify = False
'''    End If
'''
'''End Function


'''Public Sub GetGFXData(strSelect As String, sDisplay As String)
'''    Dim sMess As String, sSize As String
'''    Dim rst As ADODB.Recordset
'''    Dim sGStatus(0 To 30) As String
'''    Dim lSize As Long
'''
'''    '///// FILE STATUS VARIABLES \\\\\
'''    sGStatus(0) = "DE-ACTIVED"
'''    sGStatus(10) = "INTERNAL"
'''    sGStatus(20) = "CLIENT DRAFT"
'''    sGStatus(27) = "RETURNED FOR CHANGES"
'''    sGStatus(30) = "APPROVED"
'''
'''
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        Select Case FileLen(Trim(rst.Fields("GPATH")))
'''            Case Is < 1000: sSize = format(FileLen(Trim(rst.Fields("GPATH"))), "#,##0") & " bytes"
'''            Case Is < 2000000: sSize = format(FileLen(Trim(rst.Fields("GPATH"))) / 1000, "#,##0") & " k"
'''            Case Else: sSize = format(FileLen(Trim(rst.Fields("GPATH"))) / 1000000, "#,##0.00") & " mb"
'''        End Select
'''
'''        sMess = "Graphic Description:" & vbTab & Trim(rst.Fields("GDESC")) & vbNewLine & _
'''                    "Database I.D.:          " & vbTab & rst.Fields("GID") & vbNewLine & _
'''                    "File Format:              " & vbTab & Trim(rst.Fields("GFORMAT")) & vbNewLine & _
'''                    "Graphic Type:           " & vbTab & GfxType(rst.Fields("GTYPE")) & vbNewLine & _
'''                    "Graphic Status:         " & vbTab & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
'''                    "File Size:                   " & vbTab & sSize & vbNewLine & vbNewLine
'''        sMess = sMess & "File Added by " & Trim(rst.Fields("ADDUSER")) & " on " & _
'''                    format(rst.Fields("ADDDTTM"), "mmmm d, yyyy") & "." & vbNewLine
'''        sMess = sMess & "File Last Edited by " & Trim(rst.Fields("UPDUSER")) & " on " & _
'''                    format(rst.Fields("UPDDTTM"), "mmmm d, yyyy") & "."
'''        rst.Close
'''        Select Case sDisplay
'''            Case "msgbox"
'''                MsgBox sMess, vbInformation, "Graphic Data..."
'''            Case "control"
'''                txtXData(1).Text = sMess
'''        End Select
''''        picXData.Visible = True
'''    Else
'''        rst.Close
''''''        picXData.Visible = False
'''        MsgBox "No Data Available.", vbInformation, "Graphic Data..."
'''    End If
'''    Set rst = Nothing
'''
'''End Sub



Public Function CheckIfFutureShow(pSHYR As Integer, pSHCD As Long) As Boolean
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bCheck As Boolean
    
    strSelect = "SELECT SHY56SHCD " & _
                "FROM " & F5601 & " " & _
                "WHERE SHY56SHYR = " & pSHYR & " " & _
                "AND SHY56SHCD = " & pSHCD & " " & _
                "AND SHY56ENDDT > " & IGLToJDEDate(Now)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then bCheck = True Else bCheck = False
    rst.Close: Set rst = Nothing
    
    CheckIfFutureShow = bCheck
    
End Function

Public Sub GetApprovalGraphics(lCUNO As Long, sOrder As String, iSHYR As Integer, lSHCD As Long, lFID As Long)
    Dim strSelect As String, sFile As String
    Dim rst As ADODB.Recordset, rstC As ADODB.Recordset
    Dim iRow As Integer, iCol As Integer
    Dim i As Long
    Dim sGStatus(10 To 30) As String
    Dim GfxType(1 To 4) As String
    Dim sPDFImage As String
    
    ''CLEAR CONTROLS''
    If imx4.Count > 1 Then
        On Error Resume Next
        For i = imx4.Count - 1 To 1 Step -1
            Unload imx4(i)
            Unload shp4(i)
            Unload imgV(i)
            Unload imgStat(i)
            Unload lblStat(i)
            Unload chk4(i)
        Next i
        On Error GoTo 0
    End If
    lstFiles(4).Clear
    Err.Clear
    
    sPDFImage = sGPath & "Acrobatid.bmp"
    
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED"
    sGStatus(30) = "APPROVED"
    
    '///// REDEFINE GRAPHIC TYPES, TEMPORARILY \\\\\
    GfxType(1) = "Digital Photo"
    GfxType(2) = "Graphic"
    GfxType(3) = "Layout"
    GfxType(4) = "Presentation"
    
    iRow = 0
    flxApprove.Rows = 1
    Select Case iApproverView
        Case 0 ''MY FILES ONLY''
            If lSHCD = 0 And lFID = 0 Then ''ALL CLIENT FILES''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND GM.GAPPROVER_ID = " & UserID & " " & _
                            "AND GM.AN8_CUNO = " & lCUNO & " " & _
                            "AND GM.GAPPROVER_ID = " & UserID & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            ElseIf lFID = 0 Then ''CLIENT - SHOW''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GS.SHYR = " & iSHYR & " " & _
                            "AND GS.AN8_SHCD = " & lSHCD & " " & _
                            "AND GS.AN8_CUNO = " & lCUNO & " " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GID > 0 " & _
                            "AND GM.GAPPROVER_ID = " & UserID & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            ElseIf lSHCD = 0 Then ''CLIENT - FOLDER''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND GM.AN8_CUNO = " & lCUNO & " " & _
                            "AND GM.GAPPROVER_ID = " & UserID & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.FLR_ID = " & lFID & " " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            Else  ''CLIENT - SHOW - FOLDER''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GS.SHYR = " & iSHYR & " " & _
                            "AND GS.AN8_SHCD = " & lSHCD & " " & _
                            "AND GS.AN8_CUNO = " & lCUNO & " " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GID > 0 " & _
                            "AND GM.GAPPROVER_ID = " & UserID & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.FLR_ID = " & lFID & " " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            End If
        Case 1 ''ALL FILES''
            If lSHCD = 0 And lFID = 0 Then ''ALL CLIENT FILES''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND GM.AN8_CUNO = " & lCUNO & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            ElseIf lFID = 0 Then ''CLIENT - SHOW''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GS.SHYR = " & iSHYR & " " & _
                            "AND GS.AN8_SHCD = " & lSHCD & " " & _
                            "AND GS.AN8_CUNO = " & lCUNO & " " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GID > 0 " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            ElseIf lSHCD = 0 Then ''CLIENT - FOLDER''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND GM.AN8_CUNO = " & lCUNO & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.FLR_ID = " & lFID & " " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            Else  ''CLIENT - SHOW - FOLDER''
                strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                            "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM " & GFXShow & " GS, " & GFXMas & " GM, " & IGLUser & " U " & _
                            "WHERE GS.SHYR = " & iSHYR & " " & _
                            "AND GS.AN8_SHCD = " & lSHCD & " " & _
                            "AND GS.AN8_CUNO = " & lCUNO & " " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GID > 0 " & _
                            "AND GM.FLR_ID = " & lFID & " " & _
                            "AND GM.GSTATUS IN (" & sIN & ") " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) "
                            
            End If
        Case 2 ''SEARCH RESULT''
            strSelect = "SELECT GM.GPATH, GM.GID, GM.GDESC, GM.GFORMAT, " & _
                        "GM.GTYPE, GM.GSTATUS, GM.ADDUSER, GM.ADDDTTM, GM.VERSION_ID, " & _
                        "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                        "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                        "WHERE GM.GID > 0 " & _
                        "AND GM.GID IN (" & sSearchList & ") " & _
                        "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) " & _
                        "AND GM.GSTATUS IN (" & sIN & ") "

    End Select
    strSelect = strSelect & sOrder
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        imx4(0).Visible = False
        shp4(0).Visible = False
        imgV(0).Visible = False
        imgStat(0).Visible = False
        lblStat(0).Visible = False
        chk4(0).Visible = False
    End If
    Do While Not rst.EOF
        lstFiles(4).AddItem Trim(rst.Fields("GDESC"))
        lstFiles(4).ItemData(lstFiles(4).NewIndex) = rst.Fields("GID")
        
        iRow = iRow + 1
        flxApprove.Rows = iRow + 1
        flxApprove.RowHeight(iRow) = 1080
        flxApprove.TextMatrix(iRow, 0) = rst.Fields("GID")
        flxApprove.TextMatrix(iRow, 3) = Trim(rst.Fields("GDESC"))
        flxApprove.TextMatrix(iRow, 5) = GfxType(rst.Fields("GTYPE"))
        
        If Trim(rst.Fields("APPROVER")) <> "" Then
            flxApprove.TextMatrix(iRow, 6) = StrConv(Trim(rst.Fields("APPROVER")), vbProperCase)
        Else
            flxApprove.Row = iRow: flxApprove.Col = 6
            Set flxApprove.CellPicture = imgApprovers.Picture
            flxApprove.CellPictureAlignment = 4
        End If
        
        flxApprove.TextMatrix(iRow, 7) = Format(rst.Fields("ADDDTTM"), "ddd, mmm d, yyyy (hh:nn ampm)")
        flxApprove.TextMatrix(iRow, 8) = StrConv(Trim(rst.Fields("ADDUSER")), vbProperCase)
        
        i = iRow - 1
        If iRow > imx4.Count Then
'            i = iRow - 1
            Load shp4(i)
            shp4(i).Top = shp4(i).Top + (i * 1080)
            If i Mod 2 = 1 Then
                shp4(i).BackColor = vb3DLight
                shp4(i).BorderColor = vb3DLight
            End If
            shp4(i).Visible = True
            
            Load imx4(i)
            imx4(i).Top = imx4(i).Top + (i * 1080) '''1500)
            If (i) Mod 2 = 1 Then imx4(i).BackColor = vb3DLight
            imx4(i).Visible = True
            imx4(i).CompressInMemory = CMEM_ALWAYS
            imx4(i).PICThumbnail = THUMB_64
            imx4(i).ZOrder
            
            Load imgV(i)
            imgV(i).Container = picInner(4)
            imgV(i).Top = imgV(i).Top + (i * 1080)
            
'''            imgV(i).ZOrder
'''            imgV(i).Tag = CStr(rst.Fields("GID"))
'''            If rst.Fields("VERSION_ID") > 0 Then imgV(i).Visible = True _
'''                        Else imgV(i).Visible = False
            
            Load imgStat(i)
            imgStat(i).Top = imgStat(i).Top + (i * 1080) '''1500)
            imgStat(i).Visible = True
            imgStat(i).ZOrder
            
            Load lblStat(i)
            lblStat(i).Top = lblStat(i).Top + (i * 1080) '''1500)
            lblStat(i).Visible = True
            lblStat(i).ZOrder
            
            Load chk4(i)
            chk4(i).Top = imx4(i).Top
            chk4(i).Left = imx4(i).Left
            chk4(i).Visible = False
        End If
        
        imx4(i).Update = False
        imx4(i).Tag = CStr(rst.Fields("VERSION_ID"))
        
''''        If rst.Fields("VERSION_ID") = 0 Then
'''            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
'''                imx4(i).FileName = sPDFImage
'''            Else
'''                imx4(i).FileName = Trim(rst.Fields("GPATH"))
'''            End If
''''        Else
''''            If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
''''                imx4(i).FileName = sPDFImage
''''            Else
''''                imx4(i).FileName = sVPath & rst.Fields("VERSION_ID") & _
''''                            "." & Trim(rst.Fields("GFORMAT"))
''''            End If
''''        End If
        
        If UCase(Trim(rst.Fields("GFORMAT"))) = "PDF" Then
            ''LOOK FOR THUMB HERE''
            sFile = sGPath & "pdf_" & rst.Fields("GID") & ".bmp"
            If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.BMP''
                sFile = sGPath & "PDF_" & rst.Fields("GID") & ".jpg"
                If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.JPG''
                    imx4(i).PICThumbnail = THUMB_64
                    imx4(i).FileName = sGPath & "pdf.bmp"
                Else
                    Select Case FileLen(sFile)
                        Case Is < 10000: imx4(i).PICThumbnail = THUMB_None
                        Case Is < 25000: imx4(i).PICThumbnail = THUMB_4
                        Case Is < 50000: imx4(i).PICThumbnail = THUMB_16
                        Case Else: imx4(i).PICThumbnail = THUMB_64
                    End Select
                    imx4(i).FileName = sFile
                End If
            Else
                Select Case FileLen(sFile)
                    Case Is < 10000: imx4(i).PICThumbnail = THUMB_None
                    Case Is < 25000: imx4(i).PICThumbnail = THUMB_4
                    Case Is < 50000: imx4(i).PICThumbnail = THUMB_16
                    Case Else: imx4(i).PICThumbnail = THUMB_64
                End Select
                imx4(i).FileName = sFile
            End If
        Else
            sFile = sGPath & "Thumbs\thb_" & rst.Fields("GID") & ".jpg"
            If Dir(sFile, vbNormal) = "" Then ''OPEN FULL FILE''
                Select Case FileLen(Trim(rst.Fields("GPATH")))
                    Case Is < 10000: imx4(i).PICThumbnail = THUMB_None
                    Case Is < 25000: imx4(i).PICThumbnail = THUMB_4
                    Case Is < 50000: imx4(i).PICThumbnail = THUMB_16
                    Case Else: imx4(i).PICThumbnail = THUMB_64
                End Select
                imx4(i).FileName = Trim(rst.Fields("GPATH"))
            Else
                Select Case FileLen(sFile)
                    Case Is < 10000: imx4(i).PICThumbnail = THUMB_None
                    Case Is < 25000: imx4(i).PICThumbnail = THUMB_4
                    Case Is < 50000: imx4(i).PICThumbnail = THUMB_16
                    Case Else: imx4(i).PICThumbnail = THUMB_64
                End Select
                imx4(i).FileName = sFile
            End If
            
        End If
        
        
        
        
        imx4(i).Update = True
'''        imx4(i).Buttonize 1, 1, 50
        imx4(i).Visible = True
        imx4(i).CompressInMemory = CMEM_ALWAYS
        imx4(i).Refresh
        
        imgV(i).ZOrder
        imgV(i).Tag = CStr(rst.Fields("GID"))
        If rst.Fields("VERSION_ID") > 0 Then imgV(i).Visible = True _
                    Else imgV(i).Visible = False
        
        Set imgStat(i).Picture = imgStatus(rst.Fields("GSTATUS")).Picture
        imgStat(i).Visible = True
        lblStat(i).Caption = sGStatus(rst.Fields("GSTATUS"))
        lblStat(i).Visible = True
        
        
        If i Mod 2 = 1 Then
            With flxApprove
                .Row = i + 1
                For iCol = 1 To .Cols - 1
                    .Col = iCol: .CellBackColor = vb3DLight
                Next iCol
                
            End With
        End If
        
        ''CHECK FOR COMMENTS''
        flxApprove.Row = iRow: flxApprove.Col = 4
        strSelect = "SELECT COUNT(COMMID) AS CNT " & _
                    "FROM " & ANOComment & " " & _
                    "WHERE REFID = " & rst.Fields("GID")
        Set rstC = Conn.Execute(strSelect)
        If rstC.Fields("CNT") = 0 Then
            Set flxApprove.CellPicture = imgMail(0).Picture
        Else
            Set flxApprove.CellPicture = imgMail(1).Picture
        End If
        rstC.Close: Set rstC = Nothing
        flxApprove.CellPictureAlignment = 4

        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    picInner(4).Top = 0
    picInner(4).Height = CLng(iRow) * 1080
    
    lblFileCount.Caption = iRow & " Files"
    
    If iRow > 0 Then fraMulti.Visible = True Else fraMulti.Visible = False
    
'    If iRow = 0 Then
'        picReview.Enabled = False
'    Else
        picReview.Enabled = True
'    End If
        
End Sub

Public Sub PrepareToOpen(Index As Integer)
    Dim sStatus As String
    If bRedSaved = True And bTeam = True Then
        With frmRedAlert
            .PassGID = lGID
            .PassBCC = CLng(redBCC)
            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 1
            .Show 1
        End With
    End If
    bRedMode = False: shpRN.Visible = False: lblRN.Visible = False
    bRedSaved = False
    redBCC = "": redSHCD = 0
    
    iApprovalRow = Index + 1
    
    picJPG.Visible = False
'''    If bAcro Then pdfGraphic.Visible = False
    
'    If chkClose(4).value = 1 Then
        picTabs.Visible = False ''sst1.Visible = False
        bDirsOpen = False
        imgDirs.ToolTipText = "Click to Open File Index..."
        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'    End If
    '///// TIME TO LOAD THE GRAPHIC \\\\\
    If UCase(lblStat(Index).Caption) = "RETURNED" Then
        sStatus = "RETURNED FOR CHANGES"
    Else
        sStatus = lblStat(Index).Caption
    End If
        
'''    If imx4(Index).Tag = "" Then ''SINGLE VERSION''
        Call LoadGraphic(14, flxApprove.TextMatrix(Index + 1, 0), _
                    flxApprove.TextMatrix(Index + 1, 3), "", _
                    sStatus & " " & flxApprove.TextMatrix(Index + 1, 5))
'''    Else
'''        Call LoadGraphic(14, CLng(imx4(Index).Tag), _
'''                    flxApprove.TextMatrix(Index + 1, 3), "", _
'''                    sStatus & " " & flxApprove.TextMatrix(Index + 1, 5))
'''    End If
    
End Sub


'''Public Function InsertComment(lGID As Long, Index As Integer) As Integer
'''    Dim strInsert As String
'''    Dim rstL As ADODB.Recordset
'''    Dim lCOMMID As Long
'''    Dim sComm As String
'''
'''    On Error Resume Next
'''
'''    Select Case Index
'''        Case 0: sComm = "Graphic Status reset to 'INTERNAL DRAFT' by " & LogName & "."
'''        Case 1: sComm = "Graphic Status reset to 'CLIENT DRAFT' by " & LogName & "."
'''        Case 2: sComm = "Graphic 'APPROVED' by " & LogName & "."
'''        Case 3: sComm = "Graphic Cancelled by " & LogName & "."
'''    End Select
'''
'''    '///// GET NEW COMMID \\\\\
'''    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
'''    lCOMMID = rstL.Fields("nextval")
'''    rstL.Close: Set rstL = Nothing
'''
'''    strInsert = "INSERT INTO " & ANOComment & " " & _
'''            "(COMMID, REFID, REFSOURCE, ANO_COMMENT, " & _
'''            "COMMUSER, COMMDATE, COMMSTATUS, " & _
'''            "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
'''            "VALUES " & _
'''            "(" & lCOMMID & ", " & lGID & ", '" & sTable & "', '" & DeGlitch(sComm) & "', " & _
'''            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
'''            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
'''            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
'''    Conn.Execute (strInsert)
'''
'''    InsertComment = Err.Number
'''
'''End Function

Private Sub txtDim_KeyPress(Index As Integer, KeyAscii As Integer)
    If iUnit = 8 Then Call CheckNumeric(KeyAscii)
End Sub

Public Sub CheckNumeric(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = vbKeyBack And Not KeyAscii = 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

''''Private Sub txtDim_LostFocus(Index As Integer)
''''    Call DimConvert(Index, txtDim(Index).Text)
''''End Sub

''''Public Sub DimConvert(Index As Integer, sBase As String)
''''    Dim iSpc As Integer, iDiv As Integer, iFoot As Integer
''''    Dim rDim As Single, rFrac As Single
''''    Dim iStart As Integer
''''    Dim sBal As String
''''    Dim iDivisor As Integer
''''
''''
''''    rDim = 0
''''
''''    On Error GoTo ErrorTrap
''''
''''    ''FIRST CHECK FOR FOOT MARK''
''''    iFoot = InStr(1, sBase, "'")
''''    If iFoot <> 0 Then
''''        rDim = CSng(Left(sBase, iFoot - 1)) * 12
''''    End If
''''
''''    If iFoot = 0 Then iStart = 1 Else iStart = iFoot + 1
''''    If iStart > Len(Trim(sBase)) Then GoTo Done
''''
''''    If Mid(sBase, iStart, 1) = " " Or Mid(sBase, iStart, 1) = "-" Then
''''        iStart = iStart + 1
''''    End If
''''
''''    sBal = Trim(Mid(sBase, iStart))
''''
''''    If Right(sBal, 1) = """" Then
''''        sBal = Left(sBal, Len(sBal) - 1)
''''    End If
''''
''''    iSpc = InStr(1, sBal, " ")
''''    If iSpc > 0 Then
''''        rDim = rDim + CSng(Left(sBal, iSpc - 1))
''''        sBal = Mid(sBal, iSpc + 1)
''''    End If
''''
''''    iSpc = InStr(1, sBal, "-")
''''    If iSpc > 0 Then
''''        rDim = rDim + CSng(Left(sBal, iSpc - 1))
''''        sBal = Mid(sBal, iSpc + 1)
''''    End If
''''
''''    iDiv = InStr(1, sBal, "/")
''''    If iDiv > 0 Then
''''        Select Case CInt(Mid(sBal, iDiv + 1))
''''            Case 2, 4, 8, 16, 32, 64
''''            Case Else
''''                MsgBox "Only the following fraction denominators are recognized:" & _
''''                            vbNewLine & vbNewLine & vbTab & _
''''                            "x/2, x/4, x/8, x/16, x/32, x/64", vbExclamation, "Problem..."
''''                GoTo ErrorTrap
''''        End Select
''''        rFrac = CSng(Left(sBal, iDiv - 1)) / CSng(Mid(sBal, iDiv + 1))
''''        rDim = rDim + rFrac
''''    Else
''''        rDim = rDim + CSng(sBal)
''''    End If
''''
''''Done:
''''
''''    lblDim(Index).Caption = Format(rDim, "###0.0000")
''''Exit Sub
''''
''''ErrorTrap:
''''    lblDim(Index).Caption = ""
''''    MsgBox "Unable to resolve entry", vbExclamation, "Hey..."
''''    txtDim(Index).Text = ""
''''
''''End Sub



Public Sub LoadNodes(Index As Integer, NodeKey As String, NodeText As String, _
            NodeParKey As String, NodeParText As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sDesc As String, sENode As String, sTNode As String, sNodePref As String
    Dim iNode As Integer
    
    tvwGraphics(1).ImageList = ImageList1
    Select Case UCase(Mid(NodeKey, 2, 1))
        Case "S"
            strSelect = "SELECT " & _
                        "FROM "
                        
'''SELECT GM.GID, GM.GDESC, GM.GTYPE, GM.GSTATUS
'''FROM GFX_SHOW GS, GFX_MASTER GM
'''Where GS.SHYR = 2002
'''AND GS.AN8_CUNO = 1161
'''AND GS.AN8_SHCD = 12978
'''and GS.ELTID IS NULL
'''AND GS.GID = GM.GID
'''AND GM.GSTATUS IN (20, 30)
'''ORDER BY GTYPE, GDESC;
            

        Case "E", "R"
            sENode = "": sTNode = ""
            
            If UCase(Mid(NodeKey, 2, 1)) = "E" Then
                strSelect = "SELECT DISTINCT EU.ELTID, GM.GSTATUS, " & _
                            "(TRIM(K.KITFNAME) || '-' || TRIM(EU.ELTFNAME))STENCIL, " & _
                            "EU.ELTDESC , K.KITREF " & _
                            "FROM " & AQUAEltU & " EU, ANNOTATOR.GFX_ELEMENT GE, IGLPROD.IGL_KIT K, ANNOTATOR.GFX_MASTER GM " & _
                            "Where EU.AN8_SHCD = " & Mid(NodeKey, 3) & " " & _
                            "AND EU.AN8_CUNO = " & CLng(fBCC(sst1.Tab)) & " " & _
                            "AND EU.SHYR = " & fSHYR(sst1.Tab) & " " & _
                            "AND EU.SHSTATUS = 1 " & _
                            "AND EU.KITID = K.KITID " & _
                            "AND EU.ELTID = GE.ELTID " & _
                            "AND GE.GID = GM.GID " & _
                            "ORDER BY K.KITREF, EU.ELTDESC"
                sNodePref = "e"
            ElseIf UCase(Mid(NodeKey, 2, 1)) = "R" Then
                strSelect = "SELECT DISTINCT GS.ELTID, GM.GSTATUS, " & _
                            "(TRIM(K.KITFNAME) || '-' || TRIM(E.ELTFNAME))STENCIL, " & _
                            "E.ELTDESC , K.KITREF " & _
                            "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K " & _
                            "Where GS.SHYR = " & fSHYR(sst1.Tab) & " " & _
                            "AND GS.AN8_CUNO = " & CLng(fBCC(sst1.Tab)) & " " & _
                            "AND GS.AN8_SHCD = " & Mid(NodeKey, 3) & " " & _
                            "AND GS.ELTID IS NOT NULL " & _
                            "AND GS.GID = GM.GID " & _
                            "AND GM.GSTATUS IN (" & defSIN & ") " & _
                            "AND GS.ELTID = E.ELTID " & _
                            "AND E.KITID = K.KITID " & _
                            "ORDER BY K.KITREF, E.ELTDESC"
                sNodePref = "r"
            End If
'''            strSelect = "SELECT DISTINCT EU.ELTID, GM.GTYPE, GM.GSTATUS, " & _
'''                        "(TRIM(K.KITFNAME) || '-' || TRIM(EU.ELTFNAME))STENCIL, " & _
'''                        "EU.ELTDESC, K.KITREF " & _
'''                        "FROM IGL_ELEMENT_USE EU, GFX_ELEMENT GE, IGL_KIT K, GFX_MASTER GM " & _
'''                        "Where EU.AN8_SHCD = " & Mid(NodeKey, 3) & " " & _
'''                        "AND EU.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                        "AND EU.SHYR = " & tSHYR & " " & _
'''                        "AND EU.SHSTATUS = 1 " & _
'''                        "AND EU.KITID = K.KITID " & _
'''                        "AND EU.ELTID = GE.ELTID " & _
'''                        "AND GE.GID = GM.GID " & _
'''                        "ORDER BY K.KITREF, EU.ELTDESC, GM.GTYPE"
            
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                Do While Not rst.EOF
                    If rst.Fields("GSTATUS") = 20 Or rst.Fields("GSTATUS") = 30 Then
                        If sENode <> sNodePref & Mid(NodeKey, 3) & "-" & rst.Fields("ELTID") Then
                            sENode = sNodePref & Mid(NodeKey, 3) & "-" & rst.Fields("ELTID")
                            sDesc = Trim(rst.Fields("STENCIL")) & Space(2) & Trim(rst.Fields("ELTDESC"))
                            Set nodX = tvwGraphics(1).Nodes.Add(NodeParKey, tvwChild, sENode, sDesc, 12)
                            sTNode = ""
                        End If
'''                        If sTNode <> "t" & Mid(NodeKey, 3) & "-" & rst.Fields("ELTID") & "-" & rst.Fields("GTYPE") Then
'''                            sTNode = "t" & Mid(NodeKey, 3) & "-" & rst.Fields("ELTID") & "-" & rst.Fields("GTYPE")
'''                            sDesc = GfxType(rst.Fields("GTYPE"))
'''                            iNode = rst.Fields("GTYPE")
'''                            Set nodX = tvwGraphics(1).Nodes.Add(sENode, tvwChild, sTNode, sDesc, iNode)
'''                        End If
                    End If
                    rst.MoveNext
                Loop
                tvwGraphics(1).Nodes.Remove (NodeKey)
            End If
            rst.Close: Set rst = Nothing
            
    End Select
    
End Sub

Public Sub SetType(Index As Integer)
    picType.Visible = False
    Select Case Index
        Case 0
            shpType.Left = 0: shpType.Width = 960
        Case Else
            shpType.Width = 480
            shpType.Left = 960 + ((Index - 1) * 480)
    End Select
    picType.Visible = True
    shpType.Refresh
End Sub

Public Sub SetStatus(Index As Integer)
    Dim i As Integer, iOffset As Integer
    Dim bCheck As Boolean
    
    If bGPJ Then iOffset = 0 Else iOffset = 1
    
    shpStatus.Left = 840 * Index ''960 * Index
    Select Case Index
        Case 0
            shpStatus.Left = 0: shpStatus.Width = 840 ''960
        Case Else
            shpStatus.Width = 720
            shpStatus.Left = 840 + ((Index - 1 - iOffset) * 720) ''960 + ((Index - 1 - iOffset) * 720)
    End Select
    
    If bResetting Then Exit Sub
    
    Me.MousePointer = 11
    cmdStatusEdit_View.Enabled = True
    optApproverView(1).Enabled = True
    Select Case Index
        Case 0
            If bGPJ Then sIN = "10, 20, 27" Else sIN = "20, 27"
            lblMess.Caption = "...Refreshing with All Files..."
'''            bCheck = True
        Case 1
            sIN = "10"
            lblMess.Caption = "...Refreshing with 'Internal Draft' Files Only..."
'''            cmdStatusEdit_View.Enabled = True
        Case 2
            sIN = "20"
            lblMess.Caption = "...Refreshing with 'Client Draft' Files Only..."
'''            cmdStatusEdit_View.Enabled = True
        Case 3
            sIN = "27"
            lblMess.Caption = "...Refreshing with 'Returned' Files Only..."
        Case 4
            sIN = "30"
            lblMess.Caption = "...Refreshing with 'Approved' Files Only..."
'''            If iApproverView <> 0 Then
                bResetting = True
                iApproverView = 0
                optApproverView(0).Value = True
                optApproverView(1).Enabled = False
                bResetting = False
'''            End If
    End Select
    
    If cboCUNO(4).Text <> "" Then
        picMess.Visible = True: picMess.Refresh
        flxApprove.Visible = False: picOuter(4).Visible = False
        Call GetApprovalGraphics(CLng(fBCC(4)), sOrder, fSHYR(4), fSHCD(4), lFID)
        picMess.Visible = False
        flxApprove.Visible = True: picOuter(4).Visible = True
'''        If bCheck Then
'''            cmdStatusEdit_View.Enabled = True
'''            For i = 0 To flxApprove.Rows - 2
'''                If lblStat(i).Caption <> lblStat(0) Then
'''                    cmdStatusEdit_View.Enabled = False
'''                    Exit For
'''                End If
'''            Next i
'''        End If
    End If
    
    Me.MousePointer = 0
    
    
End Sub

Public Sub GetPartNodes(lEID As Long, sElem As String, strSelect As String, bNodes As Boolean)
    Dim rst As ADODB.Recordset
    Dim sENode As String, sPNode As String, sTNode As String
    Dim sDesc As String, sMess As String
    Dim nodX As Node
    Dim dVolT As Double, dWgtT As Double
    Dim bBadWgt As Boolean, bBadVol As Boolean
    Dim iWgtU As Integer, iVolU As Integer
    Dim sWgtU(1 To 2) As String
    Dim sUnit As String
    
    sWgtU(1) = " lbs": sWgtU(2) = " kg"
    If bNodes Then tvwGraphics(2).ImageList = ImageList1
    
    sENode = "e" & lEID
'''    sDesc = sElem
'''    Set nodX = tvwGraphics(2).Nodes.Add(, , sENode, sDesc, 3)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        iWgtU = rst.Fields("WTUNIT")
        iVolU = rst.Fields("SIZEUNIT")
        Select Case iVolU
            Case 1: sUnit = """"
            Case 2: sUnit = "'"
            Case 3: sUnit = "yd"
            Case 4: sUnit = "mi"
            Case 5: sUnit = "cm"
            Case 6: sUnit = "m"
            Case 7: sUnit = "km"
        End Select
    End If
    Do While Not rst.EOF
        sPNode = "p" & rst.Fields("PARTID")
        sDesc = "Part:  " & UCase(Trim(rst.Fields("PARTDESC")))
        If bNodes Then Set nodX = tvwGraphics(2).Nodes.Add(sENode, tvwChild, sPNode, sDesc, 23)
            sTNode = "t" & rst.Fields("PARTID") & "-1"
            sDesc = "Part No:  " & Trim(rst.Fields("PNUM"))
            If bNodes Then Set nodX = tvwGraphics(2).Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 24)
            
            sTNode = "t" & rst.Fields("PARTID") & "-2"
            sDesc = "Pkg Type:  " & Trim(rst.Fields("PKGTYPE"))
            If bNodes Then Set nodX = tvwGraphics(2).Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 24)
            
            sTNode = "t" & rst.Fields("PARTID") & "-3"
            sDesc = "Part Size:  L-" & rst.Fields("LENGTH") & sUnit & " x " & _
                        "W-" & rst.Fields("WIDTH") & sUnit & " x H-" & rst.Fields("HEIGHT") & sUnit
            If rst.Fields("LENGTH") = 0 Or rst.Fields("WIDTH") = 0 _
                        Or rst.Fields("HEIGHT") = 0 Then
                bBadVol = True
            Else
                dVolT = dVolT + (rst.Fields("LENGTH") * rst.Fields("WIDTH") * rst.Fields("HEIGHT"))
            End If
            If bNodes Then Set nodX = tvwGraphics(2).Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 24)
            
            sTNode = "t" & rst.Fields("PARTID") & "-4"
            sDesc = "Part Wgt:  " & rst.Fields("WEIGHT") & sWgtU(rst.Fields("WTUNIT"))
            If rst.Fields("WEIGHT") = 0 Then
                bBadWgt = True
            Else
                dWgtT = dWgtT + rst.Fields("WEIGHT")
            End If
            If bNodes Then Set nodX = tvwGraphics(2).Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 24)
            
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    sMess = ""
    Select Case iWgtU
        Case 1
            sMess = sMess & "Total Element Weight:  " & Format(dWgtT, "#,##0") & " lbs"
        Case 2
            sMess = sMess & "Total Element Weight:  " & Format(dWgtT, "#,##0") & " kg"
    End Select
    If bBadWgt Then
        sMess = sMess & "*" & vbNewLine
        sMess = sMess & "* Unweighed Parts found.  Actual weight may be greater."
    End If
    sMess = sMess & vbNewLine
    
    Select Case iVolU
        Case 1
            dVolT = dVolT / 1728
            sMess = sMess & "Total Element Volume:  " & Format(dVolT, "#,##0.00") & " cu ft"
        Case 5
            dVolT = dVolT / 1000000
            sMess = sMess & "Total Element Volume:  " & Format(dVolT, "#,##0.00") & " cu M"
    End Select
    If bBadVol Then
        sMess = sMess & "**" & vbNewLine
        sMess = sMess & "** Undimensioned Parts found.  Actual volume may be greater."
    End If
    
'''    lblMess.Caption = sMess
    MsgBox sMess, vbInformation, "Weight & Volume Totals..."

End Sub

Public Sub CheckIfGfxMandRecip(pBCC As Long, lUID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    bGfxMandRecip = False
    strSelect = "SELECT T.AN8_CUNO AS CUNO " & _
                "FROM " & ANOETeamUR & " R, " & ANOETeam & " T " & _
                "WHERE R.USER_SEQ_ID = " & lUID & " " & _
                "AND R.RECIPIENT_FLAG1 = 1 " & _
                "AND R.TEAM_ID = T.TEAM_ID " & _
                "AND T.AN8_CUNO = " & pBCC
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        bGfxMandRecip = True
'''        mnuGfxApproval.Visible = True
    Else
        bGfxMandRecip = False
'''        mnuGfxApproval.Visible = False
    End If
    rst.Close: Set rst = Nothing
End Sub

Public Sub CheckIfClientShowGfxMandRecip(pBCC As Long, pSHCD As Long, lUID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    bGfxMandRecip = False
    strSelect = "SELECT T.AN8_CUNO AS CUNO " & _
                "FROM " & ANOETeamUR & " R, " & ANOETeam & " T " & _
                "WHERE R.USER_SEQ_ID = " & lUID & " " & _
                "AND R.RECIPIENT_FLAG1 = 1 " & _
                "AND R.TEAM_ID = T.TEAM_ID " & _
                "AND T.AN8_CUNO = " & pBCC & " " & _
                "AND T.AN8_SHCD = " & pSHCD
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        bGfxMandRecip = True
'''        mnuGfxApproval.Visible = True
    Else
        bGfxMandRecip = False
'''        mnuGfxApproval.Visible = False
    End If
    rst.Close: Set rst = Nothing
End Sub

Public Sub GetApprovalShowYears(pBCC As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    cboSHYR(4).Clear
    cboASHCD.Clear
    cboASHCD.Enabled = False
    strSelect = "SELECT DISTINCT GS.SHYR " & _
                "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM " & _
                "Where GS.AN8_CUNO = " & pBCC & " " & _
                "AND GS.GID = GM.GID " & _
                "AND GM.GSTATUS IN (" & defEditSIN & ") " & _
                "ORDER BY SHYR"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        cboSHYR(4).Enabled = True
        Do While Not rst.EOF
            cboSHYR(4).AddItem rst.Fields("SHYR")
            rst.MoveNext
        Loop
        txtNoShows.Visible = False
    Else
        cboSHYR(4).Enabled = False
        cboASHCD.Enabled = False
        txtNoShows.Text = sNoShows
        txtNoShows.Visible = True
    End If
    rst.Close: Set rst = Nothing
End Sub

Public Sub GetApprovalShows(pBCC As Long, pSHYR As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    cboASHCD.Clear
    cboASHCD.AddItem "<All Client Files>"
    strSelect = "SELECT DISTINCT GS.AN8_SHCD, AB.ABALPH AS SHOW " & _
                "FROM ANNOTATOR.GFX_SHOW GS, ANNOTATOR.GFX_MASTER GM, " & F0101 & " AB " & _
                "Where GS.AN8_CUNO = " & pBCC & " " & _
                "AND GS.SHYR = " & pSHYR & " " & _
                "AND GS.GID = GM.GID " & _
                "AND GM.GSTATUS IN (" & defEditSIN & ") " & _
                "AND GS.AN8_SHCD = AB.ABAN8 " & _
                "ORDER BY UPPER(AB.ABALPH)"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        txtNoShows.Visible = False
        cboASHCD.Enabled = True
        Do While Not rst.EOF
            cboASHCD.AddItem Trim(rst.Fields("SHOW"))
            cboASHCD.ItemData(cboASHCD.NewIndex) = rst.Fields("AN8_SHCD")
            rst.MoveNext
        Loop
    Else
        txtNoShows.Text = sNoShows
        txtNoShows.Visible = True
        fSHYR(4) = 0
        cboASHCD.Enabled = False
        cboSHYR(4).Enabled = False
    End If
    rst.Close: Set rst = Nothing
    
    
    
End Sub

Public Sub ResetBatch(iStart As Integer, iCnt As Integer, CntIndex As Integer)
    Dim sMess As String
    
    If iStart = 1 Then
        lblFirst(CntIndex).Enabled = False
        lblPrevious(CntIndex).Enabled = False
    Else
        lblFirst(CntIndex).Enabled = True
        lblPrevious(CntIndex).Enabled = True
    End If
    
    If iCnt >= iStart + (iCols * iRows) Then
        lblNext(CntIndex).Enabled = True
        lblLast(CntIndex).Enabled = True
    Else
        lblNext(CntIndex).Enabled = False
        lblLast(CntIndex).Enabled = False
    End If
    
    If iCnt > (iCols * iRows) Then lblList(CntIndex).Visible = True Else lblList(CntIndex).Visible = False
    
    If iStart + (iCols * iRows) < iCnt Then
        lblCnt(CntIndex).Caption = "Images: " & iStart & " - " & iStart + ((iCols * iRows) - 1) & _
                    " of " & iCnt
    Else
        lblCnt(CntIndex).Caption = "Images: " & iStart & " - " & iCnt & _
                    " of " & iCnt
    End If
    
End Sub



Public Sub SetBatch(Index As Integer, sType As String)
    Select Case sType
        Case "PREVIOUS"
            iListStart(Index) = iListStart(Index) - (iCols * iRows)
            If iListStart(Index) < 1 Then iListStart(Index) = 1
            
        Case "NEXT"
            iListStart(Index) = iListStart(Index) + (iCols * iRows)
            If iListStart(Index) > iGFXCount(Index) Then _
                iListStart(Index) = (Int((iGFXCount(Index) - 1) / (iCols * iRows)) * (iCols * iRows)) + 1
            
        Case "FIRST"
            iListStart(Index) = 1
            
        Case "LAST"
            iListStart(Index) = (Int((iGFXCount(Index) - 1) / (iCols * iRows)) * (iCols * iRows)) + 1
            
    End Select
    
    Me.MousePointer = 11
    picWait.Visible = True
    picWait.Refresh
    
    TNode = tvwGraphics(Index).SelectedItem.Key
    picInner(Index).Visible = False
    Call GetGraphics(Index, Index, CurrSelect(Index), iListStart(Index), TNode)
    picInner(Index).Visible = True
    
    picWait.Visible = False
    Me.MousePointer = 0
    
    Call ResetBatch(iListStart(Index), iGFXCount(Index), 0)
    

End Sub

Public Sub ClearModes(Index As Integer)
    Dim i As Integer
    bDMode = False: mnuDownloadMode.Checked = False: mnuDownloadSels.Enabled = False: lblDownload.ForeColor = vbButtonText: mnuDownloadSels2.Visible = False
    bEMode = False: mnuEmailMode.Checked = False: mnuEmailSels.Enabled = False: lblEmail.ForeColor = vbButtonText: mnuEmailSels2.Visible = False
'    iModeTab = 0
    Select Case Index
        Case 0
            For i = 0 To imx0.Count - 1
                imx0(i).Enabled = True: chk0(i).Visible = False: chk0(i).Value = 0
            Next i
        Case 1
            For i = 0 To imx1.Count - 1
                imx1(i).Enabled = True: chk1(i).Visible = False: chk1(i).Value = 0
            Next i
        Case 2
            For i = 0 To imx2.Count - 1
                imx2(i).Enabled = True: chk2(i).Visible = False: chk2(i).Value = 0
            Next i
        Case 3
            For i = 0 To imx3.Count - 1
                imx3(i).Enabled = True: chk3(i).Visible = False: chk3(i).Value = 0
            Next i
        Case 4
            For i = 0 To imx4.Count - 1
                imx4(i).Enabled = True: chk4(i).Visible = False: chk4(i).Value = 0
            Next i
    End Select
'    iModeTab = 0
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


Public Sub GetApprovalFolders(pCUNO As Long, pSHCD As Long, pSHYR As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If pSHCD = 0 Then ''ALL SHOWS''
        If bGPJ Then
            strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC " & _
                        "FROM ANNOTATOR.GFX_FOLDER GF, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GF.AN8_CUNO = " & pCUNO & " " & _
                        "AND GF.FLR_ID = GM.FLR_ID " & _
                        "AND GF.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GM.GSTATUS IN (10, 20, 27) " & _
                        "ORDER BY FLRDESC"
        Else
            strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC " & _
                        "FROM ANNOTATOR.GFX_FOLDER GF, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GF.AN8_CUNO = " & pCUNO & " " & _
                        "AND GF.CLIENTRESTRICT_FLAG = 0 " & _
                        "AND GF.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GF.FLR_ID = GM.FLR_ID " & _
                        "AND GM.GSTATUS IN (20, 27) " & _
                        "ORDER BY FLRDESC"
        End If
    
    Else ''SPECIFIC SHOW''
        If bGPJ Then
            strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC " & _
                        "FROM ANNOTATOR.GFX_FOLDER GF, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_SHOW GS " & _
                        "Where GF.AN8_CUNO = " & pCUNO & " " & _
                        "AND GF.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GF.FLR_ID = GM.FLR_ID " & _
                        "AND GM.GSTATUS IN (10, 20, 27) " & _
                        "AND GM.GID = GS.GID " & _
                        "AND GS.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GS.SHYR = " & pSHYR & " " & _
                        "AND GS.AN8_SHCD = " & pSHCD & " " & _
                        "ORDER BY FLRDESC"
        Else
            strSelect = "SELECT DISTINCT GF.FLR_ID, GF.FLRDESC " & _
                        "FROM ANNOTATOR.GFX_FOLDER GF, ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_SHOW GS " & _
                        "Where GF.AN8_CUNO = " & pCUNO & " " & _
                        "AND GF.CLIENTRESTRICT_FLAG = 0 " & _
                        "AND GF.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GF.FLR_ID = GM.FLR_ID " & _
                        "AND GM.GSTATUS IN (20, 27) " & _
                        "AND GM.GID = GS.GID " & _
                        "AND GS.AN8_CUNO = GM.AN8_CUNO " & _
                        "AND GS.SHYR = " & pSHYR & " " & _
                        "AND GS.AN8_SHCD = " & pSHCD & " " & _
                        "ORDER BY FLRDESC"
        End If
    End If
    
    cboFolder.Clear
    cboFolder.AddItem "<All Client Folders>"
    cboFolder.ItemData(cboFolder.NewIndex) = 0
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboFolder.AddItem Trim(rst.Fields("FLRDESC"))
        cboFolder.ItemData(cboFolder.NewIndex) = rst.Fields("FLR_ID")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Private Sub txtRed_Change()
    lblRed(iRed).Caption = txtRed.Text
    If Len(txtRed.Text) > 5 Then lblEsc.Visible = False Else lblEsc.Visible = True
    bRedded = True
End Sub


Private Sub txtRed_KeyPress(KeyAscii As Integer)
    Debug.Print KeyAscii
    If KeyAscii = 27 Then
        lblEsc.Visible = False
'''        lblRed(iRed).ForeColor = vbGreen '' RGB(111, 175, 28) ''(255, 160, 0)
        lblRed(iRed).WordWrap = False
'        cmdRed(1).SetFocus
        bRedded = True
    ElseIf lblRed(iRed).Height + lblRed(iRed).Top > picRed.Height And KeyAscii <> 8 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub vsc1_Change()
    picRed.Top = picROuter.Height - vsc1.Value
End Sub

Private Sub vsc1_Scroll()
    picRed.Top = picROuter.Height - vsc1.Value
End Sub

Private Sub Xpdf1_MouseDown2(ByVal Button As Integer, ByVal Shift As Integer, ByVal page As Long, ByVal X As Double, ByVal Y As Double)
    xs = X: ys = Y
    Debug.Print "xS=" & xs & ":yS=" & ys
    
    If Button = 1 And bPan Then
        xpdf1.convertPDFToWindowCoords X, Y, panX, panY
    ElseIf Button = 1 And bSelMode Then
        xpdf1.convertPDFToWindowCoords X, Y, pxS, pyS
    ElseIf Button = vbRightButton Then
        Me.PopupMenu mnuRightClick
    End If
End Sub

Private Sub Xpdf1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim wX As Long, wY As Long
    
    If Button = 1 And bPan Then
        xpdf1.convertPDFToWindowCoords X, Y, wX, wY
        xpdf1.scrollBy -panSpeed * (wX - panX), -panSpeed * (wY - panY)
        panX = wX
        panY = wY
    End If
End Sub

Private Sub Xpdf1_MouseUp2(ByVal Button As Integer, ByVal Shift As Integer, ByVal page As Long, ByVal X As Double, ByVal Y As Double)
    xE = X: yE = Y
    Debug.Print "xE=" & xE & ":yE=" & yE
'    If Xpdf1.getCurrentSelection2(Xpdf1.currentPage, xS, yS, xE, yE) Then
'        Clipboard.SetData Xpdf1.convertRegionToPicture2(Xpdf1.currentPage, xS, yS, xE, yE, 72, Xpdf1.imageRGB)
'    End If
'''    If bSelMode Then
'''        Xpdf1.convertPDFToWindowCoords X, Y, pxE, pyE
'''        Clipboard.SetData Xpdf1.convertRegionToPicture2(Xpdf1.CurrentPage, pxS, pyS, pxE, pyE, 72, Xpdf1.imageRGB)
'''        bSelMode = False
'''        imgSelect.Picture = imlOpts.ListImages(2).Picture
'''    End If
End Sub

Private Sub Xpdf1_pageChange()
    iPDFPage = xpdf1.currentPage
    lblPage.Caption = "Page " & iPDFPage & " of " & xpdf1.NumPages
    RedName = lRedID & "-" & iPDFPage & "RED.bmp"
    RedFile = sGPath & RedName
    
    If xpdf1.NumPages = 1 Then
        ''DISABLE ALL''
        imgPage(0).Enabled = False
        imgPage(0).Picture = imlPageMode.ListImages(5).Picture
        imgPage(1).Enabled = False
        imgPage(1).Picture = imlPageMode.ListImages(7).Picture
    ElseIf xpdf1.NumPages > 1 And xpdf1.currentPage = 1 Then
        ''DISABLE BACK''
        imgPage(0).Enabled = False
        imgPage(0).Picture = imlPageMode.ListImages(5).Picture
        ''ENABLE FORWARD''
        imgPage(1).Enabled = True
        imgPage(1).Picture = imlPageMode.ListImages(8).Picture
    ElseIf xpdf1.NumPages = xpdf1.currentPage Then
        ''ENABLE BACK''
        imgPage(0).Enabled = True
        imgPage(0).Picture = imlPageMode.ListImages(6).Picture
        ''DISABLE FORWARD''
        imgPage(1).Enabled = False
        imgPage(1).Picture = imlPageMode.ListImages(7).Picture
    Else
        ''ENABLE ALL''
        imgPage(0).Enabled = True
        imgPage(0).Picture = imlPageMode.ListImages(6).Picture
        imgPage(1).Enabled = True
        imgPage(1).Picture = imlPageMode.ListImages(8).Picture
    End If
    
End Sub

Private Sub Xpdf1_selectDone()
'''    Debug.Print "Select"
    If bZWindow Then
        Dim x0 As Double, y0 As Double, X1 As Double, Y1 As Double
        xpdf1.enableMouseEvents = False
        If xpdf1.getCurrentSelection(x0, y0, X1, Y1) Then
            xpdf1.zoomToRect x0, y0, X1, Y1
        End If
        cboZoom.Text = CInt(xpdf1.zoomPercent) & "%"
        mnuMaxGraphic.Enabled = True
        If xpdf1.zoomPercent <> 100 Then
            mnuResizeGraphic.Enabled = True
        Else
            mnuResizeGraphic.Enabled = False
        End If
    Else
        xpdf1.copySelection
    End If
End Sub

Public Function GetNextLabel() As Integer
    Dim iVal As Integer
    If lblRed.Count = 1 And lblRed(0).Visible = False Then
        iVal = 0
    Else
        iVal = lblRed.Count
        Load lblRed(iVal)
        lblRed(iVal).Caption = ""
    End If
    GetNextLabel = iVal
End Function

Public Sub ClearLabels()
    Dim i As Integer
    iRed = 0
    On Error Resume Next
    For i = lblRed.UBound To 1 Step -1
        Unload lblRed(i)
    Next i
    lblRed(0).Caption = ""
'    Text1.Text = ""
    bRedMode = False: shpRN.Visible = False: lblRN.Visible = False
    bRedLine = False
    bRedText = False
    
    strSearchText = ""
    imgSearchPDF(1).Visible = False
    cmdFindNext.Default = False
    
    picRed.MousePointer = 0
End Sub


Public Sub PopZooms()
    lstZoom.Clear
''    lstZoom.AddItem "1%": lstZoom.ItemData(lstZoom.NewIndex) = "1"
''    lstZoom.AddItem "6.25%": lstZoom.ItemData(lstZoom.NewIndex) = "6.25"
''    lstZoom.AddItem "8.33%": lstZoom.ItemData(lstZoom.NewIndex) = "8.33"
''    lstZoom.AddItem "12.5%": lstZoom.ItemData(lstZoom.NewIndex) = "12.5"
    
    lstZoom.AddItem "10%": lstZoom.ItemData(lstZoom.NewIndex) = "10"
    lstZoom.AddItem "25%": lstZoom.ItemData(lstZoom.NewIndex) = "25"
    lstZoom.AddItem "33.33%": lstZoom.ItemData(lstZoom.NewIndex) = "33.33"
    lstZoom.AddItem "44%": lstZoom.ItemData(lstZoom.NewIndex) = "44"
    lstZoom.AddItem "50%": lstZoom.ItemData(lstZoom.NewIndex) = "50"
    lstZoom.AddItem "66.67%": lstZoom.ItemData(lstZoom.NewIndex) = "66.67"
    lstZoom.AddItem "75%": lstZoom.ItemData(lstZoom.NewIndex) = "75"
    lstZoom.AddItem "100%": lstZoom.ItemData(lstZoom.NewIndex) = "100"
    lstZoom.AddItem "125%": lstZoom.ItemData(lstZoom.NewIndex) = "125"
    lstZoom.AddItem "150%": lstZoom.ItemData(lstZoom.NewIndex) = "150"
    lstZoom.AddItem "200%": lstZoom.ItemData(lstZoom.NewIndex) = "200"
    lstZoom.AddItem "300%": lstZoom.ItemData(lstZoom.NewIndex) = "300"
    lstZoom.AddItem "400%": lstZoom.ItemData(lstZoom.NewIndex) = "400"
    lstZoom.AddItem "500%": lstZoom.ItemData(lstZoom.NewIndex) = "500"
    
''    lstZoom.AddItem "600%": lstZoom.ItemData(lstZoom.NewIndex) = "600"
''    lstZoom.AddItem "800%": lstZoom.ItemData(lstZoom.NewIndex) = "800"
''    lstZoom.AddItem "1200%": lstZoom.ItemData(lstZoom.NewIndex) = "1200"
''    lstZoom.AddItem "1600%": lstZoom.ItemData(lstZoom.NewIndex) = "1600"
''    lstZoom.AddItem "2400%": lstZoom.ItemData(lstZoom.NewIndex) = "2400"
''    lstZoom.AddItem "3200%": lstZoom.ItemData(lstZoom.NewIndex) = "3200"
''    lstZoom.AddItem "6400%": lstZoom.ItemData(lstZoom.NewIndex) = "6400"

End Sub


Public Sub SetPageMode(pMode As Integer)
    xpdf1.continuousMode = CBool(pMode)
    If xpdf1.NumPages > 1 Then
        imgPageMode(1).Enabled = True
        imgPageMode(1).Picture = imlPageMode.ListImages(4).Picture
    Else
        imgPageMode(1).Enabled = False
        imgPageMode(1).Picture = imlPageMode.ListImages(4).Picture
    End If
    lblPage.Caption = "Page 1 of " & xpdf1.NumPages
End Sub

Public Sub SetZoomMode(pMode As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        imgPDF(i).Picture = imlZoomMode.ListImages((i * 3) + 2).Picture
    Next i
    
'''    bZWindow = False
'''    Xpdf1.mouseCursor = imgCur(0).Picture
    Select Case pMode
        Case 0 ''FIT PAGE''
            bZWindow = False: bPan = False
            xpdf1.mouseCursor = imgCur(0).Picture
            imgPDF(0).Picture = imlZoomMode.ListImages(3).Picture
            xpdf1.Zoom = xpdf1.zoomPage
            
        Case 1 ''FIT WIDTH''
            bZWindow = False: bPan = False
            xpdf1.mouseCursor = imgCur(0).Picture
            imgPDF(1).Picture = imlZoomMode.ListImages(6).Picture
            xpdf1.Zoom = xpdf1.zoomWidth
            
        Case 2 ''ZOOM WINDOW''
            bZWindow = True: bPan = False
            xpdf1.mouseCursor = imgCur(2).Picture
            imgPDF(2).Picture = imlZoomMode.ListImages(9).Picture
            imgPDF(2).Enabled = True
            xpdf1.enableMouseEvents = False
            xpdf1.enableSelect = True
            
        Case 3 ''PAN MODE''
            bPan = True: bZWindow = False
            xpdf1.mouseCursor = imgCur(1).Picture
            imgPDF(3).Picture = imlZoomMode.ListImages(12).Picture
            xpdf1.enableMouseEvents = True
            xpdf1.enableSelect = False
            
    End Select
    
    dZF = xpdf1.zoomPercent
    If dZF >= 500 Then
        imgPDF(2).Picture = imlZoomMode.ListImages(7).Picture
        imgPDF(2).Enabled = False
        imgZoom(1).Picture = imlZoomMode.ListImages(15).Picture
        imgZoom(1).Enabled = False
    Else
        imgZoom(1).Picture = imlZoomMode.ListImages(16).Picture
        imgZoom(1).Enabled = True
    End If
        
End Sub

Public Sub SetNavCnt(i As Integer)
    Dim tFile As Integer, tCnt As Integer
    
    Select Case sst1.Tab
        Case 4: tFile = iImage(i) + 1
        Case Else: tFile = iListStart(i) + iImage(i)
    End Select
    tCnt = lstFiles(i).ListCount
    lblNavCnt.Caption = tFile & " of " & tCnt
    If tFile > 1 Then
        imgNav(0).Enabled = True
        imgNav(0).Picture = imlNav.ListImages(1).Picture
    Else
        imgNav(0).Enabled = False
        imgNav(0).Picture = imlNav.ListImages(2).Picture
    End If
    If tFile < tCnt Then
        imgNav(1).Enabled = True
        imgNav(1).Picture = imlNav.ListImages(3).Picture
    Else
        imgNav(1).Enabled = False
        imgNav(1).Picture = imlNav.ListImages(4).Picture
    End If
    picNav.Visible = True
End Sub

Public Sub SetupPDFRed(pRedMode As Integer, PRFile As String)
    Dim Resp As VbMsgBoxResult
    
    bRedding = False
    If bRedMode = False Then
        
        xpdf1.Visible = False
        If Me.WindowState <> 2 Then Me.WindowState = 2
        
        If xpdf1.Zoom <> xpdf1.zoomWidth Then
            xpdf1.Zoom = xpdf1.zoomWidth
            Call SetZoomMode(1)
        End If
        
        bRedMode = True: shpRN.Visible = True: lblRN.Visible = True
        
        If Dir(PRFile, vbNormal) = "" Or PRFile = "" Then
            imgSize.Picture = xpdf1.convertPageToPicture(xpdf1.currentPage, (72 * (xpdf1.zoomPercent / 100)))
            Call UpdateSCD(0, 0, 0)
        Else
            ''FOUND A REDLINE''
            Resp = MsgBox("There is a current Redline File.  Do you want to load it?", _
                        vbQuestion + vbYesNo, "Active Redline File...")
            If Resp = vbYes Then
                imgSize.Picture = LoadPicture(PRFile)
                Call UpdateSCD(0, 0, 1)
            Else
                imgSize.Picture = xpdf1.convertPageToPicture(xpdf1.currentPage, (72 * (xpdf1.zoomPercent / 100)))
                Call UpdateSCD(0, 0, 0)
            End If
            
        End If
        If imgSize.Width <= picRed.Width Then
            picRed.Width = imgSize.Width
            picRed.Height = imgSize.Height
            picRed.Picture = imgSize.Picture '' Xpdf1.convertPageToPicture(Xpdf1.currentPage, (72 * (Xpdf1.zoomPercent / 100)))
        Else
            picRed.Height = (picRed.Width / imgSize.Width) * imgSize.Height
            picRed.PaintPicture imgSize.Picture, 0, 0, picRed.Width, picRed.Height
        End If
        
        ''STORE INITIAL VIEW IN UNDO''
        Call AddToUndo("PICRED", -1)
        
        picPDFTools.Visible = False
            
        picROuter.Visible = True
        picRed.Visible = True
        
        Set picCurrentRed = picRed
        
        vsc1.Max = picRed.Height
        vsc1.Min = picROuter.ScaleHeight
        vsc1.Value = picROuter.ScaleHeight
        If vsc1.Max - vsc1.Value <= vsc1.Value Then
            vsc1.LargeChange = picROuter.ScaleHeight
        Else
            vsc1.LargeChange = vsc1.Max - vsc1.Value
        End If
    '        vsc1.LargeChange = vsc1.Value
        vsc1.Visible = True
        
    ''    picPDFOptions.Visible = False
        
        picRedTools.Visible = True
        Me.MousePointer = 0
    End If
    
    Select Case pRedMode
        Case 1 ''RedLine''
            picRed.MousePointer = 99
            imgRedMode(0).Picture = imlRedMode.ListImages(1).Picture
            imgRedMode(1).Picture = imlRedMode.ListImages(4).Picture
            bRedLine = True
            bRedText = False
        Case 2 ''RedText''
            If lblRed(0).Container.Name <> picRed.Name Then
                Call ResetLBLREDContainer(picRed)
            End If
            picRed.MousePointer = 3
            imgRedMode(1).Picture = imlRedMode.ListImages(3).Picture
            imgRedMode(0).Picture = imlRedMode.ListImages(2).Picture
            bRedLine = False
            bRedText = True
    End Select
End Sub

Public Sub BurnishIt(pPic As PictureBox)
    Dim i As Integer, iLR As Integer, iStart As Integer, iRow As Integer
    Dim lHgt As Long
    
    pPic.FontSize = lblRed(0).FontSize
    pPic.FontBold = lblRed(0).FontBold
''    ppic.ForeColor = vbRed
    For i = lblRed.UBound To lblRed.LBound Step -1
        If lblRed(i).Visible Then
            lblRed(i).Visible = False
            pPic.ForeColor = CLng(lblRed(i).Tag)
            pPic.CurrentX = lblRed(i).Left
            iStart = 1
            iRow = 0
            iLR = InStr(iStart, lblRed(i).Caption, Chr(10))
            Do While iLR <> 0
                pPic.CurrentX = lblRed(i).Left
                pPic.CurrentY = lblRed(i).Top + (270 * (iRow))
                pPic.Print Mid(lblRed(i).Caption, iStart, iLR - iStart)
                iStart = iLR + 1
                iLR = InStr(iStart, lblRed(i).Caption, Chr(10))
                iRow = iRow + 1
            Loop
            pPic.CurrentX = lblRed(i).Left
            pPic.CurrentY = lblRed(i).Top + (270 * (iRow))
            pPic.Print Mid(lblRed(i).Caption, iStart)
            lblRed(i).Visible = False
            
            If i > 0 Then Unload lblRed(i) Else lblRed(i).Caption = ""
            
        End If
    Next i
    
End Sub

Public Sub ResetLBLREDContainer(pContainer As PictureBox)
    Dim i As Integer
    
    ''FIRST CLEAR EXISTING LBLRED ARRAY''
    For i = lblRed.Count - 1 To 1 Step -1
        Unload lblRed(i)
    Next i
    
    lblRed(0).Caption = ""
    Set lblRed(0).Container = pContainer
    Set lblEsc.Container = pContainer
    Set txtRed.Container = pContainer
End Sub

Public Function ShallWeSave() As Integer
    Dim Resp As VbMsgBoxResult
    
    Resp = MsgBox("Redline file has been changed.  Do you want to Save your changes?", _
                vbExclamation + vbYesNoCancel, "Save...")
    Select Case Resp
        Case Is = vbYes
            Call mnuGRedSave_Click
            ShallWeSave = 1
        Case Is = vbNo
            bRedded = False
            ShallWeSave = 0
        Case Else
            ShallWeSave = 2
    End Select
    
End Function

Public Sub UpdateSCD(p0 As Integer, p1 As Integer, p2 As Integer)
    ''SAVE IMAGES 11-DISABLED, 12-ENABLED''
    ''CLEAR IMAGES 13-DISABLED, 14-ENABLED''
    ''DELETE IMAGES 15-DISABLED, 16-ENABLED''
    
    Select Case p0
        Case 0
            imgUtility(0).Picture = imlRedMode.ListImages(11).Picture
            imgUtility(0).Enabled = False
            mnuGRedSave.Enabled = False
        Case 1
            imgUtility(0).Picture = imlRedMode.ListImages(12).Picture
            imgUtility(0).Enabled = True
            mnuGRedSave.Enabled = True
    End Select
    Select Case p1
        Case 0
            imgUtility(1).Picture = imlRedMode.ListImages(13).Picture
            imgUtility(1).Enabled = False
            mnuGRedClear.Enabled = False
        Case 1
            imgUtility(1).Picture = imlRedMode.ListImages(14).Picture
            imgUtility(1).Enabled = True
            mnuGRedClear.Enabled = True
    End Select
    Select Case p2
        Case 0
            imgUtility(2).Picture = imlRedMode.ListImages(15).Picture
            imgUtility(2).Enabled = False
            mnuGRedDelete.Enabled = False
        Case 1
            imgUtility(2).Picture = imlRedMode.ListImages(16).Picture
            imgUtility(2).Enabled = True
            mnuGRedDelete.Enabled = True
    End Select
End Sub

Public Sub AddToUndo(sControl As String, Index As Integer)
'    MsgBox "iUndoIndex = " & Index
    
'    If iUndoListIndex < lstUndo.ListCount - 1 Then Call ResetUndoList(iUndoListIndex)
    
    If Index = -1 Then
        Call ClearUndo(1)
        Index = 0
    End If
    
    If UCase(sControl) <> "LBL" Then
        If Index < iUndoMax And Index >= imgUndo.Count Then
            Load imgUndo(Index)
            imgUndo(Index).Visible = False
        End If
        Select Case UCase(sControl)
            Case "PICRED"
                imgUndo(Index Mod iUndoMax).Picture = picRed.Image
                lstUndo.AddItem "img-" & iCurrUndo
                lstUndo.ItemData(lstUndo.NewIndex) = iCurrUndo - 1
            Case "PICJPG"
                imgUndo(Index Mod iUndoMax).Picture = picJPG.Image
                lstUndo.AddItem "img-" & iCurrUndo
                lstUndo.ItemData(lstUndo.NewIndex) = iCurrUndo - 1
        End Select
        
    ElseIf UCase(sControl) = "LBL" Then
        lstUndo.AddItem "lbl-" & Index
        lstUndo.ItemData(lstUndo.NewIndex) = Index
    End If
    
    
'''    If Left(UCase(sControl), 3) = "PIC" Then
'''        If Index < iUndoMax And Index >= imgUndo.Count Then
'''            Load imgUndo(Index)
'''            imgUndo(Index).Visible = False
'''        End If
'''        Select Case UCase(sControl)
'''            Case "PICRED"
'''                imgUndo(Index Mod iUndoMax).Picture = picRed.Image
'''            Case "PICJPG"
'''                imgUndo(Index Mod iUndoMax).Picture = picJPG.Image
'''        End Select
'''    End If
        
        
'''''        If bUndoCleared = True Then
'''''            ''THIS IS A FIRST TIME''
'''''            Select Case UCase(sControl)
'''''                Case "PICRED"
'''''                    imgUndo(Index).Picture = picRed.Image
'''''                    lstUndo.AddItem "imgUndo-0"
'''''                    lstUndo.ItemData(lstUndo.NewIndex) = 0
'''''                Case "PICJPG"
'''''                    imgUndo(0).Picture = picJPG.Image
'''''                    lstUndo.AddItem "imgUndo-0"
'''''                    lstUndo.ItemData(lstUndo.NewIndex) = 0
'''''            End Select
'''''            iCurrUndo = 0
'''''        Else
'''''            iCurrUndo = iCurrUndo
'''''            If Index = 0 Then
'''''                lstUndo.AddItem "imgUndo-" & CStr(imgUndo.UBound)
'''''                lstUndo.ItemData(lstUndo.NewIndex) = imgUndo.UBound
'''''                iCurrUndo = imgUndo.UBound
'''''            Else
'''''                lstUndo.AddItem "imgUndo-" & CStr(Index - 1)
'''''                lstUndo.ItemData(lstUndo.NewIndex) = Index - 1
'''''                iCurrUndo = Index
'''''            End If
'''''        End If
'''''    ElseIf UCase(sControl) = "LBLRED" Then
'''''
'''''    End If
        
    If lstUndo.ListCount > iUndoMax Then lstUndo.RemoveItem (0)
    
'''    If Left(UCase(sControl), 3) = "PIC" Then
'''        If Index < iUndoMax And Index >= imgUndo.Count Then
'''            Load imgUndo(Index)
'''            imgUndo(Index).Visible = False
'''        End If
'''        Select Case UCase(sControl)
'''            Case "PICRED"
'''                imgUndo(Index Mod iUndoMax).Picture = picRed.Image
'''            Case "PICJPG"
'''                imgUndo(Index Mod iUndoMax).Picture = picJPG.Image
'''        End Select
'''    End If
    
'    iUndoListIndex = lstUndo.ListCount - 1
    
'    iUndoIndex = Index + 1
'    If iUndoIndex = iUndoMax Then iUndoIndex = 0
    
    ''RESET UNDO ICONS''
    If lstUndo.ListCount > 1 Then
        imgDo(0).Picture = imlRedMode.ListImages(7).Picture
        imgDo(0).Enabled = True
    Else
        imgDo(0).Picture = imlRedMode.ListImages(8).Picture
        imgDo(0).Enabled = False
    End If
'''''    imgDo(1).Picture = imlRedMode.ListImages(10).Picture
'''''    imgDo(1).Enabled = False

    bUndoCleared = False
End Sub

Public Sub ClearUndo(Index As Integer)
    Dim i As Integer
    
    On Error Resume Next
    bUndoCleared = True
    iUndoIndex = 0
    iRed = 0
    iCurrUndo = 0
    For i = imgUndo.UBound To 1 Step -1
        Unload imgUndo(i)
    Next i
    imgUndo(0).Picture = LoadPicture("")
    imgDo(0).Picture = imlRedMode.ListImages(8).Picture
    imgDo(0).Enabled = False
    
    For i = lblRed.UBound To 1 Step -1
        Unload lblRed(i)
    Next i
    lblRed(0).Caption = ""
    lblRed(0).Visible = False
    
    lstUndo.Clear
    
'    If Index > 0 Then
'        If picJPG.Visible Then
'            Call AddToUndo("picjpg", -1)
'        ElseIf picRed.Visible Then
'            Call AddToUndo("picred", -1)
'        End If
'    End If
End Sub

Public Sub ResetUndo(pListIndex As Integer)
'    MsgBox "Undoing " & pListIndex '' - 1
    Select Case UCase(Left(lstUndo.List(pListIndex), 3))
'''        Case "PICJPG"
'''            picJPG.Picture = imgUndo(lstUndo.ItemData(pListIndex) Mod iUndoMax).Picture
'''        Case "PICRED"
'''            picRed.Picture = imgUndo(lstUndo.ItemData(pListIndex) Mod iUndoMax).Picture
        Case "LBL"
            lblRed(lstUndo.ItemData(pListIndex)).Visible = False
            lblRed(lstUndo.ItemData(pListIndex)).Caption = ""
            lstUndo.RemoveItem (lstUndo.ListCount - 1)
            
        Case "IMG"
            If picJPG.Visible Then
                picJPG.Picture = imgUndo(lstUndo.ItemData(pListIndex) Mod iUndoMax).Picture '' imgUndo(lstUndo.ItemData(pListIndex)).Picture
            ElseIf picRed.Visible Then
                picRed.Picture = imgUndo(lstUndo.ItemData(pListIndex) Mod iUndoMax).Picture
            End If
            lstUndo.RemoveItem (lstUndo.ListCount - 1)
            iCurrUndo = iCurrUndo - 1
'            iCurrUndo = lstUndo.ItemData(pListIndex)
'            MsgBox "Curr = " & iCurrUndo
'''            Set imgUndo(lstUndo.ItemData(pListIndex)).Picture = LoadPicture("")
    End Select
'''    iUndoIndex = iUndoIndex - 1
End Sub

Public Sub ResetUndoList(pListIndex As Integer)
    Dim i As Integer
    For i = lstUndo.ListCount - 1 To pListIndex Step -1
        Select Case UCase(lstUndo.List(i))
            Case "LBLRED"
                lblRed(lstUndo.ItemData(i)).Visible = False
                lblRed(lstUndo.ItemData(i)).Caption = ""
            Case Else
                imgUndo(lstUndo.ItemData(i)).Visible = False
                imgUndo(lstUndo.ItemData(i)).Picture = LoadPicture("")
        End Select
        lstUndo.RemoveItem (i)
        
    Next i
End Sub

Public Sub ResetJPGZoom(pResize As Integer, pMax As Integer, pFull As Integer)
    Dim rFactor As Double
    Select Case pResize
        Case 0
            imgJPGZoom(0).Picture = imlJPGTools.ListImages(1).Picture
            imgJPGZoom(0).Enabled = False
        Case 1
            imgJPGZoom(0).Picture = imlJPGTools.ListImages(2).Picture
            imgJPGZoom(0).Enabled = True
        Case 2
            imgJPGZoom(0).Picture = imlJPGTools.ListImages(3).Picture
            imgJPGZoom(0).Enabled = True
    End Select
    Select Case pMax
        Case 0
            imgJPGZoom(1).Picture = imlJPGTools.ListImages(4).Picture
            imgJPGZoom(1).Enabled = False
        Case 1
            imgJPGZoom(1).Picture = imlJPGTools.ListImages(5).Picture
            imgJPGZoom(1).Enabled = True
        Case 2
            imgJPGZoom(1).Picture = imlJPGTools.ListImages(6).Picture
            imgJPGZoom(1).Enabled = True
    End Select
    Select Case pFull
        Case 0
            imgJPGZoom(2).Picture = imlJPGTools.ListImages(7).Picture
            imgJPGZoom(2).Enabled = False
        Case 1
            imgJPGZoom(2).Picture = imlJPGTools.ListImages(8).Picture
            imgJPGZoom(2).Enabled = True
        Case 2
            imgJPGZoom(2).Picture = imlJPGTools.ListImages(9).Picture
            imgJPGZoom(2).Enabled = True
    End Select
    
    rFactor = picJPG.Width / imgSize.Width
    lblSize.Caption = CLng(rFactor * 100) & "%"
End Sub

Public Sub RedNoteVis(bSet As Boolean)
    shpRN.Visible = bSet
    lblRN.Visible = bSet
End Sub

Public Function GetPDFRedCount(pRID As Long) As String
    Dim strSelect As String, sMess As String
    Dim rst As ADODB.Recordset
    Dim bFirst As Boolean
    
    ''FIRST, CHECK COUNT''
    strSelect = "SELECT COUNT(REF_ID) AS RCNT " & _
                "FROM " & GFXRed & " " & _
                "WHERE REF_ID = " & pRID & " " & _
                "AND RED_STATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If rst.Fields("RCNT") > 1 Then
        sMess = "[" & rst.Fields("RCNT") & " Redline Files Found (Pages "
    ElseIf rst.Fields("RCNT") = 1 Then
        sMess = "[A Redline File Found (Page "
    Else
        rst.Close: Set rst = Nothing
        GetPDFRedCount = ""
        Exit Function
    End If
    rst.Close
    
    strSelect = "SELECT PAGE_ID " & _
                "FROM " & GFXRed & " " & _
                "WHERE REF_ID = " & pRID & " " & _
                "AND RED_STATUS > 0 " & _
                "ORDER BY PAGE_ID"
    Set rst = Conn.Execute(strSelect)
    bFirst = True
    Do While Not rst.EOF
        If bFirst Then
            sMess = sMess & rst.Fields("PAGE_ID")
            bFirst = False
        Else
            sMess = sMess & ", " & rst.Fields("PAGE_ID")
        End If
        rst.MoveNext
    Loop
    sMess = sMess & ")]"
    GetPDFRedCount = sMess
    
    rst.Close: Set rst = Nothing

End Function

Public Function CheckForRed(pRID As Long, pPID As Integer) As Boolean
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT REF_ID FROM " & GFXRed & " " & _
                "WHERE REF_ID = " & pRID & " " & _
                "AND PAGE_ID = " & pPID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        CheckForRed = True
    Else
        CheckForRed = False
    End If
    rst.Close: Set rst = Nothing
End Function
