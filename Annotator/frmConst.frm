VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConst 
   BackColor       =   &H00000000&
   Caption         =   "GPJ Engineering & Property Drawing Viewer"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmConst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":0B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":1118
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":16B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":19CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":1FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":2540
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sst1 
      Height          =   5955
      HelpContextID   =   2
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   10504
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      ShowFocusRect   =   0   'False
      BackColor       =   0
      TabCaption(0)   =   "Drawings of Show Props"
      TabPicture(0)   =   "frmConst.frx":2ADA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgIcon(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgIcon(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblName(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgIcon(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblEltID(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblKitID(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "volView(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tvwConst(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkClose(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdLoadDWF(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboSHCD"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboCUNO(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboSHYR"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lstBlocks(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdElemInfo(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "fraFP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Inventoried Drawings"
      TabPicture(1)   =   "frmConst.frx":2AF6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblName(1)"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "lblEltID(1)"
      Tab(1).Control(3)=   "lblKitID(1)"
      Tab(1).Control(4)=   "volView(1)"
      Tab(1).Control(5)=   "tvwConst(1)"
      Tab(1).Control(6)=   "chkClose(1)"
      Tab(1).Control(7)=   "cmdLoadDWF(1)"
      Tab(1).Control(8)=   "cboCUNO(1)"
      Tab(1).Control(9)=   "lstBlocks(1)"
      Tab(1).Control(10)=   "cmdElemInfo(1)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Engineering Projects"
      TabPicture(2)   =   "frmConst.frx":2B12
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCPRJ"
      Tab(2).Control(1)=   "chkClose(2)"
      Tab(2).Control(2)=   "cmdLoadDWF(2)"
      Tab(2).Control(3)=   "cboCUNO(2)"
      Tab(2).Control(4)=   "tvwConst(2)"
      Tab(2).Control(5)=   "volView(2)"
      Tab(2).Control(6)=   "lblName(2)"
      Tab(2).Control(7)=   "Label1"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Collaboration Projects"
      TabPicture(3)   =   "frmConst.frx":2B2E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdLoadDWF(3)"
      Tab(3).Control(1)=   "chkClose(3)"
      Tab(3).Control(2)=   "cboCUNO(3)"
      Tab(3).Control(3)=   "tvwConst(3)"
      Tab(3).Control(4)=   "volView(3)"
      Tab(3).Control(5)=   "lblName(3)"
      Tab(3).Control(6)=   "Label7"
      Tab(3).ControlCount=   7
      Begin VB.Frame fraFP 
         Caption         =   "Floorplan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   8220
         TabIndex        =   50
         Top             =   600
         Visible         =   0   'False
         Width           =   3200
         Begin VOLOVIEWXLibCtl.AvViewX volFP 
            Height          =   1335
            Left            =   180
            TabIndex        =   51
            Top             =   330
            Width           =   2930
            _cx             =   5168
            _cy             =   2355
            Appearance      =   0
            BorderStyle     =   0
            BackgroundColor =   "DefaultColors"
            Enabled         =   -1  'True
            UserMode        =   "ZoomToRect"
            HighlightLinks  =   0   'False
            src             =   ""
            LayersOn        =   ""
            LayersOff       =   ""
            SrcTemp         =   ""
            SupportPath     =   $"frmConst.frx":2B4A
            FontPath        =   $"frmConst.frx":2D7B
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
         Begin VB.Label lblFP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Click on Elem to view drawing"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   990
            TabIndex        =   52
            Top             =   30
            Width           =   2100
         End
      End
      Begin VB.CommandButton cmdElemInfo 
         Caption         =   "Info"
         Height          =   435
         Index           =   0
         Left            =   10380
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5355
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdElemInfo 
         Caption         =   "Info"
         Height          =   435
         Index           =   1
         Left            =   -64620
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   5355
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox lstBlocks 
         Height          =   255
         Index           =   0
         ItemData        =   "frmConst.frx":2FAC
         Left            =   2340
         List            =   "frmConst.frx":2FAE
         TabIndex        =   38
         Top             =   5535
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ListBox lstBlocks 
         Height          =   255
         Index           =   1
         Left            =   -71100
         TabIndex        =   37
         Top             =   5475
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdLoadDWF 
         Caption         =   "Load File into Viewer"
         Height          =   435
         Index           =   3
         Left            =   -66360
         TabIndex        =   34
         Top             =   5340
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Auto-Close with Selection"
         Height          =   195
         Index           =   3
         Left            =   -74820
         MaskColor       =   &H8000000F&
         TabIndex        =   31
         Top             =   5535
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         Index           =   3
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   765
         Width           =   7875
      End
      Begin VB.Frame fraCPRJ 
         Height          =   585
         Left            =   -71820
         TabIndex        =   26
         Top             =   555
         Width           =   4455
         Begin VB.CommandButton cmdGo 
            Caption         =   "Go!"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3900
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   180
            Width           =   435
         End
         Begin VB.TextBox txtCPRJ 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   27
            Top             =   180
            Width           =   1995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Project Number Search:"
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   240
            Width           =   1710
         End
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Auto-Close with Selection"
         Height          =   195
         Index           =   2
         Left            =   -74820
         MaskColor       =   &H8000000F&
         TabIndex        =   23
         Top             =   5535
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoadDWF 
         Caption         =   "Load File into Viewer"
         Height          =   435
         Index           =   2
         Left            =   -66870
         TabIndex        =   22
         Top             =   5355
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         Index           =   2
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1215
         Width           =   7875
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         Index           =   1
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   765
         Width           =   7935
      End
      Begin VB.ComboBox cboSHYR 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   765
         Width           =   855
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         Index           =   0
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   765
         Width           =   6915
      End
      Begin VB.ComboBox cboSHCD 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1305
         Width           =   7875
      End
      Begin VB.CommandButton cmdLoadDWF 
         Caption         =   "Load File into Viewer"
         Height          =   435
         Index           =   1
         Left            =   -66780
         TabIndex        =   10
         Top             =   5355
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton cmdLoadDWF 
         Caption         =   "Load File into Viewer"
         Height          =   435
         Index           =   0
         Left            =   8160
         TabIndex        =   9
         Top             =   5340
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Auto-Close with Selection"
         Height          =   195
         Index           =   1
         Left            =   -74820
         MaskColor       =   &H8000000F&
         TabIndex        =   5
         Top             =   5535
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Auto-Close with Selection"
         Height          =   195
         Index           =   0
         Left            =   180
         MaskColor       =   &H8000000F&
         TabIndex        =   3
         Top             =   5535
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.TreeView tvwConst 
         Height          =   3615
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   1740
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6376
         _Version        =   393217
         HideSelection   =   0   'False
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwConst 
         Height          =   4215
         Index           =   1
         Left            =   -74820
         TabIndex        =   6
         Top             =   1155
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   7435
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwConst 
         Height          =   3735
         Index           =   2
         Left            =   -74820
         TabIndex        =   24
         Top             =   1635
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6588
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwConst 
         Height          =   4215
         Index           =   3
         Left            =   -74820
         TabIndex        =   32
         Top             =   1155
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   7435
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VOLOVIEWXLibCtl.AvViewX volView 
         Height          =   4275
         Index           =   3
         Left            =   -66780
         TabIndex        =   35
         Top             =   675
         Visible         =   0   'False
         Width           =   3135
         _cx             =   5530
         _cy             =   7541
         Appearance      =   0
         BorderStyle     =   0
         BackgroundColor =   "DefaultColors"
         Enabled         =   -1  'True
         UserMode        =   "ZoomToRect"
         HighlightLinks  =   0   'False
         src             =   ""
         LayersOn        =   ""
         LayersOff       =   ""
         SrcTemp         =   ""
         SupportPath     =   $"frmConst.frx":2FB0
         FontPath        =   $"frmConst.frx":31DE
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
      Begin VOLOVIEWXLibCtl.AvViewX volView 
         Height          =   4275
         Index           =   2
         Left            =   -66780
         TabIndex        =   48
         Top             =   675
         Visible         =   0   'False
         Width           =   7875
         _cx             =   13891
         _cy             =   7541
         Appearance      =   0
         BorderStyle     =   0
         BackgroundColor =   "DefaultColors"
         Enabled         =   -1  'True
         UserMode        =   "ZoomToRect"
         HighlightLinks  =   0   'False
         src             =   ""
         LayersOn        =   ""
         LayersOff       =   ""
         SrcTemp         =   ""
         SupportPath     =   $"frmConst.frx":340C
         FontPath        =   $"frmConst.frx":363D
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
      Begin VOLOVIEWXLibCtl.AvViewX volView 
         Height          =   4275
         Index           =   1
         Left            =   -66780
         TabIndex        =   49
         Top             =   675
         Visible         =   0   'False
         Width           =   3135
         _cx             =   5530
         _cy             =   7541
         Appearance      =   0
         BorderStyle     =   0
         BackgroundColor =   "DefaultColors"
         Enabled         =   -1  'True
         UserMode        =   "ZoomToRect"
         HighlightLinks  =   0   'False
         src             =   ""
         LayersOn        =   ""
         LayersOff       =   ""
         SrcTemp         =   ""
         SupportPath     =   $"frmConst.frx":386E
         FontPath        =   $"frmConst.frx":3A9B
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
      Begin VOLOVIEWXLibCtl.AvViewX volView 
         Height          =   1695
         Index           =   0
         Left            =   8225
         TabIndex        =   53
         Top             =   2800
         Visible         =   0   'False
         Width           =   3135
         _cx             =   5530
         _cy             =   2990
         Appearance      =   0
         BorderStyle     =   0
         BackgroundColor =   "DefaultColors"
         Enabled         =   -1  'True
         UserMode        =   "ZoomToRect"
         HighlightLinks  =   0   'False
         src             =   ""
         LayersOn        =   ""
         LayersOff       =   ""
         SrcTemp         =   ""
         SupportPath     =   $"frmConst.frx":3CC8
         FontPath        =   $"frmConst.frx":3EF7
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
      Begin VB.Label lblKitID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   10710
         TabIndex        =   43
         Top             =   5115
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblEltID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   10740
         TabIndex        =   42
         Top             =   5775
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblKitID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   -64200
         TabIndex        =   41
         Top             =   5175
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblEltID 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   -64200
         TabIndex        =   40
         Top             =   5775
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   -65820
         TabIndex        =   36
         Top             =   5055
         UseMnemonic     =   0   'False
         Width           =   75
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   33
         Top             =   555
         Width           =   465
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   -65820
         TabIndex        =   25
         Top             =   5055
         UseMnemonic     =   0   'False
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   21
         Top             =   975
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   19
         Top             =   555
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Year:"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   555
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   1140
         TabIndex        =   16
         Top             =   555
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1095
         Width           =   450
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   4
         Left            =   5700
         Picture         =   "frmConst.frx":4126
         Stretch         =   -1  'True
         ToolTipText     =   "Kit Sheet Image"
         Top             =   5415
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   9180
         TabIndex        =   8
         Top             =   5055
         UseMnemonic     =   0   'False
         Width           =   75
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   -65820
         TabIndex        =   7
         Top             =   5055
         UseMnemonic     =   0   'False
         Width           =   75
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   1
         Left            =   6540
         Picture         =   "frmConst.frx":4270
         Stretch         =   -1  'True
         ToolTipText     =   "Engineering Construction Drawing"
         Top             =   5445
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Index           =   0
         Left            =   6120
         Picture         =   "frmConst.frx":47FA
         Stretch         =   -1  'True
         ToolTipText     =   "Floorplan Property Block"
         Top             =   5445
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VOLOVIEWXLibCtl.AvViewX volConst 
      Height          =   3795
      Left            =   1140
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   10215
      _cx             =   4212322
      _cy             =   4200998
      Appearance      =   0
      BorderStyle     =   0
      BackgroundColor =   "DefaultColors"
      Enabled         =   -1  'True
      UserMode        =   "Zoom"
      HighlightLinks  =   0   'False
      src             =   ""
      LayersOn        =   ""
      LayersOff       =   ""
      SrcTemp         =   ""
      SupportPath     =   $"frmConst.frx":4944
      FontPath        =   $"frmConst.frx":4B86
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
   Begin MSComctlLib.ImageList imlDirs 
      Left            =   5280
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":4DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":56A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":5F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConst.frx":6856
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   308
      MouseIcon       =   "frmConst.frx":7130
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   765
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmConst.frx":743A
      MousePointer    =   99  'Custom
      Picture         =   "frmConst.frx":7744
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Width           =   720
   End
   Begin VB.Label lblSettings 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Settings..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10958
      MouseIcon       =   "frmConst.frx":828E
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   720
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   765
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
      MouseIcon       =   "frmConst.frx":8598
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   180
      Width           =   510
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   0
      Left            =   10800
      Picture         =   "frmConst.frx":88A2
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   1
      Left            =   10800
      Picture         =   "frmConst.frx":8CE4
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   2
      Left            =   10800
      Picture         =   "frmConst.frx":9126
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgComm 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "frmConst.frx":9430
      MousePointer    =   99  'Custom
      Picture         =   "frmConst.frx":973A
      Top             =   660
      Visible         =   0   'False
      Width           =   480
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
      Left            =   1860
      MouseIcon       =   "frmConst.frx":9B7C
      TabIndex        =   11
      Top             =   780
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Construction Drawing Viewer is loading..."
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
      Left            =   1020
      TabIndex        =   0
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   4710
   End
   Begin VB.Image imgClose 
      Height          =   945
      Left            =   10800
      Top             =   0
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
   Begin VB.Image imgMenu 
      Height          =   570
      Left            =   0
      Picture         =   "frmConst.frx":9E86
      Top             =   600
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuRCPan 
         Caption         =   "Pan"
      End
      Begin VB.Menu mnuRCZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuRCZoomW 
         Caption         =   "Zoom Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRCFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuDash00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCCancel 
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
      Begin VB.Menu mnuRedlining 
         Caption         =   "Annotation"
         Begin VB.Menu mnuVRedlines 
            Caption         =   "Redlines"
            Begin VB.Menu mnuVRedLoad 
               Caption         =   "Load Redline File"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuVRedReturn 
               Caption         =   "Return to Original Drawing"
            End
            Begin VB.Menu mnuVRedSave 
               Caption         =   "Save Redline File"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuVRedClear 
               Caption         =   "Clear Redlines"
               Enabled         =   0   'False
            End
            Begin VB.Menu mnuVRedDelete 
               Caption         =   "Delete Redline File"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu mnuVSketch 
            Caption         =   "  Redline ""Sketch"" Mode"
         End
         Begin VB.Menu mnuVText 
            Caption         =   "  Redline ""Text"" Mode"
         End
         Begin VB.Menu mnuVRedEnd 
            Caption         =   "  End Redline Mode"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuVDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVNamedViews 
         Caption         =   "Named Views..."
      End
      Begin VB.Menu mnuVLayers 
         Caption         =   "Layers..."
      End
      Begin VB.Menu mnuVMainDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuVDisplay 
            Caption         =   "Default Colors"
            Checked         =   -1  'True
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
      Begin VB.Menu mnuVPrintSet 
         Caption         =   "Print Set"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyClip 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuDownloadPDF 
         Caption         =   "Download PDF..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDownloadZip 
         Caption         =   "Download DWG zip file"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help..."
      End
      Begin VB.Menu mnuVCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuElemInfo 
      Caption         =   "mnuElemInfo"
      Visible         =   0   'False
      Begin VB.Menu mnuVElemInfo 
         Caption         =   "View Element Information..."
      End
      Begin VB.Menu mnuPhoto 
         Caption         =   "View Photo..."
      End
   End
   Begin VB.Menu mnuFindElem 
      Caption         =   "mnuFindElem"
      Visible         =   0   'False
      Begin VB.Menu mnuFindElems 
         Caption         =   "Find Elements this DWG has been attached to..."
      End
   End
   Begin VB.Menu mnuDWFEdit 
      Caption         =   "mnuDWFEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuDWFData 
         Caption         =   "Posted by..."
      End
      Begin VB.Menu mnuDWFDelete 
         Caption         =   "Delete DWF file..."
      End
   End
   Begin VB.Menu mnuSort 
      Caption         =   "mnuSort"
      Visible         =   0   'False
      Begin VB.Menu mnuSortNodes 
         Caption         =   "Sort Drawings by Drawing Name"
         Index           =   0
      End
      Begin VB.Menu mnuSortNodes 
         Caption         =   "Sort Drawings by Post Date (descending)"
         Index           =   1
      End
      Begin VB.Menu mnuSortNodes 
         Caption         =   "Sort Drawings by Post Date (ascending)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuFPVolo 
      Caption         =   "mnuFPVolo"
      Visible         =   0   'False
      Begin VB.Menu mnuFPPan 
         Caption         =   "Pan"
      End
      Begin VB.Menu mnuFPZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuFPZoomW 
         Caption         =   "Zoom Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFPFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuFPDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFPCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iDisplay As Integer

Public sOpenFile As String
'''Dim sPath As String
Dim sDWFPath As String
Dim dLeft As Double, dRight As Double, dTop As Double, dBottom As Double
Dim dLeft2(0 To 3) As Double, dRight2(0 To 3) As Double, dTop2(0 To 3) As Double, dBottom2(0 To 3) As Double
Dim dLeftFP As Double, dRightFP As Double, dTopFP As Double, dBottomFP As Double
Dim bViewSet As Boolean, bLoading As Boolean, bTeam As Boolean, bDirsOpen As Boolean
Dim tSHYR As Integer
Dim tBCC As String
Dim tSHCD As Long
Dim tFBCN As String
Dim tSHNM As String
Dim sPrevMode As String, sNewWelcome As String, sNewName As String
Dim iD As Integer
Dim tRefID As Long, tRedID As Long, lRefID As Long, lRedID As Long
Dim RedName As String, RedFile As String, redBCC As String
Dim bSaveRed As Boolean, bReded As Boolean, bRightButton As Boolean, _
            bResortingNodes As Boolean
Dim sHDR(0 To 1) As String
Dim iTab As Integer
Dim tHDR As String
Dim tELTID As Long, tKitID As Long
Dim tDWGID(0 To 3) As Long, tSHTID(0 To 3) As Long, tDWFID(0 To 3) As Long
Dim lDWGID As Long, lSHTID As Long, lDWFID As Long

Dim sPDFFile As String



Public Property Get PassBCC() As String
    PassBCC = tBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As String)
    tBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = tFBCN
End Property
Public Property Let PassFBCN(ByVal vNewValue As String)
    tFBCN = vNewValue
End Property

Public Property Get PassSHNM() As String
    PassSHNM = tSHNM
End Property
Public Property Let PassSHNM(ByVal vNewValue As String)
    tSHNM = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = tSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    tSHYR = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = tSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    tSHCD = vNewValue
End Property




Private Sub cboCUNO_Click(Index As Integer)
    Dim i As Integer

    If cboCUNO(Index).Text <> "" And bLoading <> True Then
        On Error Resume Next
        Screen.MousePointer = 11
        tBCC = Right("00000000" & CStr(cboCUNO(Index).ItemData(cboCUNO(Index).ListIndex)), 8)
        tFBCN = GetBCN(tBCC)
        For i = 0 To cboCUNO.Count - 1
            If i <> Index Then
                If cboCUNO(i).Text = "" Then cboCUNO(i).Text = tFBCN
            End If
        Next i
        Select Case Index
            Case 0
                tvwConst(0).Nodes.Clear
                Call GetShows(cboSHCD, tSHYR, tBCC)
            Case 1: Call PopInventory(tBCC)
            Case 2: Call PopDrawings("CUNO", tBCC)
            Case 3
                If tBCC = "00057548" Or tBCC = "00054216" Then
                    Call PopAICHIDrawings(tBCC)
                Else
                    Call PopNonGPJDrawings(tBCC)
                End If
        End Select
        
        fraFP.Visible = False
        volView(Index).src = ""
        volView(Index).Visible = False
        cmdLoadDWF(Index).Visible = False
        lblName(Index).Caption = ""
        
'''        sst1.Width = 7000
                    
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cboSHCD_Click()
    If cboSHCD.Text <> "" Then
        volView(0).src = ""
        volView(0).Visible = False
        cmdLoadDWF(0).Visible = False
        lblName(0).Caption = ""
        
        tSHCD = cboSHCD.ItemData(cboSHCD.ListIndex)
        tSHNM = GetSHNM(tSHCD, tSHYR)
        Call PopUse(tBCC, tSHYR, tSHCD)
        fraFP.Visible = CheckForFloorplan(CLng(tBCC), tSHYR, tSHCD)
'        lblFP.Visible = volFP.Visible
    End If
End Sub

Private Sub cboSHYR_Change()
    If cboSHYR.Text <> "" Then
        tSHYR = CInt(cboSHYR.Text)
        Call GetShowClients(cboCUNO(0), tSHYR)
    End If
End Sub

Private Sub cboSHYR_Click()
    If cboSHYR.Text <> "" Then
        tSHYR = CInt(cboSHYR.Text)
        Call GetShowClients(cboCUNO(0), tSHYR)
    End If
End Sub

Private Sub imgDirs_Click()
    If sst1.Visible = False Then
        sst1.Visible = True
        imgDirs.ToolTipText = "Click to Close File Index"
'''        Set imgDirs.Picture = imlDirs.ListImages(4).Picture
        bDirsOpen = True
    Else
        sst1.Visible = False
        imgDirs.ToolTipText = "Click to Open File Index..."
'''        Set imgDirs.Picture = imlDirs.ListImages(2).Picture
        bDirsOpen = False
'''        sst1.Width = 7000 '''6120 '''4800
    End If
End Sub

Private Sub cmdElemInfo_Click(Index As Integer)
    With frmElemInfo
        .PassELTID = CLng(lblEltID(Index))
        .PassKITID = CLng(lblKitID(Index))
        .PassHDR = sHDR(Index)
        .Show 1
    End With
End Sub

Private Sub cmdGo_Click()
    Dim sCPRJ As String
    bLoading = True
    sCPRJ = txtCPRJ.Text
    Call PopDrawings("CPRJ", sCPRJ)
    cmdGo.Default = False
    bLoading = False
End Sub

Private Sub cmdLoadDWF_Click(Index As Integer)
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim Resp As VbMsgBoxResult
    
    
    '///// CHECK IF RED SHOULD BE SAVED \\\\\
    If bSaveRed = True Then
        Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNoCancel, "Redline Changes...")
        If Resp = vbYes Then
            mnuVRedSave_Click
        ElseIf Resp = vbNo Then
            volConst.ClearMarkup
            bSaveRed = False
        ElseIf Resp = vbCancel Then
            Exit Sub
        End If
    End If
    
    If bReded = True And bTeam = True And bPerm(15) Then
        With frmRedAlert
            .PassBCC = CLng(redBCC)
            .PassHDR = lblWelcome
            .PassType = 2
            .Show 1
        End With
'''        Call RedAlert(0, lblWelcome, redBCC, redSHCD) 'AlertOfRed
    End If
    bReded = False
    redBCC = ""
    
    sDWGZip = ""
    RedFile = ""
    
    sst1.Visible = False
'''    sst1.Width = 7000 '''6120 '''4800
'''''    lblDWF = lblName(Index).Caption
'''    picJPG.Visible = False
'''    lblByGeorge(0).Visible = False
'''    lblByGeorge(1).Visible = False
    bDirsOpen = False
    imgDirs.ToolTipText = "Click to Open File Index..."
'''    Set imgDirs.Picture = imlDirs.ListImages(1).Picture
            
    bViewSet = False
    volConst.src = volView(Index).src
    sPDFFile = Left(volView(Index).src, Len(volView(Index).src) - 3) & "pdf"
    
    mnuDownloadPDF.Enabled = CBool(Dir(sPDFFile, vbNormal) <> "")
    
    sOpenFile = tvwConst(Index).SelectedItem.Text
    lDWGID = tDWGID(Index): lSHTID = tSHTID(Index): lDWFID = tDWFID(Index)
    
    
    '///// CHECK FOR TEAM \\\\\
    mnuVRedReturn.Enabled = False
    mnuVRedClear.Enabled = False
    mnuVRedDelete.Enabled = False
    mnuVRedSave.Enabled = False
    If bPerm(34) Then
        bTeam = CheckForTeam(tBCC, tSHCD, frmConst)
        imgComm.Enabled = True
        If bTeam Then
             '///// CHECK FOR COMMENT \\\\\
             lRefID = lDWFID '' tRefID
             strSelect = "SELECT COMMID FROM " & ANOComment & " " & _
                        "WHERE REFID = " & lRefID & " " & _
                        "AND COMMSTATUS > 0"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                imgComm.Picture = imgMail(1).Picture
            Else
                imgComm.Picture = imgMail(0).Picture
            End If
            rst.Close: Set rst = Nothing
            
            '///// CHECK FOR REDLINE \\\\\'
            strSelect = "SELECT DWFID, DWFPATH FROM ANNOTATOR.DWG_DWF " & _
                        "Where DWGID = " & lDWGID & " " & _
                        "AND DWFTYPE = -9"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                RedFile = Trim(rst.Fields("DWFPATH"))
                lRedID = rst.Fields("DWFID")
                lblDWF.ForeColor = vbRed
                lblDWF.Caption = "A Redline File exists for this drawing.  " & _
                            "To Load it, use the Viewer Menu, or click this link."
                lblDWF.MousePointer = 99
                mnuVRedLoad.Enabled = True
                
            Else
                RedFile = "": lRedID = 0
                lblDWF.ForeColor = vbWhite
                lblDWF.Caption = sNewName
                lblDWF.MousePointer = 0
                mnuVRedLoad.Enabled = False
                mnuVRedReturn.Enabled = False
            End If
            rst.Close: Set rst = Nothing
            
            mnuRedlining.Enabled = True
            
            '///// CHECK FOR ZIP FILE \\\\\'
            sDWGZip = Left(volConst.src, Len(volConst.src) - 4) & ".zip"
            If Dir(sDWGZip, vbNormal) <> "" Then
                mnuDownloadZip.Enabled = True
            Else
                mnuDownloadZip.Enabled = False
                sDWGZip = ""
            End If
            
        Else
            imgComm.Picture = imgMail(2).Picture
            imgComm.Enabled = False
            mnuRedlining.Enabled = False
            lblDWF.ForeColor = vbWhite
        End If
        imgComm.Visible = True
    Else
        imgComm.Visible = False
        mnuRedlining.Enabled = False
        lblDWF.ForeColor = vbWhite
    End If
    
    imgMenu.Visible = True
    lblMenu.Visible = True
    
'    volView(Index).src = ""
    volConst.Visible = True
    lblWelcome = sNewWelcome
    
    lblWelcome.WordWrap = False
    If lblWelcome.Left + lblWelcome.Width > lblClose.Left - 240 Then
        lblWelcome.WordWrap = True
        lblWelcome.Width = lblClose.Left - 240 - lblWelcome.Left
        If lblWelcome.Height > shpHDR.Height Then
            lblWelcome.Height = shpHDR.Height
        End If
    End If
    lblWelcome.Top = (shpHDR.Height - lblWelcome.Height) / 2
            
''    lblDWF = sNewName
    
    lblSettings.Visible = True
End Sub

Private Sub Form_Load()
    Dim i As Integer, tmpSHYR As Integer
'''''    Dim Con As Control
'''''
'''''    On Error Resume Next
'''''    Me.BackColor = RGB(227, 229, 208)
'''''    For Each Con In Me.Controls
'''''        Con.BackColor = RGB(227, 229, 208)
'''''    Next
    
    Screen.MousePointer = 11
    
    lblFP.Caption = "Mouse over Element and click to view drawing"
    sDWFPath = "\\DETMSFS01\GPJAnnotator\Engineering\"
    
    sst1.Tab = 1
    
    mnuRedlining.Enabled = True '' = False '///// CHANGE LATER \\\\\
    
'''    sPath = "D:\DWF Test\"
'''    sPath = "\\DETMSFS01\GPJAnnotator\Engineering\"
'''    sPath = "\\Detsfa01\Annotator\Engineering\"
    
'''    If bAnnoOpen = True Then
        Me.WindowState = AppWindowState
'''        If Me.WindowState = 0 Then
'''            Me.Width = frmAnnotator.Width
'''            Me.Height = frmAnnotator.Height
'''            Me.Top = frmAnnotator.Top
'''            Me.Left = frmAnnotator.Left
'''        End If
'''    Else
'''
'''    End If
    
    With volConst
        .Top = 1440 ''675
        .Left = 1080 '' 120
    End With
'''    with picJPG
'''        .Top = 1440 '' 675
'''        .Left = 1080 '' 120
'''    End With
    
'''    lblByGeorge(0).ForeColor = lGeo_Back '' RGB(30, 30, 21)
'''    lblByGeorge(1).ForeColor = lGeo_Fore '' RGB(100, 100, 68)
    
        '///// FIRST, POP SHYR COMBOS \\\\\
    tmpSHYR = CInt(Format(Now, "YYYY"))
    For i = -2 To 2
        cboSHYR.AddItem tmpSHYR + i
    Next i
    
    '///// NEXT, POP INVENTORY CLIENT LIST \\\\\
    Call PopClientsWithInventory(cboCUNO(1))
    Call PopClientsWithEngProjects(cboCUNO(2))
    Call PopClientsWithNonGPJDwgs(cboCUNO(3))
    
    '///// NEXT, PASS IN SHYR (IF IT EXISTS) \\\\\
    Err = 0
    On Error Resume Next
    If tSHYR <> 0 Then
        cboSHYR.Text = tSHYR
        If Err Then
            cboSHYR.Text = tmpSHYR
            Err.Clear
        End If
        cboSHYR.Text = tSHYR
        If Err Then
            cboSHYR.Text = tmpSHYR
            Err.Clear
        End If
    Else
        cboSHYR.Text = tmpSHYR
    End If
    
    '///// NOW PASS IN CUNOS IF THEY EXIST \\\\\
    If tBCC <> "" And tFBCN <> "" Then
        cboCUNO(0).Text = tFBCN
        cboCUNO(1).Text = tFBCN
        cboCUNO(2).Text = tFBCN
    End If
    
    '///// NOW PASS IN SHCD IF COMING FROM FLOORPLAN \\\\\
    If tSHCD <> 0 And tSHNM <> "" Then
        cboSHCD.Text = tSHNM
    End If
   
'''    sst1.Width = 7000 '''6120 '''4800
    bViewSet = False
'''    picJPG.Visible = True
    volConst.Visible = False
    
    '///// CHECK PERMS \\\\\
    If Not bPerm(33) Then '/// REDLINING \\\
        mnuRedlining.Visible = False
        mnuVDash02.Visible = False
    End If
    
    If bPerm(70) Then ''HIDE GPJ-BASED TABS''
        sst1.TabVisible(0) = False
        sst1.TabVisible(1) = False
        sst1.TabVisible(2) = False
        sst1.TabVisible(3) = True
        iTab = 3
        If tvwConst(iTab).Nodes.Count = 0 And cboCUNO(iTab).ListCount = 1 Then
            cboCUNO(iTab).Text = cboCUNO(iTab).List(0)
        End If
        mnuDWFDelete.Visible = True
    Else
        sst1.TabVisible(0) = True
        sst1.TabVisible(1) = True
        sst1.TabVisible(2) = True
        sst1.TabVisible(3) = True
        If bPerm(36) Then
            mnuDWFDelete.Visible = True
        Else
            mnuDWFDelete.Visible = False
        End If
    End If
        
    '///// EDITDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(False)
    If bPerm(43) And bENABLE_PRINTERS Then  '/// PRINT CAPABILITY \\\
        mnuVPrint.Visible = True
        mnuVPrintSet.Visible = True
    Else
        mnuVPrint.Visible = False
        mnuVPrintSet.Visible = False
    End If
    '\\\\\ ---------------------------------------------------------- /////
    
    sst1.Visible = True
    imgDirs.ToolTipText = "Click to Close File Index"
    bDirsOpen = True
    
    
    
    lblWelcome = "The Construction Drawing Viewer is waiting for your selection..."
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    If Me.WindowState <> 1 Then
        shpHDR.Width = Me.ScaleWidth
        If Me.Width > 2400 And Me.Height > 2400 Then
            sst1.Width = Me.ScaleWidth
            sst1.Height = Me.ScaleHeight - sst1.Top
            
            volConst.Width = Me.ScaleWidth - volConst.Left - 120
            volConst.Height = Me.ScaleHeight - volConst.Top - 120
'''            picJPG.Width = Me.ScaleWidth - picJPG.Left - 120
'''            picJPG.Height = Me.ScaleHeight - picJPG.Left - 120
            
            imgClose.Left = Me.ScaleWidth - imgClose.Width
            lblSettings.Left = imgClose.Left + (imgClose.Width / 2) - (lblSettings.Width / 2)
            lblClose.Left = imgClose.Left + (imgClose.Width / 2) - (lblClose.Width / 2)
            
            
            lblWelcome.WordWrap = False
            If lblWelcome.Left + lblWelcome.Width > lblClose.Left - 240 Then
                lblWelcome.WordWrap = True
                lblWelcome.Width = lblClose.Left - 240 - lblWelcome.Left
                If lblWelcome.Height > shpHDR.Height Then
                    lblWelcome.Height = shpHDR.Height
                End If
            End If
            lblWelcome.Top = (shpHDR.Height - lblWelcome.Height) / 2
            
            
            If sst1.Width > 10000 And sst1.Height > 2400 Then
                tvwConst(0).Height = sst1.Height - tvwConst(0).Top - 180 '' tvwConst(i).Left
                tvwConst(0).Width = (sst1.Width / 2) - 180
                cboCUNO(0).Width = tvwConst(0).Left + tvwConst(0).Width - cboCUNO(0).Left
                cboSHCD.Width = tvwConst(0).Width
                volView(0).Width = (sst1.Width / 2) - 360
                volView(0).Left = tvwConst(0).Left + tvwConst(0).Width + 180
                volView(0).Height = (sst1.Height - sst1.TabHeight - 360 - (300 * 2) - (cmdLoadDWF(0).Height * 2)) / 2
                
                cmdLoadDWF(0).Top = sst1.Height - 180 - cmdLoadDWF(0).Height '' volView(0).Top + volView(0).Height + 300
                cmdLoadDWF(0).Left = volView(0).Left + (volView(0).Width / 2) - (cmdLoadDWF(0).Width / 2)
                lblName(0).Top = cmdLoadDWF(0).Top - 50 - lblName(0).Height '' volView(0).Top + volView(0).Height + 50
                lblName(0).Left = volView(0).Left + (volView(0).Width / 2) - (lblName(0).Width / 2)
                volView(0).Top = lblName(0).Top - 50 - volView(0).Height
                
                fraFP.Width = volView(0).Width
                fraFP.Height = volView(0).Height + 300
                fraFP.Left = volView(0).Left
                volFP.Width = fraFP.Width - 270 - 120 '' volView(0).Width
                volFP.Height = fraFP.Height - 270 - 135 - 120 '' volView(0).Height
'                volFP.Top = 720
'                volFP.Left = volView(0).Left
'                lblFP.Top = volFP.Top + volFP.Height + 60
'                lblFP.Left = volFP.Left + (volFP.Width / 2) - (lblFP.Width / 2)
                lblFP.Left = fraFP.Width - 120 - lblFP.Width
                
                For i = tvwConst.LBound + 1 To tvwConst.UBound
                    tvwConst(i).Height = sst1.Height - tvwConst(i).Top - 180 '' tvwConst(i).Left
'                    Debug.Print tvwConst(i).Height
                    volView(i).Width = sst1.Width - 8220 - 180
                    volView(i).Height = volView(i).Width / 3 * 2
                    lblName(i).Top = volView(i).Top + volView(i).Height + 50
                    lblName(i).Left = 8220 + (volView(i).Width / 2) - (lblName(i).Width / 2)
                    cmdLoadDWF(i).Top = volView(i).Top + volView(i).Height + 300
                    cmdLoadDWF(i).Left = 8220 + (volView(i).Width / 2) - (cmdLoadDWF(i).Width / 2)
                Next i
            End If
            cmdElemInfo(0).Left = cmdLoadDWF(0).Left + cmdLoadDWF(0).Width + 120
            cmdElemInfo(1).Left = cmdLoadDWF(1).Left + cmdLoadDWF(1).Width + 120
            cmdElemInfo(0).Top = cmdLoadDWF(0).Top
            cmdElemInfo(1).Top = cmdLoadDWF(1).Top
        End If
        AppWindowState = Me.WindowState
    End If
End Sub

Public Sub PopInventory(tmpBCC As String)
    Dim sDesc As String, sDescPar As String
    Dim sKNode As String, sENode As String, sDNode As String
    Dim nodX As Node
    Dim iType As Integer, ictr As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sChk As String
    
    sKNode = "": sENode = "": sDNode = ""
    tvwConst(1).ImageList = ImageList1
    tvwConst(1).Visible = False
    tvwConst(1).Nodes.Clear
    strSelect = "SELECT DISTINCT K.AN8_CUNO, K.KITID, K.KITREF, K.KITFNAME, " & _
                "E.ELTID, E.ELTFNAME, E.ELTCODE, E.ELSUFFIX, E.ELTDESC, " & _
                "PB.DWFID, PB.DWFTYPE, PB.DWFPATH " & _
                "FROM " & IGLKit & " K, " & _
                "" & IGLElt & " E, " & DWGElt & " DE, " & _
                "(select * FROM " & DWGDwf & " " & _
                "where dwftype = 20) PB " & _
                "WHERE K.AN8_CUNO = " & CLng(tBCC) & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ESTATUS > 2 " & _
                "AND E.ELTID = DE.INVID (+) " & _
                "AND DE.DWGID = PB.DWGID (+) " & _
                "ORDER BY K.AN8_CUNO, K.KITREF, E.ELTCODE, E.ELSUFFIX"
                
    Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
        If sKNode <> "k" & rst.Fields("KITID") Then
            sKNode = "k" & rst.Fields("KITID")
            sDesc = UCase(Trim(rst.Fields("KITFNAME")))
            sDescPar = sDesc
            iType = 5
            Set nodX = tvwConst(1).Nodes.Add(, , sKNode, sDesc, iType)
        End If
        If sENode <> "e" & rst.Fields("ELTID") Then
            '///// THIS IS A CHILD \\\\\
            sENode = "e" & rst.Fields("ELTID")
            sDesc = sDescPar & "-" & UCase(Trim(rst.Fields("ELTFNAME"))) & "  " & _
                        UCase(Trim(rst.Fields("ELTDESC")))
            If rst.Fields("DWFTYPE") = 20 And rst.Fields("DWFPATH") <> "" Then
                iType = 2
                lstBlocks(1).AddItem Trim(rst.Fields("DWFPATH"))
                lstBlocks(1).ItemData(lstBlocks(1).NewIndex) = rst.Fields("ELTID")
                If tvwConst(1).Nodes(sKNode).Image <> 1 Then tvwConst(1).Nodes(sKNode).Image = 4
            Else
                iType = 5
            End If
            Set nodX = tvwConst(1).Nodes.Add(sKNode, tvwChild, sENode, sDesc, iType)
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing

'''    strSelect = "SELECT DISTINCT K.AN8_CUNO, K.KITID, K.KITREF, K.KITFNAME, " & _
'''                "E.ELTID, E.ELTFNAME, E.ELTCODE, E.ELSUFFIX, E.ELTDESC, " & _
'''                "NVL(EE.PRGID, 0)PRGID, SAW.* " & _
'''                "FROM " & IGLKit & " K, " & _
'''                "" & IGLElt & " E, " & DWGElt & " DE, ENG_ELTID EL, ENG_ELEMENT EE, " & _
'''                "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
'''                "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
'''                "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'''                "WHERE DM.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                "AND DM.DWGTYPE = 10 " & _
'''                "AND DM.DWGID = DS.DWGID " & _
'''                "AND DM.DWGID = DD.DWGID " & _
'''                "AND DS.DWGID = DD.DWGID " & _
'''                "AND DS.SHTID = DD.SHTID) SAW " & _
'''                "WHERE K.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                "AND K.KSTATUS > 0 " & _
'''                "AND K.KITID = E.KITID " & _
'''                "AND E.ESTATUS > 2 " & _
'''                "AND E.ELTID = EL.ELTID (+) " & _
'''                "AND EL.ELEMID = EE.ELEMID (+) " & _
'''                "AND E.ELTID = DE.INVID " & _
'''                "AND DE.DWGID = SAW.DWGID " & _
'''                "ORDER BY K.AN8_CUNO, K.KITREF, E.ELTCODE, E.ELSUFFIX,  " & _
'''                "PRGID, SAW.DWGNUM, SAW.SHTSEQ"
    
'''    strSelect = "SELECT DISTINCT K.AN8_CUNO, K.KITID, K.KITREF, K.KITFNAME, " & _
'''                "E.ELTID, E.ELTFNAME, E.ELTCODE, E.ELSUFFIX, E.ELTDESC, NVL(EE.PRGID, 0)PRGID, SAW.* " & _
'''                "FROM IGL_KIT K, IGL_ELEMENT E, DWG_ELEMENT DE, ENG_ELTID EL, ENG_ELEMENT EE, " & _
'''                "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
'''                "FROM DWG_MASTER DM, DWG_SHEET DS, DWG_DWF DD " & _
'''                "Where DM.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                "AND DM.DWGTYPE = 10 " & _
'''                "AND DM.DWGID = DS.DWGID " & _
'''                "AND DM.DWGID = DD.DWGID " & _
'''                "AND DS.DWGID = DD.DWGID " & _
'''                "AND DS.SHTID = DD.SHTID) SAW " & _
'''                "Where K.AN8_CUNO = " & CLng(tBCC) & " " & _
'''                "AND K.KSTATUS > 0 " & _
'''                "AND K.KITID = E.KITID " & _
'''                "AND E.ESTATUS > 2 " & _
'''                "AND E.ELTID = DE.INVID " & _
'''                "AND DE.DWGID = SAW.DWGID " & _
'''                "AND E.ELTID = EL.ELTID (+) " & _
'''                "AND EL.ELEMID = EE.ELEMID (+) " & _
'''                "ORDER BY K.AN8_CUNO, K.KITREF, E.ELTCODE, E.ELSUFFIX,  " & _
'''                "PRGID, SAW.DWGNUM, SAW.SHTSEQ"

    strSelect = "SELECT DISTINCT K.AN8_CUNO, K.KITID, K.KITREF, K.KITFNAME, " & _
                "E.ELTID, E.ELTFNAME, E.ELTCODE, E.ELSUFFIX, E.ELTDESC, NVL(EE.PRGID, 0)PRGID, " & _
                "DM.DWGID, NVL(EE.DWGID, -1)DWG_CHK, DM.DWGDESC, DS.SHTDESC, DM.DWGYR , DM.MCU, " & _
                "DM.DWGNUM , dS.SHTID, dS.SHTSEQ, DD.DWFID " & _
                "FROM IGLPROD.IGL_KIT K, IGLPROD.IGL_ELEMENT E, ANNOTATOR.DWG_ELEMENT DE, " & _
                "ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DD, " & _
                "ANNOTATOR.ENG_ELTID EL, ANNOTATOR.ENG_ELEMENT EE " & _
                "Where K.AN8_CUNO = " & CLng(tBCC) & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ESTATUS > 2 " & _
                "AND E.ELTID = DE.INVID " & _
                "AND DE.DWGID = DM.DWGID " & _
                "AND DM.DWGTYPE = 10 " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DD.DWGID " & _
                "AND DS.SHTID = DD.SHTID " & _
                "AND DD.DWFTYPE = -1 " & _
                "AND E.ELTID = EL.ELTID (+) " & _
                "AND EL.ELEMID = EE.ELEMID (+) " & _
                "ORDER BY K.AN8_CUNO, K.KITREF, E.ELTCODE, E.ELSUFFIX, PRGID, DWG_CHK DESC, DM.DWGNUM, DS.SHTSEQ"

    Set rst = Conn.Execute(strSelect)
    iType = 3
    ictr = 1
    Do While Not rst.EOF
        sENode = "e" & rst.Fields("ELTID")
        sDNode = "d" & ictr & "-" & rst.Fields("DWFID")
'        If rst.Fields("PRGID") = 1163 Then
'            Debug.Print "STOP"
'        End If
        If rst.Fields("PRGID") = 0 Or rst.Fields("DWGID") <> rst.Fields("DWG_CHK") Then
            Select Case Len(Trim(rst.Fields("MCU")))
                Case 6
                    sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 2, 4) & "-" & _
                                Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                Case 9
                    sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 4) & "-" & _
                                Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                Case 10
                    sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 5) & "-" & _
                                Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                Case Else
                    sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Trim(rst.Fields("MCU")) & "-" & _
                                Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                "  [" & Trim(rst.Fields("SHTDESC")) & "]"
            End Select
        Else
            sDesc = rst.Fields("PRGID") & "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & _
                        Trim(rst.Fields("SHTSEQ")) & "  [" & Trim(rst.Fields("SHTDESC")) & "]"
        End If
        
        Set nodX = tvwConst(1).Nodes.Add(sENode, tvwChild, sDNode, sDesc, iType)
        If nodX.Parent.Image <> 2 Then nodX.Parent.Image = 4
        If nodX.Parent.Parent.Image <> 1 Then nodX.Parent.Parent.Image = 4
        ictr = ictr + 1
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    Set nodX = Nothing
    tvwConst(1).Visible = True
End Sub

Public Sub PopDrawings(sType As String, tmpStr As String)
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim nodX As Node
    Dim sYNode As String, sPNode As String, sSNode As String, sDNode As String
    Dim sDesc As String, sYear As String, sCPRJ As String
    Dim iType As Integer, i As Integer
    
    tvwConst(2).Nodes.Clear
    tvwConst(2).ImageList = ImageList1
    sYNode = "": sPNode = "": sSNode = "": sDNode = ""
    Select Case sType
        Case "CUNO"
'''            strSelect = "SELECT DM.DWGID, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
'''                        "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
'''                        "DD.DWFID, DD.DWFPATH " & _
'''                        "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'''                        "WHERE DM.AN8_CUNO = " & CLng(tmpStr) & " " & _
'''                        "AND DM.DWGTYPE IN (10, 11, 12) " & _
'''                        "AND DM.DWGID = DS.DWGID " & _
'''                        "AND DS.DWGID = DD.DWGID " & _
'''                        "AND DS.SHTID = DD.SHTID " & _
'''                        "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
            
            strSelect = "SELECT NVL(EE.PRGID, 0)PRGID, DM.DWGID, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
                        "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, " & _
                        "dS.SHTDESC , DD.DWFID, DD.DWFPATH " & _
                        "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DD, ANNOTATOR.ENG_ELEMENT EE " & _
                        "Where DM.AN8_CUNO = " & CLng(tmpStr) & " " & _
                        "AND DM.DWGTYPE IN (10, 11, 12) " & _
                        "AND DM.DWGID = DS.DWGID " & _
                        "AND DS.DWGID = DD.DWGID " & _
                        "AND DS.SHTID = DD.SHTID " & _
                        "AND DM.DWGID = EE.DWGID (+) " & _
                        "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
        Case "CPRJ"
            If bClientAll_Enabled Then
                strSelect = "SELECT NVL(EE.PRGID, 0)PRGID, DM.DWGID, DM.AN8_CUNO, C.ABALPH, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
                            "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
                            "DD.DWFID, DD.DWFPATH " & _
                            "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD, " & F0101 & " C, ANNOTATOR.ENG_ELEMENT EE " & _
                            "WHERE DM.MCU = '" & tmpStr & "' " & _
                            "AND DM.DWGTYPE IN (10, 11, 12) " & _
                            "AND DM.DWGID = DS.DWGID " & _
                            "AND DS.DWGID = DD.DWGID " & _
                            "AND DS.SHTID = DD.SHTID " & _
                            "AND DM.AN8_CUNO = C.ABAN8 " & _
                            "AND DM.DWGID = EE.DWGID (+) " & _
                            "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
                            
'''                strSelect = "SELECT DM.DWGID, DM.AN8_CUNO, C.ABALPH, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
'''                            "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
'''                            "DD.DWFID, DD.DWFPATH " & _
'''                            "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD, " & F0101 & " C " & _
'''                            "WHERE DM.MCU = '" & tmpStr & "' " & _
'''                            "AND DM.DWGTYPE IN (10, 11, 12) " & _
'''                            "AND DM.DWGID = DS.DWGID " & _
'''                            "AND DS.DWGID = DD.DWGID " & _
'''                            "AND DS.SHTID = DD.SHTID " & _
'''                            "AND DM.AN8_CUNO = C.ABAN8 " & _
'''                            "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
                            
            Else
                strSelect = "SELECT NVL(EE.PRGID, 0)PRGID, DM.DWGID, DM.AN8_CUNO, C.ABALPH, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
                            "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
                            "DD.DWFID, DD.DWFPATH " & _
                            "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD, " & F0101 & " C, ANNOTATOR.ENG_ELEMENT EE " & _
                            "WHERE DM.AN8_CUNO IN (" & strCunoList & ") " & _
                            "AND DM.MCU = '" & tmpStr & "' " & _
                            "AND DM.DWGTYPE IN (10, 11, 12) " & _
                            "AND DM.DWGID = DS.DWGID " & _
                            "AND DS.DWGID = DD.DWGID " & _
                            "AND DS.SHTID = DD.SHTID " & _
                            "AND DM.AN8_CUNO = C.ABAN8 " & _
                            "AND DM.DWGID = EE.DWGID (+) " & _
                            "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
                            
'''                strSelect = "SELECT DM.DWGID, DM.AN8_CUNO, C.ABALPH, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
'''                            "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
'''                            "DD.DWFID, DD.DWFPATH " & _
'''                            "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD, " & F0101 & " C " & _
'''                            "WHERE DM.AN8_CUNO IN (" & strCunoList & ") " & _
'''                            "AND DM.MCU = '" & tmpStr & "' " & _
'''                            "AND DM.DWGTYPE IN (10, 11, 12) " & _
'''                            "AND DM.DWGID = DS.DWGID " & _
'''                            "AND DS.DWGID = DD.DWGID " & _
'''                            "AND DS.SHTID = DD.SHTID " & _
'''                            "AND DM.AN8_CUNO = C.ABAN8 " & _
'''                            "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
            End If
'''            strSelect = "SELECT DM.DWGID, DM.AN8_CUNO, C.ABALPH, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
'''                        "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
'''                        "DD.DWFID, DD.DWFPATH " & _
'''                        "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD, " & F0101 & " C " & _
'''                        "WHERE DM.MCU = '" & tmpStr & "' " & _
'''                        "AND DM.DWGTYPE IN (10, 11, 12) " & _
'''                        "AND DM.DWGID = DS.DWGID " & _
'''                        "AND DS.DWGID = DD.DWGID " & _
'''                        "AND DS.SHTID = DD.SHTID " & _
'''                        "AND DM.AN8_CUNO = C.ABAN8 " & _
'''                        "ORDER BY DM.DWGYR, DM.MCU, DM.DWGTYPE, DM.DWGNUM, DS.SHTSEQ"
    End Select
    Debug.Print strSelect
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        On Error Resume Next
        cboCUNO(2).Text = UCase(Trim(rst.Fields("ABALPH")))
        tBCC = Right("00000000" & CStr(cboCUNO(2).ItemData(cboCUNO(2).ListIndex)), 8)
        Do While Not rst.EOF
            If sYNode <> "y" & rst.Fields("DWGYR") Then
                sYNode = "y" & rst.Fields("DWGYR")
                sDesc = rst.Fields("DWGYR")
                iType = 4
                Set nodX = tvwConst(2).Nodes.Add(, , sYNode, sDesc, iType)
                sYear = rst.Fields("DWGYR"): sPNode = "": sSNode = "": sDNode = ""
            End If
            If sPNode <> "p" & sYear & "-" & Trim(rst.Fields("MCU")) Then
                sPNode = "p" & sYear & "-" & Trim(rst.Fields("MCU"))
                sDesc = Trim(rst.Fields("MCU")) ''' & " -- " & Trim(rst.Fields("PROJ"))
                iType = 4
                Set nodX = tvwConst(2).Nodes.Add(sYNode, tvwChild, sPNode, sDesc, iType)
                sSNode = "": sDNode = ""
            End If
            If sSNode <> "s" & rst.Fields("DWGID") Then
                sSNode = "s" & rst.Fields("DWGID")
                sDesc = UCase(Trim(rst.Fields("DWGDESC")))
                iType = 4
                Set nodX = tvwConst(2).Nodes.Add(sPNode, tvwChild, sSNode, sDesc, iType)
                sDNode = ""
            End If
            
            If sDNode <> "d" & rst.Fields("DWFID") Then
                sDNode = "d" & rst.Fields("DWFID")
                Select Case rst.Fields("PRGID")
                    Case 0
                        Select Case Len(Trim(rst.Fields("MCU")))
                            Case 6
                                sDesc = UCase(Trim(rst.Fields("SHTDESC"))) & "  [" & _
                                            Right(CStr(rst.Fields("DWGYR")), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 2, 4) & _
                                            "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & "]"
                            Case 9
                                sDesc = UCase(Trim(rst.Fields("SHTDESC"))) & "  [" & _
                                            Right(CStr(rst.Fields("DWGYR")), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 4) & _
                                            "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & "]"
                            Case 10
                                sDesc = UCase(Trim(rst.Fields("SHTDESC"))) & "  [" & _
                                            Right(CStr(rst.Fields("DWGYR")), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 5) & _
                                            "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & "]"
                            Case Else
                                sDesc = UCase(Trim(rst.Fields("SHTDESC"))) & "  [" & _
                                            Right(CStr(rst.Fields("DWGYR")), 2) & "-" & Trim(rst.Fields("MCU")) & _
                                            "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & "]"
                        End Select
                    Case Else
                        sDesc = UCase(Trim(rst.Fields("SHTDESC"))) & "  [" & _
                                    rst.Fields("PRGID") & "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & _
                                    Trim(rst.Fields("SHTSEQ")) & "]"
                End Select
                iType = 3
                Set nodX = tvwConst(2).Nodes.Add(sSNode, tvwChild, sDNode, sDesc, iType)
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    ''GET PROJECT NAMES FROM JDE''
    For i = 1 To tvwConst(2).Nodes.Count
        If UCase(Left(tvwConst(2).Nodes(i).Key, 1)) = "P" Then
            strSelect = "SELECT MCDL01 AS PROJ " & _
                        "FROM " & F0006 & " " & _
                        "WHERE MCMCU = '" & Right(Space(12) & Mid(tvwConst(2).Nodes(i).Key, 7), 12) & "'"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                tvwConst(2).Nodes(i).Text = tvwConst(2).Nodes(i).Text & " -- " & Trim(rst.Fields("PROJ"))
            End If
            rst.Close: Set rst = Nothing
        End If
    Next i
End Sub

Public Sub PopUse(tmpBCC As String, tmpSHYR As Integer, tmpSHCD As Long)
    Dim sNode As String, sParent As String, sDesc As String, sDescPar As String
    Dim sKNode As String, sENode As String, sDNode As String
    Dim nodX As Node
    Dim iFile As Integer, i As Integer, iNode As Integer, iNo As Integer, ictr As Integer
    Dim lParent As Long
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sCPath As String, sChk As String, sClient As String
    Dim iType As Integer
    Dim sCList As String
    Dim lCUNO As Long
    
    Screen.MousePointer = 11
    sKNode = "": sENode = "": sDNode = ""
    sNode = "": sParent = ""
    tvwConst(0).Visible = False
    tvwConst(0).Nodes.Clear
    tvwConst(0).ImageList = ImageList1
    sCList = "": lCUNO = 0 '' CLng(tmpBCC)
    
'''    strSelect = "SELECT K.AN8_CUNO, K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
'''                "PB.DWFID, PB.DWFTYPE, PB.DWFPATH " & _
'''                "FROM " & IGLKitU & " KU, " & IGLEltU & " EU, " & IGLKit & " K, " & _
'''                "" & DWGElt & " DE, " & _
'''                "(select * FROM " & DWGDwf & " " & _
'''                "where dwftype = 20) PB " & _
'''                "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
'''                "AND KU.SHYR = " & tmpSHYR & " " & _
'''                "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'''                "AND KU.KITUSEID = EU.KITUSEID " & _
'''                "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
'''                "AND KU.SHYR = EU.SHYR " & _
'''                "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
'''                "AND EU.ESTATUS > 2 " & _
'''                "AND EU.SHSTATUS <> 3 " & _
'''                "AND EU.KITID = K.KITID " & _
'''                "AND EU.ELTID = DE.INVID (+) " & _
'''                "AND DE.DWGID = PB.DWGID (+) " & _
'''                "ORDER BY K.AN8_CUNO, K.KITREF, EU.ELTCODE, EU.ELSUFFIX"
                
    strSelect = "SELECT K.AN8_CUNO, AB.ABALPH AS CLIENT, K.KITID, K.KITFNAME, " & _
                "EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
                "PB.DWFID, PB.DWFTYPE, PB.DWFPATH " & _
                "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & IGLKit & " K, " & _
                "ANNOTATOR.ENG_ELTID EL, ANNOTATOR.ENG_ELEMENT EE, " & F0101 & " AB, " & _
                "(select * FROM " & DWGDwf & " " & _
                "where dwftype = 20) PB " & _
                "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND KU.SHYR = " & tmpSHYR & " " & _
                "AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                "AND KU.KITUSEID = EU.KITUSEID " & _
                "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                "AND KU.SHYR = EU.SHYR " & _
                "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                "AND EU.ESTATUS > 2 " & _
                "AND EU.SHSTATUS <> 3 " & _
                "AND EU.KITID = K.KITID " & _
                "AND K.AN8_CUNO = AB.ABAN8 " & _
                "AND EU.ELTID = EL.ELTID (+) " & _
                "AND EL.ELEMID = EE.ELEMID (+) " & _
                "AND EE.DWGID = PB.DWGID (+) " & _
                "ORDER BY K.AN8_CUNO, K.KITREF, EU.ELTCODE, EU.ELSUFFIX"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If rst.Fields("AN8_CUNO") <> lCUNO Then
            lCUNO = rst.Fields("AN8_CUNO")
            If sCList = "" Then
                sCList = CStr(lCUNO)
            Else
                sCList = sCList & ", " & CStr(lCUNO)
            End If
        End If
        
        If lCUNO <> CLng(tmpBCC) Then
            sClient = " [" & Trim(rst.Fields("CLIENT")) & "]"
        Else
            sClient = ""
        End If
        
        If sKNode <> "k" & rst.Fields("KITID") Then
            sKNode = "k" & rst.Fields("KITID")
            sDesc = UCase(Trim(rst.Fields("KITFNAME"))) & sClient
            sDescPar = sDesc
            iType = 5
            Set nodX = tvwConst(0).Nodes.Add(, , sKNode, sDesc, iType)
        End If
        If sENode <> "e" & rst.Fields("ELTID") Then
            '///// THIS IS A CHILD \\\\\
            sENode = "e" & rst.Fields("ELTID")
            sDesc = sDescPar & "-" & UCase(Trim(rst.Fields("ELTFNAME"))) & "  " & _
                        UCase(Trim(rst.Fields("ELTDESC")))
            If rst.Fields("DWFTYPE") = 20 And rst.Fields("DWFPATH") <> "" Then
                iType = 2
                lstBlocks(0).AddItem Trim(rst.Fields("DWFPATH"))
                lstBlocks(0).ItemData(lstBlocks(0).NewIndex) = rst.Fields("ELTID")
                If tvwConst(0).Nodes(sKNode).Image <> 1 Then tvwConst(0).Nodes(sKNode).Image = 4
            Else
                iType = 5
            End If
            Set nodX = tvwConst(0).Nodes.Add(sKNode, tvwChild, sENode, sDesc, iType)
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    If Len(sCList) <> 0 Then
'''        strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, SAW.* " & _
'''                    "FROM " & IGLKitU & " KU, " & IGLEltU & " EU, " & IGLKit & " K, " & _
'''                    "" & DWGElt & " DE, " & _
'''                    "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
'''                    "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
'''                    "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'''                    "WHERE DM.AN8_CUNO IN (" & sCList & ") AND DM.DWGTYPE = 10 " & _
'''                    "AND DM.DWGID = DS.DWGID " & _
'''                    "AND DM.DWGID = DD.DWGID " & _
'''                    "AND DS.DWGID = DD.DWGID " & _
'''                    "AND DS.SHTID = DD.SHTID) SAW " & _
'''                    "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " AND KU.SHYR = " & tmpSHYR & " AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'''                    "AND KU.KITUSEID = EU.KITUSEID " & _
'''                    "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
'''                    "AND KU.SHYR = EU.SHYR " & _
'''                    "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
'''                    "AND EU.ESTATUS > 2 AND EU.SHSTATUS <> 3 " & _
'''                    "AND EU.KITID = K.KITID " & _
'''                    "AND EU.ELTID = DE.INVID " & _
'''                    "AND DE.DWGID = SAW.DWGID " & _
'''                    "ORDER BY KU.AN8_CUNO, KU.KITREF, EU.ELTCODE, EU.ELSUFFIX, " & _
'''                    "SAW.DWGNUM , SAW.SHTSEQ"
                    
'        strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, SAW.* " & _
'                    "FROM " & IGLKitU & " KU, " & IGLEltU & " EU, " & IGLKit & " K, " & _
'                    "ENG_ELTID DE, ENG_ELEMENT EE, " & _
'                    "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
'                    "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
'                    "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'                    "WHERE DM.AN8_CUNO IN (" & sCList & ") AND DM.DWGTYPE = 10 " & _
'                    "AND DM.DWGID = DS.DWGID " & _
'                    "AND DM.DWGID = DD.DWGID " & _
'                    "AND DS.DWGID = DD.DWGID " & _
'                    "AND DS.SHTID = DD.SHTID) SAW " & _
'                    "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " AND KU.SHYR = " & tmpSHYR & " AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'                    "AND KU.KITUSEID = EU.KITUSEID " & _
'                    "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
'                    "AND KU.SHYR = EU.SHYR " & _
'                    "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
'                    "AND EU.ESTATUS > 2 AND EU.SHSTATUS <> 3 " & _
'                    "AND EU.KITID = K.KITID " & _
'                    "AND EU.ELTID = DE.ELTID " & _
'                    "AND DE.ELEMID = EE.ELEMID " & _
'                    "AND EE.DWGID = SAW.DWGID " & _
'                    "ORDER BY KU.AN8_CUNO, KU.KITREF, EU.ELTCODE, EU.ELSUFFIX, " & _
'                    "SAW.DWGNUM , SAW.SHTSEQ"
                    
'        strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
'                    "NVL(EE.PRGID, 0)PRGID, SAW.* " & _
'                    "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & IGLKit & " K, " & _
'                    "ENG_ELTID EL, ENG_ELEMENT EE, DWG_ELEMENT DE, " & _
'                    "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
'                    "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
'                    "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'                    "WHERE DM.AN8_CUNO IN (" & sCList & ") AND DM.DWGTYPE = 10 " & _
'                    "AND DM.DWGID = DS.DWGID " & _
'                    "AND DM.DWGID = DD.DWGID " & _
'                    "AND DS.DWGID = DD.DWGID " & _
'                    "AND DS.SHTID = DD.SHTID) SAW " & _
'                    "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " AND KU.SHYR = " & tmpSHYR & " AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'                    "AND KU.KITUSEID = EU.KITUSEID " & _
'                    "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
'                    "AND KU.SHYR = EU.SHYR " & _
'                    "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
'                    "AND EU.ESTATUS > 2 AND EU.SHSTATUS <> 3 " & _
'                    "AND EU.KITID = K.KITID " & _
'                    "AND EU.ELTID = DE.INVID " & _
'                    "AND DE.DWGID = SAW.DWGID " & _
'                    "AND EU.ELTID = EL.ELTID (+) " & _
'                    "AND EL.ELEMID = EE.ELEMID (+) " & _
'                    "ORDER BY KU.AN8_CUNO, KU.KITREF, EU.ELTCODE, EU.ELSUFFIX, " & _
'                    "PRGID, SAW.DWGNUM , SAW.SHTSEQ"
                    
        strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
                    "NVL(EE.PRGID, 0)PRGID, SAW.* " & _
                    "FROM " & AQUAKitU & " KU, " & AQUAEltU & " EU, " & IGLKit & " K, " & _
                    "ANNOTATOR.ENG_ELTID EL, ANNOTATOR.ENG_ELEMENT EE, ANNOTATOR.DWG_ELEMENT DE, " & _
                    "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
                    "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID " & _
                    "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
                    "WHERE DM.DWGTYPE = 10 " & _
                    "AND DM.DWGID = DS.DWGID " & _
                    "AND DM.DWGID = DD.DWGID " & _
                    "AND DS.DWGID = DD.DWGID " & _
                    "AND DS.SHTID = DD.SHTID) SAW " & _
                    "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " AND KU.SHYR = " & tmpSHYR & " AND KU.AN8_SHCD = " & tmpSHCD & " " & _
                    "AND KU.KITUSEID = EU.KITUSEID " & _
                    "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
                    "AND KU.SHYR = EU.SHYR " & _
                    "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
                    "AND EU.ESTATUS > 2 AND EU.SHSTATUS <> 3 " & _
                    "AND EU.KITID = K.KITID " & _
                    "AND EU.ELTID = DE.INVID " & _
                    "AND DE.DWGID = SAW.DWGID " & _
                    "AND EU.ELTID = EL.ELTID (+) " & _
                    "AND EL.ELEMID = EE.ELEMID (+) " & _
                    "ORDER BY KU.AN8_CUNO, KU.KITREF, EU.ELTCODE, EU.ELSUFFIX, " & _
                    "PRGID, SAW.DWGNUM , SAW.SHTSEQ"
                    
                    
'        strSelect = "SELECT K.KITID, K.KITFNAME, EU.ELTID, EU.ELTFNAME, EU.ELTDESC, " & _
'                    "NVL(EE.PRGID, 0)PRGID, SAW.* " & _
'                    "FROM " & IGLKitU & " KU, " & IGLEltU & " EU, " & IGLKit & " K, " & _
'                    "ENG_ELTID EL, ENG_ELEMENT EE, DWG_ELEMENT DE, " & _
'                    "(SELECT DM.DWGID, DM.DWGDESC, DS.SHTDESC, " & _
'                    "DM.DWGYR , DM.MCU, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DD.DWFID "
'
'        strSelect = strSelect & _
'                    "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & DWGDwf & " DD " & _
'                    "WHERE DM.AN8_CUNO IN (" & sCList & ") AND DM.DWGTYPE = 10 " & _
'                    "AND DM.DWGID = DS.DWGID " & _
'                    "AND DS.SSTATUS > 1 " & _
'                    "AND DS.DWGID = DD.DWGID " & _
'                    "AND DS.SHTID = DD.SHTID) SAW " & _
'                    "WHERE KU.AN8_CUNO = " & CLng(tmpBCC) & " AND KU.SHYR = " & tmpSHYR & " AND KU.AN8_SHCD = " & tmpSHCD & " " & _
'                    "AND KU.KITUSEID = EU.KITUSEID " & _
'                    "AND KU.AN8_CUNO = EU.AN8_CUNO " & _
'                    "AND KU.SHYR = EU.SHYR " & _
'                    "AND KU.AN8_SHCD = EU.AN8_SHCD " & _
'                    "AND EU.ESTATUS > 2 AND EU.SHSTATUS <> 3 " & _
'                    "AND EU.KITID = K.KITID " & _
'                    "AND EU.ELTID = DE.INVID " & _
'                    "AND DE.DWGID = SAW.DWGID " & _
'                    "AND EU.ELTID = EL.ELTID (+) " & _
'                    "AND EL.ELEMID = EE.ELEMID (+) " & _
'                    "ORDER BY KU.AN8_CUNO, KU.KITREF, EU.ELTCODE, EU.ELSUFFIX, " & _
'                    "PRGID, SAW.DWGNUM , SAW.SHTSEQ"
        Set rst = Conn.Execute(strSelect)
        iType = 3
        ictr = 1
        Do While Not rst.EOF
            sENode = "e" & rst.Fields("ELTID")
            sDNode = "d" & ictr & "-" & rst.Fields("DWFID")
            If rst.Fields("PRGID") = 0 Then
                Select Case Len(Trim(rst.Fields("MCU")))
                    Case 6
                        sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 2, 4) & "-" & _
                                    Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                    "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                    Case 9
                        sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 4) & "-" & _
                                    Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                    "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                    Case 10
                        sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Mid(Trim(rst.Fields("MCU")), 3, 5) & "-" & _
                                    Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                    "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                    Case Else
                        sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Trim(rst.Fields("MCU")) & "-" & _
                                    Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
                                    "  [" & Trim(rst.Fields("SHTDESC")) & "]"
                End Select
            Else
                sDesc = rst.Fields("PRGID") & "-" & Right("00" & CStr(rst.Fields("DWGNUM")), 2) & _
                            Trim(rst.Fields("SHTSEQ")) & "  [" & Trim(rst.Fields("SHTDESC")) & "]"
            End If
            
'''            sDesc = Right(rst.Fields("DWGYR"), 2) & "-" & Trim(rst.Fields("MCU")) & "-" & _
'''                        Right("00" & CStr(rst.Fields("DWGNUM")), 2) & Trim(rst.Fields("SHTSEQ")) & _
'''                        "  [" & Trim(rst.Fields("SHTDESC")) & "]"
            Set nodX = tvwConst(0).Nodes.Add(sENode, tvwChild, sDNode, sDesc, iType)
            If nodX.Parent.Image <> 2 Then nodX.Parent.Image = 4
            If nodX.Parent.Parent.Image <> 1 Then nodX.Parent.Parent.Image = 4
            ictr = ictr + 1
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
    End If
    Set nodX = Nothing
    tvwConst(0).Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sCheck As String, strUpdate As String
    Dim RetVal As Variant
    Dim i As Integer
    Dim Resp As VbMsgBoxResult

     '///// CHECK IF RED SHOULD BE SAVED \\\\\
    If bSaveRed = True Then
        Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNo, "Redline Changes...")
        If Resp = vbYes Then
            mnuVRedSave_Click
        ElseIf Resp = vbNo Then
            volConst.ClearMarkup
        End If
    End If
    
    If bReded = True And bTeam = True And bPerm(15) Then
        With frmRedAlert
'            .PassSHYR = redSHYR
            .PassBCC = CLng(redBCC)
'            .PassSHCD = redSHCD
            .PassHDR = lblWelcome
            .PassType = 2
            .Show 1
        End With
'''        Call RedAlert(0, lblWelcome, redBCC, redSHCD)
    End If
    bReded = False
    redBCC = ""
    
'    Unload Me
End Sub

Private Sub imgComm_Click()
    With frmComments
        .PassREFID = lRefID
        .PassTable = "ANNOTATOR.DWG_DWF"
        .PassIType = 2
        .PassBCC = tBCC
        .PassFBCN = tFBCN
        .PassSHCD = tSHCD
        .PassMessPath = lblWelcome.Caption
        .PassMessSub = lblDWF.Caption
        .PassForm = "frmConst"
        .Show 1
    End With
End Sub

Private Sub imgMenu_Click()
    If volConst.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
    End If
End Sub

'''Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case bDirsOpen
'''        Case True: Set imgDirs.Picture = imlDirs.ListImages(3).Picture
'''        Case False: Set imgDirs.Picture = imlDirs.ListImages(1).Picture
'''    End Select
'''End Sub

'''Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.ForeColor = lGeo_Back ''vbWhite
'''    lblSettings.ForeColor = lGeo_Back ''vbWhite
'''End Sub

'''Private Sub imgDirs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    Select Case bDirsOpen
'''        Case True: Set imgDirs.Picture = imlDirs.ListImages(4).Picture
'''        Case False: Set imgDirs.Picture = imlDirs.ListImages(2).Picture
'''    End Select
'''End Sub
'''
'''Private Sub imgMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblMenu.ForeColor = lGeo_Back ''vbWhite
'''End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

'''Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.ForeColor = vbWhite '' vbYellow
'''End Sub

Private Sub lblDWF_Click()
    If InStr(1, lblDWF.Caption, "click this link") <> 0 Then
        Call mnuVRedLoad_Click
    End If
End Sub

Private Sub lblMenu_Click()
'''    lblMenu.ForeColor = vbWhite
    If volConst.Visible = True Then
        Me.PopupMenu mnuVolo, 0, imgMenu.Left, imgMenu.Top + imgMenu.Height
    End If
End Sub

'''Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblMenu.ForeColor = vbWhite
'''End Sub

Private Sub lblSettings_Click()
'    frmSettings.PassFrom = "FP"
'    frmSettings.PassBCC = CLng(BCC)
'    frmSettings.PassFBCN = FBCN
'    frmSettings.PassBCC_DEF = defCUNO ''' lBCC_Def
'    frmSettings.PassFBCN_DEF = defFBCN ''' sFBCN_Def
'    frmSettings.Show 1
End Sub

Private Sub lblSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSettings.ForeColor = vbWhite '' vbYellow
End Sub

Private Sub mnuCopyClip_Click()
    SendKeys "^C"
End Sub

Private Sub mnuDownloadPDF_Click()
'    MsgBox "Coming soon", vbExclamation, "Sorry..."
    frmBrowse.PassFBCN = cboCUNO(iTab).Text
    frmBrowse.PassFILETYPE = "pdf"
    frmBrowse.PassFrom = UCase(Me.Name)
    frmBrowse.PassDWGID = tDWGID(iTab)
    frmBrowse.Show 1, Me
End Sub

Private Sub mnuDownloadZip_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, _
                sChk As String, sPath As String, sFile As String, _
                strInsert As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lID As Long
    
    
    If shlShell Is Nothing Then
        Set shlShell = New Shell32.Shell
    End If
    
    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
                "Select Folder to download DWG zip file into:", _
                BIF_RETURNONLYFSDIRS)
    
    If shlFolder Is Nothing Then
        Exit Sub
    Else
        Screen.MousePointer = 11
        
        On Error GoTo BadFile
        sFolder = shlFolder.Items.Item.Path
        
        If UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
'''        ''STOPPED HERE''
'''        MsgBox "Stopped.  Under Construction..."
'''        Exit Sub
        
        On Error GoTo ErrorTrap
        strSelect = "SELECT DM.DWGDESC, DD.DWFPATH " & _
                    "FROM " & DWGMas & " DM, " & DWGDwf & " DD " & _
                    "WHERE DD.DWFID = " & lRefID & " " & _
                    "AND DD.DWGID = DM.DWGID"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sFile = Trim(rst.Fields("DWFPATH"))
            sFile = Left(sFile, Len(sFile) - 4) & ".zip"
            sPath = sFolder & "\" & Trim(rst.Fields("DWGDESC")) & ".zip"
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
        
        ''WRITE DOWNLOAD TO ANO_LOCKLOG''
        Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
        lID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        
        strInsert = "INSERT INTO ANNOTATOR.ANO_LOCKLOG " & _
                    "(LOCKID, LOCKREFID, LOCKREFSOURCE, " & _
                    "USER_SEQ_ID, LOCKOPENDTTM, LOCKSTATUS, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "Values " & _
                    "(" & lID & ", " & lRefID & ", 'DWG_DOWNLOAD', " & _
                    UserID & ", SYSDATE, 99, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
        
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

Private Sub mnuDWFData_Click()
    Dim sMess As String, strSelect As String
    Dim rst As ADODB.Recordset
    
'''    MsgBox "Edit node - " & tvwConst(iTab).SelectedItem.key & _
'''                "  [" & tvwConst(iTab).SelectedItem.Text & "]"
    sMess = ""
    strSelect = "SELECT ADDUSER, ADDDTTM, UPDUSER, UPDDTTM " & _
                "FROM ANNOTATOR.DWG_DWF WHERE DWFID = " & Mid(tvwConst(iTab).SelectedItem.Key, 2)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If rst.Fields("ADDDTTM") <> rst.Fields("UPDDTTM") Then
            sMess = "Drawing originally posted by " & _
                        Trim(rst.Fields("ADDUSER")) & " on " & _
                        Format(rst.Fields("ADDDTTM"), "DDDD, MMMM D, YYYY") & "." & _
                        vbNewLine & _
                        "File last updated by " & Trim(rst.Fields("UPDUSER")) & " on " & _
                        Format(rst.Fields("UPDDTTM"), "DDDD, MMMM D, YYYY") & "."
        Else
            sMess = "Drawing originally posted by " & _
                        Trim(rst.Fields("ADDUSER")) & " on " & _
                        Format(rst.Fields("ADDDTTM"), "DDDD, MMMM D, YYYY") & "."
        End If
    End If
    rst.Close: Set rst = Nothing
    
    If sMess <> "" Then MsgBox sMess, vbInformation, tvwConst(iTab).SelectedItem.Text
End Sub

Private Sub mnuDWFDelete_Click()
        
    MsgBox "Delete node - " & tvwConst(iTab).SelectedItem.Key & _
                "  [" & tvwConst(iTab).SelectedItem.Text & "]" & _
                vbNewLine & vbNewLine & _
                "UserType = " & UserType & vbNewLine & _
                "Group = " & tvwConst(iTab).SelectedItem.Parent.Parent.Text
    
    Call DeleteDWF(CLng(Mid(tvwConst(iTab).SelectedItem.Key, 2)))
End Sub

Private Sub mnuFindElems_Click()
    Dim sClient As String, sMess As String
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    
    sClient = "": sMess = ""
    strSelect = "SELECT AB.ABALPH, K.KITFNAME, E.ELTFNAME, E.ELTDESC " & _
                "FROM IGLPROD.IGL_KIT K, IGLPROD.IGL_ELEMENT E, " & F0101 & " AB " & _
                "WHERE E.ELTID IN " & _
                "(SELECT INVID FROM ANNOTATOR.DWG_ELEMENT " & _
                "WHERE DWGID = " & tDWGID(iTab) & ") " & _
                "AND E.KITID = K.KITID " & _
                "AND K.AN8_CUNO = AB.ABAN8 " & _
                "ORDER BY UPPER(AB.ABALPH), K.KITREF, E.ELTFNAME"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If UCase(Trim(rst.Fields("ABALPH"))) <> sClient Then
            sClient = UCase(Trim(rst.Fields("ABALPH")))
            sMess = sMess & vbNewLine & sClient & vbNewLine
        End If
        sMess = sMess & Space(8) & UCase(Trim(rst.Fields("KITFNAME"))) & " - " & _
                    UCase(Trim(rst.Fields("ELTFNAME"))) & "  " & UCase(Trim(rst.Fields("ELTDESC"))) & _
                    vbNewLine
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
'''    Screen.MousePointer = 0
    If sMess <> "" Then
        sMess = "The selected Drawing has been attached to the following Elements:" & _
                    vbNewLine & sMess
        MsgBox sMess, vbInformation, "Inventory Information..."
    Else
        MsgBox "The selected Drawing has not been attached to any Elements in the Inventory.", _
                    vbInformation, "Inventory Information..."
    End If

End Sub

Private Sub mnuFPFullView_Click()
    volFP.SetCurrentView dLeftFP, dRightFP, dBottomFP, dTopFP
End Sub

Private Sub mnuFPPan_Click()
    ClearFPChecks
    mnuFPPan.Checked = True
    volFP.UserMode = "Pan"
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

Private Sub mnuHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub mnuPhoto_Click()
    frmPhoto.PassEID = tELTID
    frmPhoto.PassIn = "1"
    frmPhoto.PassLink = tHDR
    frmPhoto.Show 1, Me
    
End Sub

Private Sub mnuRCFullView_Click()
    
    volView(iD).SetCurrentView dLeft2(iD), dRight2(iD), dBottom2(iD), dTop2(iD)
End Sub

Private Sub mnuRCPan_Click()
    Dim i As Integer
    mnuRCPan.Checked = True
    mnuRCZoom.Checked = False
    For i = 0 To 3
        volView(i).UserMode = "Pan"
    Next i
End Sub

Private Sub mnuRCZoom_Click()
    Dim i As Integer
    mnuRCPan.Checked = False
    mnuRCZoomW.Checked = False
    mnuRCZoom.Checked = True
    For i = 0 To 3
        volView(i).UserMode = "Zoom"
    Next i
End Sub

Private Sub mnuRCZoomW_Click()
    Dim i As Integer
    mnuRCPan.Checked = False
    mnuRCZoom.Checked = False
    mnuRCZoomW.Checked = True
    For i = 0 To 3
        volView(i).UserMode = "ZoomToRect"
    Next i
End Sub

Private Sub mnuSortNodes_Click(Index As Integer)
    Dim sOrderBy As String, strSelect As String, sLoc As String, sMCU As String
    Dim rst As ADODB.Recordset
    Dim nodPar As Node, nodX As Node
    Dim sDNode As String, sDesc As String
    Dim iType As Integer
    
    
    Set nodPar = tvwConst(iTab).SelectedItem
    
    ''SET SLOC''
    sLoc = UCase(Mid(nodPar.Parent.Key, 2))
    
    ''SET SMCU''
    sMCU = Right(Space(12) & Mid(nodPar.Key, 6), 12)
    
    Select Case Index
        Case 0: sOrderBy = "ORDER BY DM.DWGDESC"
        Case 1: sOrderBy = "ORDER BY DD.UPDDTTM DESC"
        Case 2: sOrderBy = "ORDER BY DD.UPDDTTM"
    End Select
    
    
    Do While nodPar.Children <> 0
        tvwConst(iTab).Nodes.Remove (nodPar.Child.Key)
    Loop
    
    ''REPOP CHILDREN BASED ON NEW ORDER BY''
    strSelect = "SELECT DM.DWGID, DD.UPDDTTM, DM.DWGDESC, " & _
                "DD.DWFID, DD.DWFPATH, MC.MCDL01 AS PROJ " & _
                "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & _
                DWGDwf & " DD, " & F0006 & " MC " & _
                "WHERE DM.AN8_CUNO = " & CLng(tBCC) & " " & _
                "AND DM.DWGTYPE = 15 " & _
                "AND DM.MCU ='" & sMCU & "' " & _
                "AND DM.DLOC = '" & sLoc & "' " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DD.DWGID " & _
                "AND DS.SHTID = DD.SHTID " & _
                "AND DD.DWFTYPE = -1 " & _
                "AND DM.MCU LIKE MC.MCMCU " & _
                sOrderBy
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        sDNode = "d" & rst.Fields("DWFID")
        sDesc = UCase(Trim(rst.Fields("DWGDESC"))) & " ---- <" & _
                    UCase(Format(rst.Fields("UPDDTTM"), "DD-MMM-YYYY")) & ">"
        If rst.Fields("UPDDTTM") + 4 > Now Then iType = 6 Else iType = 3
        Set nodX = tvwConst(3).Nodes.Add(nodPar.Key, tvwChild, sDNode, sDesc, iType)
        If iType = 6 Or iType = 8 Then ''DO PARENTS''
            nodX.Parent.Image = 7
            nodX.Parent.Parent.Image = 7
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    
    
End Sub

Private Sub mnuVDisplay_Click(Index As Integer)
    Dim i As Integer
    
    iDisplay = Index
    
    For i = 0 To 2
        If i = Index Then mnuVDisplay(i).Checked = True Else mnuVDisplay(i).Checked = False
    Next i
    Select Case Index
    Case 0
        volConst.GeometryColor = "DefaultColors"
        volConst.BackgroundColor = "DefaultColors"
    Case 1
        volConst.GeometryColor = vbBlack
        volConst.BackgroundColor = vbWhite
    Case 2
        volConst.GeometryColor = "ClearScale"
        volConst.BackgroundColor = "ClearScale"
    End Select
End Sub

Private Sub mnuVElemInfo_Click()
    With frmElemInfo
        .PassELTID = tELTID
        .PassKITID = tKitID
        .PassHDR = tHDR
        .Show 1
    End With
    tELTID = 0: tKitID = 0: tHDR = ""
End Sub

Private Sub mnuVFullView_Click()
    volConst.SetCurrentView dLeft, dRight, dBottom, dTop
End Sub

Private Sub mnuVLayers_Click()
    volConst.ShowLayersDialog
End Sub

Private Sub mnuVNamedViews_Click()
    volConst.ShowNamedViewsDialog
End Sub

Private Sub mnuVPan_Click()
    ClearChecks
    mnuVPan.Checked = True
    volConst.UserMode = "Pan"
End Sub

Private Sub mnuVPrint_Click()
    volConst.ShowPrintDialog
End Sub

Private Sub mnuVRedClear_Click()
    Dim Resp As VbMsgBoxResult
    
    If bSaveRed Then
        Resp = MsgBox("The annotation has been edited.  Are you certain you want " & _
                    "to clear the Redlines?", vbYesNo + vbExclamation, "Hey...")
        If Resp = vbYes Then
            volConst.ClearMarkup
        End If
    Else
        volConst.ClearMarkup
    End If
End Sub

Private Sub mnuVRedDelete_Click()
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
            volConst.src = volView(iTab).src

            
            '///// DELETE REDLINE FROM DATABASE \\\\\
            strDelete = "DELETE FROM " & DWGDwf & " " & _
                        "WHERE DWFID = " & lRedID
            Conn.Execute (strDelete)
                
            
            '///// DELETE ACTUAL DWF FILE \\\\\
            Kill RedFile
            
            '///// NOW, CLEAN UP \\\\\
            lRedID = 0
            RedFile = ""
            lblDWF.ForeColor = vbWhite
            lblDWF.Caption = "NO Redline File exists for this Floor Plan."
            mnuVRedLoad.Enabled = False
            mnuVRedDelete.Enabled = False
            mnuVRedReturn.Enabled = False
            If volConst.UserMode = "Sketch" Or volConst.UserMode = "Text" Then
                mnuVRedClear.Enabled = True
                mnuVRedSave.Enabled = True
            Else
                mnuVRedClear.Enabled = False
                mnuVRedSave.Enabled = False
            End If
        End If
    Else
        MsgBox "You do not have permission to delete this file." & vbNewLine & _
                "To delete Redline Files, you must be a member" & vbNewLine & _
                "of the Email Notification Team for this Client.", vbCritical, "Sorry..."
    End If
End Sub

Private Sub mnuVRedEnd_Click()
    ClearChecks
    mnuVRedEnd.Enabled = False
    Select Case sPrevMode
        Case "Zoom"
            mnuVZoom.Checked = True
            volConst.UserMode = "Zoom"
        Case "ZoomToRect"
            mnuVZoomW.Checked = True
            volConst.UserMode = "ZoomToRect"
        Case "Pan"
            mnuVPan.Checked = True
            volConst.UserMode = "Pan"
    End Select
End Sub

Private Sub mnuVRedLoad_Click()
    Dim Resp As VbMsgBoxResult
    
    Screen.MousePointer = 11
    
    bViewSet = False
    volConst.src = RedFile
    volConst.Update
    bViewSet = False
    volConst.src = RedFile
    lblDWF.Caption = "Redline File Loaded"
    If volConst.UserMode = "sketch" Or volConst.UserMode = "text" Then
        mnuVRedSave.Enabled = True
    Else
        mnuVRedSave.Enabled = False
    End If
    mnuVRedClear.Enabled = True
    mnuVRedDelete.Enabled = True
    mnuVRedReturn.Enabled = True
    
'''    mnuDownloadPDF.Enabled = False
'''    mnuEmailPDF.Enabled = False
    
    Screen.MousePointer = 0
    
    
    
'''    If RedFile = "" Or volConst.src = RedFile Then
'''        ClearChecks
'''        mnuVSketch.Checked = True
'''        mnuVText.Checked = False
'''        mnuVRedEnd.Enabled = True
'''        volConst.UserMode = "Sketch"
'''        mnuVRedSave.Enabled = True
'''        mnuVRedClear.Enabled = True
'''    ElseIf volConst.src <> RedFile Then
'''        Resp = MsgBox("A Redline File already exists, but is not currently loaded." & _
'''                    vbCr & "Select 'YES' to begin a New Redline, Select 'NO' to Abort." & _
'''                    vbCr & vbCr & "NOTE: Original Redline will not be overwritten until 'Saved'.", _
'''                    vbYesNo + vbCritical + vbDefaultButton2, "Existing Redline File...")
'''        If Resp = vbYes Then
'''            ClearChecks
'''            mnuVSketch.Checked = True
'''            mnuVText.Checked = False
'''            volConst.UserMode = "Sketch"
'''            mnuVRedSave.Enabled = True
'''            mnuVRedClear.Enabled = True
'''        End If
'''    End If
End Sub

Private Sub mnuVRedReturn_Click()
    Screen.MousePointer = 11
    cmdLoadDWF_Click (iTab)
    Screen.MousePointer = 0
End Sub

Private Sub mnuVRedSave_Click()
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
        lRedID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        RedFile = sDWFPath & CStr(lRedID) & ".dwf"
        strInsert = "INSERT INTO " & DWGDwf & " " & _
                    "(DWGID, SHTID, DWFID, DWFTYPE, " & _
                    "DWFDESC, DWFPATH, DWFSTATUS, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lDWGID & ", " & lSHTID & ", " & lRedID & ", -9, " & _
                    "'REDLINE', '" & RedFile & "', 1, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
        
'''        lRedID = lDWFID
    Else
        strUpdate = "UPDATE " & DWGDwf & " " & _
                    "SET UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE DWGID = " & lDWGID & " " & _
                    "AND SHTID = " & lSHTID & " " & _
                    "AND DWFID = " & lRedID
        Conn.Execute (strUpdate)
        
    End If
    volConst.SaveMarkup (RedFile)
    If Err Then
        Conn.RollbackTrans
        MsgBox "Redline Annotation cannot be saved." & vbNewLine & vbNewLine & _
                    "Error:  " & Err.Description, vbExclamation, "Error Encountered..."
    Else
        Conn.CommitTrans
        lblDWF.ForeColor = vbRed
        lblDWF.Caption = "A Redline File exists for this drawing.  " & _
                    "To Load it, use the Viewer Menu, or click this link."
        bReded = True
'''        If Not bLoading Then
            redBCC = tBCC '': redSHCD = SHCD: redSHYR = SHYR
'''        End If
        bSaveRed = False
        lblDWF.Visible = True
        mnuVRedLoad.Enabled = True
    End If

End Sub

Private Sub mnuVSketch_Click()
    Dim Resp As VbMsgBoxResult
'    ClearChecks
'    mnuVSketch.Checked = True
    If volConst.UserMode <> "Sketch" Or volConst.UserMode = "Text" Then sPrevMode = volConst.UserMode
'    volConst.UserMode = "Sketch"
'    mnuVRedSave.Enabled = True
'    mnuVRedClear.Enabled = True
    
    If RedFile = "" Or volConst.src = RedFile Then
        ClearChecks
        mnuVSketch.Checked = True
        mnuVText.Checked = False
        mnuVRedEnd.Enabled = True
        volConst.UserMode = "Sketch"
        mnuVRedSave.Enabled = True
        mnuVRedClear.Enabled = True
    ElseIf volConst.src <> RedFile Then
        Resp = MsgBox("A Redline File already exists, but is not currently loaded." & _
                    vbCr & "Select 'YES' to begin a New Redline, Select 'NO' to Abort." & _
                    vbCr & vbCr & "NOTE: Original Redline will not be overwritten until 'Saved'.", _
                    vbYesNo + vbCritical + vbDefaultButton2, "Existing Redline File...")
        If Resp = vbYes Then
            ClearChecks
            mnuVSketch.Checked = True
            mnuVText.Checked = False
            volConst.UserMode = "Sketch"
            mnuVRedSave.Enabled = True
            mnuVRedClear.Enabled = True
        End If
    End If
End Sub

Private Sub mnuVText_Click()
    ClearChecks
    mnuVText.Checked = True
    If volConst.UserMode <> "Sketch" Or volConst.UserMode = "Text" Then sPrevMode = volConst.UserMode
    volConst.UserMode = "Text"
    mnuVRedSave.Enabled = True
    mnuVRedClear.Enabled = True
End Sub

Private Sub mnuVZoom_Click()
    ClearChecks
    mnuVZoom.Checked = True
    volConst.UserMode = "Zoom"
End Sub

Private Sub mnuVZoomW_Click()
    ClearChecks
    mnuVZoomW.Checked = True
    volConst.UserMode = "ZoomToRect"
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    iTab = sst1.Tab
    If iTab = 3 Then
        If tvwConst(iTab).Nodes.Count = 0 And cboCUNO(iTab).ListCount = 1 Then
            cboCUNO(iTab).Text = cboCUNO(iTab).List(0)
        End If
    End If
End Sub

Private Sub tvwConst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tvwConst(Index).Nodes.Count = 0 Then Exit Sub
    If UCase(Left(tvwConst(Index).SelectedItem.Key, 1)) = "E" _
                And Button = vbRightButton Then
        tHDR = tvwConst(Index).SelectedItem.Text
        tELTID = Mid(tvwConst(Index).SelectedItem.Key, 2)
        tKitID = Mid(tvwConst(Index).SelectedItem.Parent.Key, 2)
        Me.PopupMenu mnuElemInfo
    End If
    If Button = vbRightButton Then bRightButton = True Else bRightButton = False
End Sub

Private Sub tvwConst_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    Dim strSelect As String, sClient As String, sMess As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iDash As Integer
    
    
    Screen.MousePointer = 11
    If Index < 2 Then cmdElemInfo(Index).Visible = False
    Debug.Print Node.Text & "   (Image=" & Node.Image & ")"
    Debug.Print Node.Key
    Select Case Index
        Case 0, 1
            If UCase(Left(Node.Key, 1)) = "K" And Node.Image = 1 Then
                bViewSet = False
                sNewWelcome = cboCUNO(Index).Text & " Kit Sheet"
                sNewName = "Kit:  " & Node.Text
            ElseIf UCase(Left(Node.Key, 1)) = "E" Then
                If Node.Image = 2 Then
                    volView(Index).Visible = False
                    sNewWelcome = cboCUNO(Index).Text & " Property Drawing"
                    sNewName = "Element:  " & Node.Text
                    sHDR(Index) = sNewName
                    lblEltID(Index) = 0: lblKitID(Index) = 0
                    cmdElemInfo(Index).Visible = False
                    For i = 0 To lstBlocks(Index).ListCount - 1
                        If lstBlocks(Index).ItemData(i) = Mid(Node.Key, 2) Then
                            volView(Index).src = lstBlocks(Index).List(i)
                            tRefID = 0: tRedID = 0: RedName = ""
                            lblEltID(Index) = Mid(Node.Key, 2)
                            lblKitID(Index) = Mid(Node.Parent.Key, 2)
                            cmdElemInfo(Index).Visible = True
                            GoTo FoundBlock
                        End If
                    Next i
FoundBlock:
                    lblName(Index) = Node.Text
                    volView(Index).Visible = True
                    cmdLoadDWF(Index).Visible = True
'''                    sst1.Width = 11520 '''10560 '''9360
                Else
                    Screen.MousePointer = 0
                    tHDR = Node.Text
                    tELTID = Mid(Node.Key, 2)
                    tKitID = Mid(Node.Parent.Key, 2)
                    Me.PopupMenu mnuElemInfo
                End If
            ElseIf UCase(Left(Node.Key, 1)) = "D" Then
''                '///// CHECK IF RED SHOULD BE SAVED \\\\\
''                If SaveRed = True Then
''                    Resp = MsgBox("Do you wish to Save the Redline Changes?", vbYesNoCancel, "Redline Changes...")
''                    If Resp = vbYes Then
''                        mnuVRedSave_Click
''                    ElseIf Resp = vbNo Then
''                        volConst.ClearMarkup
''            '''            SaveRed = False
''                    ElseIf Resp = vbCancel Then
''                        Exit Sub
''                    End If
''                End If
''
''                If bReded = True And bTeam = True And bPerm(15) Then
''                    With frmRedAlert
'''''                        .PassSHYR = redSHYR
''                        .PassBCC = CLng(redBCC)
'''''                        .PassSHCD = redSHCD
''                        .PassHDR = lblWelcome.Caption
''                        .PassType = 2 ''ENGINEERING''
''                        .Show 1
''                    End With
''            '''        Call RedAlert(0, lblWelcome, redBCC, redSHCD) 'AlertOfRed
''                End If
''                bReded = False
''                redBCC = "" '': redSHCD = 0: redSHYR = 0
                
                
                
                iDash = InStr(1, Node.Key, "-")
                tDWFID(Index) = CLng(Mid(Node.Key, iDash + 1))
                strSelect = "SELECT DWGID, SHTID, DWFPATH " & _
                            "FROM " & DWGDwf & " " & _
                            "WHERE DWFID = " & tDWFID(Index)
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    bViewSet = False
                    tDWGID(Index) = rst.Fields("DWGID"): tSHTID(Index) = rst.Fields("SHTID")
                    tRefID = tDWFID(Index): tRedID = tDWFID(Index) '': RedName = CStr(tDWFID) & "RED.bmp"
                    volView(Index).src = Trim(rst.Fields("DWFPATH"))
                    volView(Index).Visible = True
                    cmdLoadDWF(Index).Visible = True
                    lblName(Index) = Node.Text
                    sNewWelcome = cboCUNO(Index).Text & " - " & Node.Parent.Text
                    sNewName = "Drawing:  " & Node.Text
'''                    sst1.Width = 11520 '''10560 '''9360
                Else
                    volView(Index).src = ""
                    volView(Index).Visible = False
                    cmdLoadDWF(Index).Visible = False
'''                    sst1.Width = 7000 '''6120 '''4800
                End If
                rst.Close
                Set rst = Nothing
            End If
        Case 2, 3
            If bRightButton And UCase(Left(Node.Key, 1)) = "D" Then
                Screen.MousePointer = 0
                If bPerm(70) Then
                    Select Case UCase(UserType)
                        Case "VENDOR - BRC"
                            If tvwConst(Index).SelectedItem.Parent.Parent.Key = "lBRC" Then
                                Me.mnuDWFDelete.Enabled = True
                            Else
                                mnuDWFDelete.Enabled = False
                            End If
                        Case "VENDOR - HOLLOMON"
                            If tvwConst(Index).SelectedItem.Parent.Parent.Key = "lHOL" Then
                                Me.mnuDWFDelete.Enabled = True
                            Else
                                mnuDWFDelete.Enabled = False
                            End If
                        Case "VENDOR - TANSEISHA"
                            If tvwConst(Index).SelectedItem.Parent.Parent.Key = "lTAN" Then
                                Me.mnuDWFDelete.Enabled = True
                            Else
                                mnuDWFDelete.Enabled = False
                            End If
                        Case "CLIENT - AICHI"
                            mnuDWFDelete.Enabled = False
                        Case Else
                            If bPerm(36) Then
                                Me.mnuDWFDelete.Enabled = True
                            Else
                                mnuDWFDelete.Enabled = False
                            End If
                    End Select
                Else
                    If bPerm(36) Then
                        Me.mnuDWFDelete.Enabled = True
                    Else
                        mnuDWFDelete.Enabled = False
                    End If
                End If
                
                Me.PopupMenu mnuDWFEdit
            
            ElseIf bRightButton And UCase(Left(Node.Key, 1)) = "P" Then
                Screen.MousePointer = 0
                Me.PopupMenu mnuSort
                
            ElseIf UCase(Left(Node.Key, 1)) = "D" Then
                tDWFID(Index) = CLng(Mid(Node.Key, 2))
                strSelect = "SELECT DWGID, SHTID, DWFPATH " & _
                            "FROM " & DWGDwf & " " & _
                            "WHERE DWFID = " & tDWFID(Index)
                Set rst = Conn.Execute(strSelect)
                If Not rst.EOF Then
                    bViewSet = False
                    tDWGID(Index) = rst.Fields("DWGID"): tSHTID(Index) = rst.Fields("SHTID")
                    tRefID = tDWFID(Index): tRedID = tDWFID(Index)
                    On Error Resume Next
                    Err = 0
                    volView(Index).src = Trim(rst.Fields("DWFPATH"))
                    If Err Then
'                        MsgBox "Error:  " & Err.Description, vbExclamation, "Error Encountered..."
                        volView(Index).Visible = False
                        lblName(Index) = ""
                        sNewWelcome = ""
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                    
                    volView(Index).Visible = True
                    cmdLoadDWF(Index).Visible = True
                    lblName(Index) = Node.Text
                    If Index = 2 Then
                        sNewWelcome = cboCUNO(Index).Text & "  -  " & Node.Parent.Parent.Parent.Text & _
                                    " - " & Node.Parent.Text & " (" & Node.Parent.Parent.Text & ")"
                    ElseIf Index = 3 Then
                        sNewWelcome = ""
                        On Error Resume Next
                        sNewWelcome = Node.Parent.Parent.Parent.Text & " - "
                        sNewWelcome = cboCUNO(Index).Text & "  -  " & sNewWelcome & Node.Parent.Parent.Text & _
                                    " (" & Node.Parent.Text & ")"
                    End If
                    sNewName = "Drawing:  " & Node.Text
'''                    sst1.Width = 11520 '''10560 '''9360
                Else
                    volView(Index).src = ""
                    volView(Index).Visible = False
                    cmdLoadDWF(Index).Visible = False
'''                    sst1.Width = 7000 '''6120 '''4800
                End If
                rst.Close
                Set rst = Nothing
            ElseIf UCase(Left(Node.Key, 1)) = "S" Then
                tDWGID(Index) = CLng(Mid(Node.Key, 2))
                Screen.MousePointer = 0
                Me.PopupMenu mnuFindElem
                
'''                sClient = "": sMess = ""
'''                strSelect = "SELECT AB.ABALPH, K.KITFNAME, E.ELTFNAME, E.ELTDESC " & _
'''                            "FROM IGL_KIT K, IGL_ELEMENT E, " & F0101 & " AB " & _
'''                            "WHERE E.ELTID IN " & _
'''                            "(SELECT INVID FROM DWG_ELEMENT " & _
'''                            "WHERE DWGID = " & tDWGID & ") " & _
'''                            "AND E.KITID = K.KITID " & _
'''                            "AND K.AN8_CUNO = AB.ABAN8 " & _
'''                            "ORDER BY UPPER(AB.ABALPH), K.KITREF, E.ELTFNAME"
'''                Set rst = Conn.Execute(strSelect)
'''                Do While Not rst.EOF
'''                    If UCase(Trim(rst.FIELDS("ABALPH"))) <> sClient Then
'''                        sClient = UCase(Trim(rst.FIELDS("ABALPH")))
'''                        sMess = sMess & vbNewLine & sClient & vbNewLine
'''                    End If
'''                    sMess = sMess & Space(8) & UCase(Trim(rst.FIELDS("KITFNAME"))) & " - " & _
'''                                UCase(Trim(rst.FIELDS("ELTFNAME"))) & "  " & UCase(Trim(rst.FIELDS("ELTDESC"))) & _
'''                                vbNewLine
'''                    rst.MoveNext
'''                Loop
'''                rst.Close: Set rst = Nothing
'''                Screen.MousePointer = 0
'''                If sMess <> "" Then
'''                    sMess = "The selected Drawing has been attached to the following Elements:" & _
'''                                vbNewLine & sMess
'''                    MsgBox sMess, vbInformation, "Inventory Information..."
'''                Else
'''                    MsgBox "The selected Drawing has not been attached to any Elements in the Inventory.", _
'''                                vbInformation, "Inventory Information..."
'''                End If
            End If
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub txtCPRJ_Change()
    ChangeCase txtCPRJ
    If Len(Trim(txtCPRJ.Text)) > 5 Then
        cmdGo.Enabled = True
        cmdGo.Default = True
    Else
        cmdGo.Enabled = False
        cmdGo.Default = False
    End If
End Sub

Private Sub volConst_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuVolo
    Else
        If volConst.UserMode = "Sketch" Or volConst.UserMode = "Text" Then
            bSaveRed = True
        End If
    End If
End Sub

Private Sub volConst_OnClearMarkup(enable_default As Boolean)
    bSaveRed = False
End Sub

Private Sub volConst_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            InitialView
            bViewSet = True
        End If
    End If
End Sub

Private Sub volConst_OnSaveMarkup(enable_default As Boolean)
    enable_default = False
    mnuVRedSave_Click
End Sub

Private Sub volFP_DoNavigateToURL(ByVal URL As String, ByVal window_name As String, enable_default As Boolean)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim pCUNO As Long
    Dim pElt As String, sFind As String, sKey As String
    Dim i As Integer
    Dim bDWFFound As Boolean
    Dim tELTID As Long
    
    
    enable_default = False
    sLinkID = URL
    pCUNO = CLng(Left(sLinkID, 8))
    pElt = Mid(sLinkID, 10)
    bDWFFound = False
    
    strSelect = "SELECT E.ELTID " & _
                "FROM IGLPROD.IGL_KIT K, IGLPROD.IGL_ELEMENT E " & _
                "Where K.AN8_CUNO = " & pCUNO & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTFNAME = '" & pElt & "' " & _
                "AND E.ESTATUS > 2"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        tELTID = rst.Fields("ELTID")
    Else
        tELTID = 0
    End If
    rst.Close
    
'    strSelect = "SELECT F.DWFID " & _
'                "From DWG_SHEET S, DWG_DWF F " & _
'                "WHERE S.DWGID IN (" & _
'                    "SELECT DWGID " & _
'                    "From DWG_ELEMENT " & _
'                    "WHERE INVID IN (" & _
'                        "SELECT E.ELTID " & _
'                        "FROM IGL_KIT K, IGL_ELEMENT E " & _
'                        "Where K.AN8_CUNO = " & pCUNO & " " & _
'                        "AND K.KITID = E.KITID " & _
'                        "AND E.ELTFNAME = '" & pElt & "'" & _
'                    ")" & _
'                ") " & _
'                "AND S.DWGID = F.DWGID " & _
'                "AND S.SHTID = F.SHTID " & _
'                "ORDER BY S.SHTSEQ"
                
    strSelect = "SELECT S.SHTSEQ, F.DWFID " & _
                "From ANNOTATOR.DWG_MASTER D, ANNOTATOR.DWG_SHEET S, ANNOTATOR.DWG_DWF F " & _
                "WHERE D.DWGID IN (" & _
                    "SELECT DWGID " & _
                    "From ANNOTATOR.DWG_ELEMENT " & _
                    "WHERE INVID = " & tELTID & _
                ") " & _
                "AND D.DWGTYPE = 10 " & _
                "AND D.DWGID = S.DWGID " & _
                "AND S.DWGID = F.DWGID " & _
                "AND S.SHTID = F.SHTID " & _
                "ORDER BY S.SHTSEQ"
                
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sFind = "-" & rst.Fields("DWFID")
        For i = 1 To tvwConst(0).Nodes.Count
            sKey = tvwConst(0).Nodes(i).Key
            If InStr(1, sKey, sFind) > 0 Then
                If UCase(tvwConst(0).Nodes(sKey).Parent.Key) = "E" & CStr(tELTID) Then
                    tvwConst(0).Nodes(sKey).Selected = True
                    tvwConst(0).Nodes(sKey).Expanded = True
                    Call tvwConst_NodeClick(0, tvwConst(0).Nodes(sKey))
                    bDWFFound = True
                    Exit For
                End If
            End If
        Next i
    End If
    rst.Close: Set rst = Nothing
    
    If Not bDWFFound Then MsgBox "No drawing files found for '" & pElt & "'", vbExclamation, "Sorry..."
    
End Sub

Private Sub volFP_MouseDown(Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        Me.PopupMenu mnuFPVolo
    End If
End Sub

Private Sub volFP_OnProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            Call InitialViewFP
            bViewSet = True
        End If
    End If
End Sub

Public Function InitialViewFP()
    volFP.GetCurrentView dLeftFP, dRightFP, dBottomFP, dTopFP
End Function

Private Sub volView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Button = vbRightButton Then
        iD = Index
        Me.PopupMenu mnuRightClick
    End If
End Sub

Public Sub ClearChecks()
    mnuVPan.Checked = False
    mnuVZoom.Checked = False
    mnuVZoomW.Checked = False
    mnuVSketch.Checked = False
    mnuVText.Checked = False
End Sub

Public Sub ClearFPChecks()
    mnuFPPan.Checked = False
    mnuFPZoom.Checked = False
    mnuFPZoomW.Checked = False
End Sub

Public Function InitialView()
    volConst.GetCurrentView dLeft, dRight, dBottom, dTop
End Function

Public Function InitialView2(Index As Integer)
    volView(Index).GetCurrentView dLeft2(Index), dRight2(Index), dBottom2(Index), dTop2(Index)
End Function

Public Function GetBCN(tmpBCC As String)
    Dim rstCN As ADODB.Recordset
    Dim strSelect As String
    
    '****'Conn.Open ALREADY****
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
                "AND SHY56SHYR = " & tmpSHYR
    Set rstSN = Conn.Execute(strSelect)
    If Not rstSN.EOF Then
        GetSHNM = UCase(Trim(rstSN.Fields("SHY56NAMA")))
    Else
        GetSHNM = ""
    End If
    rstSN.Close
    Set rstSN = Nothing
End Function


Private Sub volView_OnProgress(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
    If bViewSet = False Then
        If StatusCode = 42 Then
            Call InitialView2(Index)
            bViewSet = True
        End If
    End If
End Sub

Public Function ChangeCase(Con As Control)
    Dim Pos As Integer
    Pos = Con.SelStart
    Con.Text = UCase(Con.Text)
    Con.SelStart = Pos
End Function

Public Sub PopClientsWithNonGPJDwgs(combo As ComboBox)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    combo.Clear
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT M.AN8_CUNO, C.ABALPH " & _
                    "FROM " & DWGMas & " M, " & F0101 & " C " & _
                    "WHERE M.DWGTYPE = 15 " & _
                    "AND M.DSTATUS > 0 " & _
                    "AND M.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY UPPER(ABALPH)"
    Else
        strSelect = "SELECT DISTINCT M.AN8_CUNO, C.ABALPH " & _
                    "FROM " & DWGMas & " M, " & F0101 & " C " & _
                    "WHERE M.AN8_CUNO IN (" & strCunoList & ") " & _
                    "AND M.DWGTYPE = 15 " & _
                    "AND M.DSTATUS > 0 " & _
                    "AND M.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY UPPER(ABALPH)"
    End If

    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        combo.AddItem UCase(Trim(rst.Fields("ABALPH")))
        combo.ItemData(combo.NewIndex) = rst.Fields("AN8_CUNO")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub PopNonGPJDrawings(tmpStr As String)
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim nodX As Node
    Dim sYNode As String, sPNode As String, sSNode As String, sDNode As String
    Dim sDesc As String, sYear As String, sCPRJ As String
    Dim iType As Integer
    
    tvwConst(3).Nodes.Clear
    tvwConst(3).ImageList = ImageList1
    sYNode = "": sPNode = "": sSNode = "": sDNode = ""
    strSelect = "SELECT DM.DWGID, DM.DWGYR, DM.MCU, DM.DWGTYPE, " & _
                "DM.DWGDESC, DM.DWGNUM, DS.SHTID, DS.SHTSEQ, DS.SHTDESC, " & _
                "DD.DWFID, DD.DWFPATH, MC.MCDL01 AS PROJ " & _
                "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & _
                DWGDwf & " DD, " & F0006 & " MC " & _
                "WHERE DM.AN8_CUNO = " & CLng(tmpStr) & " " & _
                "AND DM.DWGTYPE = 15 " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DD.DWGID " & _
                "AND DS.SHTID = DD.SHTID " & _
                "AND DM.MCU LIKE MC.MCMCU " & _
                "ORDER BY DM.DWGYR, PROJ, DM.DWGDESC, DS.SHTDESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sYNode <> "y" & rst.Fields("DWGYR") Then
            sYNode = "y" & rst.Fields("DWGYR")
            sDesc = rst.Fields("DWGYR")
            iType = 4
            Set nodX = tvwConst(3).Nodes.Add(, , sYNode, sDesc, iType)
            sYear = rst.Fields("DWGYR"): sPNode = "": sSNode = "": sDNode = ""
        End If
        
        If sPNode <> "p" & sYear & "-" & Trim(rst.Fields("MCU")) Then
            sPNode = "p" & sYear & "-" & Trim(rst.Fields("MCU"))
            sDesc = Trim(rst.Fields("MCU")) & " -- " & Trim(rst.Fields("PROJ"))
            iType = 4
            Set nodX = tvwConst(3).Nodes.Add(sYNode, tvwChild, sPNode, sDesc, iType)
            sSNode = "": sDNode = ""
        End If
        
'''        If sSNode <> "s" & rst.Fields("DWGID") Then
'''            sSNode = "s" & rst.Fields("DWGID")
'''            sDesc = UCase(Trim(rst.Fields("DWGDESC")))
'''            iType = 4
'''            Set nodX = tvwConst(3).Nodes.Add(sPNode, tvwChild, sSNode, sDesc, iType)
'''            sDNode = ""
'''        End If
        
        If sDNode <> "d" & rst.Fields("DWFID") Then
            sDNode = "d" & rst.Fields("DWFID")
            sDesc = UCase(Trim(rst.Fields("DWGDESC")))
            iType = 3
            Set nodX = tvwConst(3).Nodes.Add(sPNode, tvwChild, sDNode, sDesc, iType)
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub



Public Sub PopAICHIDrawings(tmpStr As String)
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim nodX As Node
    Dim sLNode As String, sYNode As String, sPNode As String, sDNode As String
    Dim sDesc As String, sYear As String, sCPRJ As String
    Dim iType As Integer
    
    tvwConst(3).Nodes.Clear
    tvwConst(3).ImageList = ImageList1
    sLNode = "": sYNode = "": sPNode = "": sDNode = ""
    strSelect = "SELECT DM.DLOC, DM.DWGID, DM.DWGYR, DM.MCU, " & _
                "DM.ADDDTTM, DD.UPDDTTM, " & _
                "DM.DWGDESC, DD.DWFID, DD.DWFPATH, MC.MCDL01 AS PROJ " & _
                "FROM " & DWGMas & " DM, " & DWGSht & " DS, " & _
                DWGDwf & " DD, " & F0006 & " MC " & _
                "WHERE DM.AN8_CUNO = " & CLng(tmpStr) & " " & _
                "AND DM.DWGTYPE = 15 " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DD.DWGID " & _
                "AND DS.SHTID = DD.SHTID " & _
                "AND DD.DWFTYPE = -1 " & _
                "AND DM.MCU LIKE MC.MCMCU " & _
                "ORDER BY DM.DLOC, DM.MCU, DM.DWGDESC"
'''                "ORDER BY DM.DLOC, DM.DWGYR, DM.MCU, DM.DWGDESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If sLNode <> "l" & rst.Fields("DLOC") Then
            sLNode = "l" & rst.Fields("DLOC")
            Select Case UCase(rst.Fields("DLOC"))
                Case "BRC": sDesc = "BRC Imagination Arts"
                Case "HOL": sDesc = "The Hollomon Group"
                Case "TAN": sDesc = "Tanseisha"
                Case "GPJ": sDesc = "George P. Johnson Company"
            End Select
            iType = 4
            Set nodX = tvwConst(3).Nodes.Add(, , sLNode, sDesc, iType)
            sYNode = "": sPNode = "": sDNode = ""
        End If
        
'''        If sYNode <> "y" & rst.Fields("DLOC") & rst.Fields("DWGYR") Then
'''            sYNode = "y" & rst.Fields("DLOC") & rst.Fields("DWGYR")
'''            sDesc = rst.Fields("DWGYR")
'''            iType = 4
'''            Set nodX = tvwConst(3).Nodes.Add(sLNode, tvwChild, sYNode, sDesc, iType)
'''            sYear = rst.Fields("DWGYR"): sPNode = "": sDNode = ""
'''        End If
        
'''        If sPNode <> "p" & rst.Fields("DLOC") & sYear & "-" & Trim(rst.Fields("MCU")) Then
'''            sPNode = "p" & rst.Fields("DLOC") & sYear & "-" & Trim(rst.Fields("MCU"))
        If sPNode <> "p" & rst.Fields("DLOC") & "-" & Trim(rst.Fields("MCU")) Then
            sPNode = "p" & rst.Fields("DLOC") & "-" & Trim(rst.Fields("MCU"))
            sDesc = Trim(rst.Fields("PROJ")) & "  (" & Trim(rst.Fields("MCU")) & ")"
            iType = 4
'            Set nodX = tvwConst(3).Nodes.Add(sYNode, tvwChild, sPNode, sDesc, iType)
            Set nodX = tvwConst(3).Nodes.Add(sLNode, tvwChild, sPNode, sDesc, iType)
            sDNode = ""
        End If
        
        If sDNode <> "d" & rst.Fields("DWFID") Then
            sDNode = "d" & rst.Fields("DWFID")
            sDesc = UCase(Trim(rst.Fields("DWGDESC"))) & " ---- <" & _
                        UCase(Format(rst.Fields("UPDDTTM"), "DD-MMM-YYYY")) & ">"
            If rst.Fields("UPDDTTM") + 4 > Now Then
                If rst.Fields("UPDDTTM") = rst.Fields("ADDDTTM") Then
                    iType = 6
                Else
                    iType = 8
                End If
            Else
                iType = 3
            End If
            Set nodX = tvwConst(3).Nodes.Add(sPNode, tvwChild, sDNode, sDesc, iType)
            If iType = 6 Or iType = 8 Then ''DO PARENTS''
                nodX.Parent.Image = 7
                nodX.Parent.Parent.Image = 7
'''                nodX.Parent.Parent.Parent.Image = 7
            End If
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub DeleteDWF(pDWFID As Long)
    Dim strSelect As String, strDelete As String
    Dim rst As ADODB.Recordset
    Dim tDWGID As Long
    Dim sDWF As String, sZip As String
    Dim i As Integer
    
    If lDWFID = pDWFID Then
        MsgBox "The selected file cannot be deleted because " & _
                    "it is currently open in the main viewer.", vbExclamation, "Sorry..."
        Exit Sub
    End If
    
    For i = tvwConst.LBound To tvwConst.UBound
        If tDWFID(i) = pDWFID Then
'''            sst1.Width = 7000
            volView(i).src = ""
            volView(i).Visible = False
            cmdLoadDWF(i).Visible = False
            tDWFID(i) = 0
        End If
    Next i
    
    strSelect = "SELECT DWGID FROM ANNOTATOR.DWG_DWF " & _
                "WHERE DWFID = " & pDWFID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        tDWGID = rst.Fields("DWGID")
        rst.Close
        strSelect = "SELECT DWFPATH FROM ANNOTATOR.DWG_DWF " & _
                    "WHERE DWGID = " & tDWGID
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            ''KILL DWF FILE''
            sDWF = Trim(rst.Fields("DWFPATH"))
            If Dir(sDWF) <> "" Then
                Kill sDWF
            End If
            
            ''KILL ZIP FILE''
            sZip = Left(sDWF, Len(sDWF) - 4) & ".zip"
            If Dir(sZip) <> "" Then
                Kill sZip
            End If
            
            rst.MoveNext
        Loop
        rst.Close: Set rst = Nothing
        
        strDelete = "DELETE FROM ANNOTATOR.DWG_DWF WHERE DWGID = " & tDWGID
        Conn.Execute (strDelete)
        
        tvwConst(iTab).Nodes.Remove ("d" & pDWFID)
        
    Else
        rst.Close: Set rst = Nothing
        MsgBox "Unable to delete selected file", vbExclamation, "Sorry..."
    End If
    
End Sub

Public Function CheckForFloorplan(pCUNO As Long, pSHYR As Integer, pSHCD As Long) As Boolean
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT DWFID, DWFPATH " & _
                "From ANNOTATOR.DWG_DWF " & _
                "WHERE DWGID IN (" & _
                    "select dwgid " & _
                    "From ANNOTATOR.DWG_show " & _
                    "Where SHYR = " & pSHYR & " " & _
                    "and an8_cuno = " & pCUNO & " " & _
                    "and an8_shcd = " & pSHCD & "" & _
                ") " & _
                "AND DWFTYPE = 0"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        bViewSet = False
        volFP.src = Trim(rst.Fields("DWFPATH"))
        volFP.Tag = rst.Fields("DWFID")
        CheckForFloorplan = True
    Else
        volFP.src = ""
        volFP.Tag = 0
        CheckForFloorplan = False
    End If
    rst.Close: Set rst = Nothing
    
End Function
