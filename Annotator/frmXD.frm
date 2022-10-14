VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmXD 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picXData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5460
      ScaleHeight     =   405
      ScaleWidth      =   1800
      TabIndex        =   58
      Top             =   3660
      Visible         =   0   'False
      Width           =   1800
      Begin VB.PictureBox picExpand 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   1420
         Picture         =   "frmXD.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   59
         Tag             =   "0"
         Top             =   90
         Width           =   240
      End
      Begin VB.PictureBox picExpand 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   1425
         Picture         =   "frmXD.frx":014A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
         Tag             =   "0"
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CommandButton cmdXData 
         Caption         =   "Extended Data...      "
         Height          =   405
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   1800
      End
   End
   Begin VB.PictureBox picXD 
      BackColor       =   &H80000002&
      Height          =   6795
      Left            =   0
      ScaleHeight     =   6735
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4995
      Begin VB.CommandButton Command1 
         Caption         =   "Save Changes"
         Height          =   405
         Left            =   30
         TabIndex        =   1
         Top             =   6300
         Width           =   1695
      End
      Begin TabDlg.SSTab sstXData 
         Height          =   6255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BackColor       =   -2147483646
         TabCaption(0)   =   "Graphic Data"
         TabPicture(0)   =   "frmXD.frx":0294
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "File Data"
         TabPicture(1)   =   "frmXD.frx":02B0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture1 
            FillColor       =   &H80000010&
            Height          =   5835
            Left            =   60
            ScaleHeight     =   5775
            ScaleWidth      =   4755
            TabIndex        =   5
            Top             =   360
            Width           =   4815
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   3720
               ScaleHeight     =   315
               ScaleWidth      =   1035
               TabIndex        =   35
               Top             =   2820
               Width           =   1035
               Begin VB.OptionButton optWgtUnit 
                  Caption         =   "lb"
                  Height          =   255
                  Index           =   1
                  Left            =   60
                  TabIndex        =   37
                  Top             =   30
                  Width           =   495
               End
               Begin VB.OptionButton optWgtUnit 
                  Caption         =   "kg"
                  Height          =   255
                  Index           =   2
                  Left            =   540
                  TabIndex        =   36
                  Top             =   30
                  Width           =   495
               End
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   2220
               Width           =   2415
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   3120
               Width           =   2415
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   3420
               Width           =   2415
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   3720
               Width           =   2415
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   1980
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   4020
               Width           =   2415
            End
            Begin VB.OptionButton optUnits 
               Caption         =   "mm"
               Height          =   315
               Index           =   8
               Left            =   2880
               TabIndex        =   29
               Top             =   60
               Width           =   615
            End
            Begin VB.OptionButton optUnits 
               Caption         =   "Inch"
               Height          =   315
               Index           =   1
               Left            =   1980
               TabIndex        =   28
               ToolTipText     =   "Interface accepts FT'-IN"" entry (12'- 4 1/2"")"
               Top             =   60
               Width           =   675
            End
            Begin VB.PictureBox picConvert 
               AutoRedraw      =   -1  'True
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   3480
               Picture         =   "frmXD.frx":02CC
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   27
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   1620
               Width           =   315
            End
            Begin VB.PictureBox picConvert 
               AutoRedraw      =   -1  'True
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   3480
               Picture         =   "frmXD.frx":0856
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   26
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   1320
               Width           =   315
            End
            Begin VB.PictureBox picConvert 
               AutoRedraw      =   -1  'True
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   3480
               Picture         =   "frmXD.frx":0DE0
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   25
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   1020
               Width           =   315
            End
            Begin VB.PictureBox picConvert 
               AutoRedraw      =   -1  'True
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   3480
               Picture         =   "frmXD.frx":136A
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   24
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   660
               Width           =   315
            End
            Begin VB.PictureBox picConvert 
               AutoRedraw      =   -1  'True
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   3480
               Picture         =   "frmXD.frx":18F4
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   23
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   360
               Width           =   315
            End
            Begin VB.PictureBox picEdit 
               AutoRedraw      =   -1  'True
               Height          =   315
               Index           =   4
               Left            =   4380
               Picture         =   "frmXD.frx":1E7E
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   22
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   4020
               Width           =   315
            End
            Begin VB.PictureBox picEdit 
               AutoRedraw      =   -1  'True
               Height          =   315
               Index           =   3
               Left            =   4380
               Picture         =   "frmXD.frx":2408
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   21
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   3720
               Width           =   315
            End
            Begin VB.PictureBox picEdit 
               AutoRedraw      =   -1  'True
               Height          =   315
               Index           =   2
               Left            =   4380
               Picture         =   "frmXD.frx":2992
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   20
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   3420
               Width           =   315
            End
            Begin VB.PictureBox picEdit 
               AutoRedraw      =   -1  'True
               Height          =   315
               Index           =   1
               Left            =   4380
               Picture         =   "frmXD.frx":2F1C
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   19
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   3120
               Width           =   315
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               Left            =   1980
               TabIndex        =   18
               Top             =   4020
               Width           =   2115
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               Left            =   1980
               TabIndex        =   17
               Top             =   3720
               Width           =   2115
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               Left            =   1980
               TabIndex        =   16
               Top             =   3420
               Width           =   2115
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               Left            =   1980
               TabIndex        =   15
               Top             =   3120
               Width           =   2115
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               Left            =   1980
               TabIndex        =   14
               Top             =   2220
               Width           =   2115
            End
            Begin VB.PictureBox picEdit 
               AutoRedraw      =   -1  'True
               Height          =   315
               Index           =   0
               Left            =   4380
               Picture         =   "frmXD.frx":34A6
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   13
               ToolTipText     =   "Click to Edit Reference List"
               Top             =   2220
               Width           =   315
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   1980
               TabIndex        =   12
               Top             =   1620
               Width           =   1515
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   1980
               TabIndex        =   11
               Top             =   2820
               Width           =   1695
            End
            Begin VB.TextBox Text1 
               Height          =   1095
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   4620
               Width           =   4635
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   1980
               TabIndex        =   9
               Top             =   1320
               Width           =   1515
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   1980
               TabIndex        =   8
               Top             =   1020
               Width           =   1515
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   1980
               TabIndex        =   7
               Top             =   660
               Width           =   1515
            End
            Begin VB.TextBox txtDim 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   1980
               TabIndex        =   6
               Top             =   360
               Width           =   1515
            End
            Begin VB.Image imgPrint 
               Height          =   240
               Left            =   4440
               Picture         =   "frmXD.frx":3A30
               ToolTipText     =   "Click to print Graphic Detail"
               Top             =   60
               Width           =   240
            End
            Begin VB.Label lblDim 
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
               Index           =   3
               Left            =   3900
               TabIndex        =   56
               Top             =   1380
               Width           =   45
            End
            Begin VB.Label lblDim 
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
               Index           =   2
               Left            =   3900
               TabIndex        =   55
               Top             =   1080
               Width           =   45
            End
            Begin VB.Label lblDim 
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
               Index           =   1
               Left            =   3900
               TabIndex        =   54
               Top             =   720
               Width           =   45
            End
            Begin VB.Label lblDim 
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
               Index           =   0
               Left            =   3900
               TabIndex        =   53
               Top             =   420
               Width           =   45
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Actual Thickness:"
               Height          =   195
               Index           =   6
               Left            =   240
               TabIndex        =   52
               Top             =   1680
               Width           =   1245
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Weight:"
               Height          =   195
               Index           =   14
               Left            =   240
               TabIndex        =   51
               Top             =   2880
               Width           =   570
            End
            Begin VB.Label lbl1X 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Notes:"
               Height          =   195
               Index           =   13
               Left            =   90
               TabIndex        =   50
               Top             =   4380
               Width           =   480
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fastener Type:"
               Height          =   195
               Index           =   12
               Left            =   240
               TabIndex        =   49
               Top             =   3780
               Width           =   1110
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graphic Producer:"
               Height          =   195
               Index           =   11
               Left            =   240
               TabIndex        =   48
               Top             =   4080
               Width           =   1290
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Media:"
               Height          =   195
               Index           =   10
               Left            =   240
               TabIndex        =   47
               Top             =   3180
               Width           =   480
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graphic Type:"
               Height          =   195
               Index           =   9
               Left            =   240
               TabIndex        =   46
               Top             =   3480
               Width           =   1005
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Miscellaneous:"
               Height          =   195
               Index           =   8
               Left            =   90
               TabIndex        =   45
               Top             =   2580
               Width           =   1035
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material:"
               Height          =   195
               Index           =   7
               Left            =   240
               TabIndex        =   44
               Top             =   2280
               Width           =   630
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Substrate:"
               Height          =   195
               Index           =   5
               Left            =   90
               TabIndex        =   43
               Top             =   1980
               Width           =   765
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Actual Height:"
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   42
               Top             =   1380
               Width           =   1020
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Actual Width/Length:"
               Height          =   195
               Index           =   3
               Left            =   240
               TabIndex        =   41
               Top             =   1080
               Width           =   1530
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nominal Height:"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   40
               Top             =   720
               Width           =   1125
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nominal Width/Length:"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   39
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label lblX 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Graphic Diminsions:"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   38
               Top             =   120
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   5835
            Left            =   -74940
            ScaleHeight     =   5775
            ScaleWidth      =   4755
            TabIndex        =   3
            Top             =   360
            Width           =   4815
            Begin VB.TextBox txtXData 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   5535
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   4
               Top             =   120
               Width           =   4515
            End
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Code does not work on this form until it is moved from frmGraphics"
      Height          =   1095
      Left            =   5340
      TabIndex        =   57
      Top             =   780
      Width           =   2955
   End
End
Attribute VB_Name = "frmXD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

