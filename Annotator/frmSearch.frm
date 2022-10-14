VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphic Search..."
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter by Approver"
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
      Left            =   5820
      TabIndex        =   60
      Top             =   6000
      Width           =   1875
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   2535
      Index           =   8
      Left            =   5700
      TabIndex        =   58
      Top             =   6000
      Width           =   2415
      Begin VB.ListBox lstApprovers 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   59
         Top             =   300
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdKeywordSearch 
      Caption         =   "Search..."
      Height          =   375
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Keyword Search"
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
      Left            =   2820
      TabIndex        =   54
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame fraFilter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   7
      Left            =   2700
      TabIndex        =   48
      Top             =   120
      Width           =   5415
      Begin VB.OptionButton optOR 
         Caption         =   "Or"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   57
         ToolTipText     =   "Search will return files referencing ANY of these Keywords"
         Top             =   180
         Width           =   555
      End
      Begin VB.OptionButton optOR 
         Caption         =   "And"
         Height          =   255
         Index           =   0
         Left            =   4005
         TabIndex        =   56
         ToolTipText     =   "Search will return only files referencing ALL of these Keywords"
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtKeyword 
         Height          =   315
         Left            =   120
         TabIndex        =   53
         Top             =   300
         Width           =   2115
      End
      Begin VB.ListBox lstKeyAvail 
         Height          =   1620
         ItemData        =   "frmSearch.frx":08CA
         Left            =   120
         List            =   "frmSearch.frx":08CC
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   690
         Width           =   2115
      End
      Begin VB.ListBox lstKeyApply 
         Height          =   1815
         ItemData        =   "frmSearch.frx":08CE
         Left            =   3180
         List            =   "frmSearch.frx":08D0
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   480
         Width           =   2115
      End
      Begin VB.CommandButton cmdApply 
         Height          =   495
         Index           =   0
         Left            =   2400
         Picture         =   "frmSearch.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdApply 
         Height          =   495
         Index           =   1
         Left            =   2400
         Picture         =   "frmSearch.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1380
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced..."
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2160
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4860
      Top             =   8940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1156
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":12B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":140A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1564
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "Search..."
      Height          =   615
      Index           =   1
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter by Post Date"
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
      Left            =   240
      TabIndex        =   31
      Top             =   4740
      Width           =   1935
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Status Filter"
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
      Left            =   4680
      TabIndex        =   30
      Top             =   2760
      Width           =   1395
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Type Filter"
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
      Left            =   2820
      TabIndex        =   29
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "File Name Search"
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
      Left            =   4680
      TabIndex        =   28
      Top             =   4740
      Width           =   1815
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Client Filter"
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
      Left            =   240
      TabIndex        =   27
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1155
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   4740
      Width           =   4275
      Begin VB.CommandButton cmdCal 
         Height          =   315
         Index           =   1
         Left            =   3780
         Picture         =   "frmSearch.frx":16BE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtDATE 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   2340
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdCal 
         Height          =   315
         Index           =   0
         Left            =   1560
         Picture         =   "frmSearch.frx":1808
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtDATE 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblDateRange 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2115
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   3
         Left            =   2025
         TabIndex        =   25
         Top             =   540
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lbl1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Index           =   1
         Left            =   472
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1875
      Index           =   2
      Left            =   4560
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
      Begin VB.CheckBox chkGSTATUS 
         Caption         =   "Returned for Changes"
         Height          =   435
         Index           =   27
         Left            =   120
         TabIndex        =   61
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkGSTATUS 
         Caption         =   "Internal Draft"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkGSTATUS 
         Caption         =   "Client Draft"
         Height          =   315
         Index           =   20
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkGSTATUS 
         Caption         =   "Approved"
         Height          =   315
         Index           =   30
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1875
      Index           =   1
      Left            =   2700
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
      Begin VB.CheckBox chkGTYPE 
         Caption         =   "Presentations"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox chkGTYPE 
         Caption         =   "Graphic Layouts"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox chkGTYPE 
         Caption         =   "Graphic Files"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox chkGTYPE 
         Caption         =   "Digital Photos"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1155
      Index           =   3
      Left            =   4560
      TabIndex        =   6
      Top             =   4740
      Width           =   3555
      Begin VB.TextBox txtGDESC 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   540
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a String to search for in the File Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Visible         =   0   'False
         Width           =   3210
      End
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   5415
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Remove Un-Selected Clients"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "Remove Selected Clients"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.ListBox lstClient 
         Height          =   255
         Left            =   7000
         TabIndex        =   32
         Top             =   60
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtCuno 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdCuno 
         Height          =   315
         Left            =   960
         Picture         =   "frmSearch.frx":1952
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txtClient 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.CommandButton cmdClient 
         Height          =   315
         Left            =   4920
         Picture         =   "frmSearch.frx":1EDC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin MSFlexGridLib.MSFlexGrid flx1 
         Height          =   1995
         Left            =   120
         TabIndex        =   5
         Top             =   690
         Visible         =   0   'False
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   3519
         _Version        =   393216
         Rows            =   10
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
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
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter by Poster"
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
      Left            =   240
      TabIndex        =   35
      Top             =   2760
      Width           =   1635
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1875
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   2760
      Width           =   2415
      Begin VB.ListBox lstPosters 
         Height          =   1410
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Format Filter"
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
      Left            =   6540
      TabIndex        =   38
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Frame fraFilter 
      Enabled         =   0   'False
      Height          =   1875
      Index           =   6
      Left            =   6420
      TabIndex        =   39
      Top             =   2760
      Width           =   1695
      Begin VB.Frame fraDILOnly 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1515
         Left            =   780
         TabIndex        =   63
         Top             =   300
         Visible         =   0   'False
         Width           =   855
         Begin VB.CheckBox chkGFORMAT 
            Caption         =   "MOV"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   1140
            Width           =   1335
         End
         Begin VB.CheckBox chkGFORMAT 
            Caption         =   "MPG"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   66
            Top             =   780
            Width           =   1335
         End
         Begin VB.CheckBox chkGFORMAT 
            Caption         =   "AVI"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   420
            Width           =   1335
         End
         Begin VB.CheckBox chkGFORMAT 
            Caption         =   "PPT"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   60
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkGFORMAT 
         Caption         =   "PDF"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkGFORMAT 
         Caption         =   "BMP"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CheckBox chkGFORMAT 
         Caption         =   "JPG"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   555
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8700
      Width           =   2415
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "Search..."
      Height          =   375
      Index           =   0
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   $"frmSearch.frx":2466
      Height          =   1725
      Left            =   180
      TabIndex        =   47
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   2355
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblHdr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   $"frmSearch.frx":2551
      Height          =   1050
      Left            =   5700
      TabIndex        =   37
      Top             =   9480
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2475
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sAndOr As String
Dim bOpen As Boolean, bPopped As Boolean, bAdvanced As Boolean
Dim tFrom As String
Dim pBCC As Long

Public Property Get PassFrom() As String
    PassFrom = tFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    tFrom = vNewValue
End Property

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property


Private Sub chkFilter_Click(Index As Integer)
    Select Case chkFilter(Index).Value
        Case 0
            fraFilter(Index).Enabled = False
            Call SetFrame(Index, False)
        Case 1
            fraFilter(Index).Enabled = True
            Call SetFrame(Index, True)
            Select Case tFrom
                Case "GH", "GA": fraDILOnly.Visible = False
                Case "DIL": fraDILOnly.Visible = True
            End Select
    End Select
End Sub

'''Private Sub chkKeyword_Click()
'''    Select Case chkKeyword.value
'''        Case 0: fraKeyword.Enabled = False
'''        Case Is > 0: fraKeyword.Enabled = True
'''    End Select
'''
'''End Sub

Private Sub cmdAdvanced_Click()
    Dim lOpenHgt As Long
    Select Case tFrom
        Case "GH", "GA"
            lOpenHgt = Me.Height - Me.ScaleHeight + 9405 ''' 9885
'            fraDILOnly.Visible = False
        Case "DIL"
            lOpenHgt = Me.Height - Me.ScaleHeight + 6015 ''' 6500
'            fraDILOnly.Visible = True
    End Select
    
    If bOpen Then
        bOpen = False
        bAdvanced = False
'''        Me.Height = 3180
        Me.Height = Me.Height - Me.ScaleHeight + 2700 ''' 3180
        cmdKeywordSearch.Visible = True
'        Me.Top = (Screen.Height - Me.Height) / 2
    Else
        cmdKeywordSearch.Visible = False
        If Not bPopped Then
            Me.MousePointer = 11
            Call GetAdvanced
            bPopped = True
            Me.MousePointer = 0
        End If
        bOpen = True
        bAdvanced = True
        Me.Height = lOpenHgt
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    
End Sub

Private Sub cmdApply_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0
            For i = lstKeyAvail.ListCount - 1 To 0 Step -1
                If lstKeyAvail.Selected(i) = True Then
                    lstKeyApply.AddItem lstKeyAvail.List(i)
                    lstKeyApply.ItemData(lstKeyApply.NewIndex) = lstKeyAvail.ItemData(i)
                    lstKeyAvail.RemoveItem (i)
                End If
            Next i
            txtKeyword.Text = ""
            txtKeyword.SetFocus
        Case 1
            For i = lstKeyApply.ListCount - 1 To 0 Step -1
                If lstKeyApply.Selected(i) = True Then
                    lstKeyAvail.AddItem lstKeyApply.List(i)
                    lstKeyAvail.ItemData(lstKeyAvail.NewIndex) = lstKeyApply.ItemData(i)
                    lstKeyApply.RemoveItem (i)
                End If
            Next i
    End Select
End Sub


Private Sub cmdCal_Click(Index As Integer)
    If txtDATE(Index).Text = "" Then
        PassDate = Empty
    Else
        PassDate = DateValue(txtDATE(Index).Text)
    End If
    With frmCal
        .PassLeft = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) _
                    + fraFilter(4).Left + txtDATE(Index).Left
        .PassTop = Me.Top + (Me.Height - Me.ScaleHeight - 75) + fraFilter(4).Top _
                    + txtDATE(Index).Top '''+ txtDATE(Index).Height
        .Show 1
    End With
    
    If PassDate = Empty Then
        txtDATE(Index).Text = ""
    Else
        txtDATE(Index).Text = Format(PassDate, "DD-MMM-YYYY")
    End If
    
    lblDateRange.Caption = GetDateRange(txtDATE(0).Text, txtDATE(1).Text)
End Sub

Private Sub cmdClient_Click()
    Dim i As Integer, iLen As Integer
'    MsgBox "CLIENT"
    flx1.Visible = False
    flx1.Rows = 0
    txtClient.Text = CheckStars(txtClient.Text)
    Select Case Left(txtClient.Text, 1)
        Case "*"
            If txtClient.Text = "*" Then
                For i = 0 To lstClient.ListCount - 1
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                Next i
            ElseIf InStr(2, txtClient.Text, "*") = Len(txtClient.Text) Then
                For i = 0 To lstClient.ListCount - 1
                    If InStr(1, UCase(lstClient.List(i)), UCase(Mid(txtClient.Text, 2, Len(txtClient.Text) - 2))) <> 0 Then
                        flx1.Rows = flx1.Rows + 1
                        flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                        flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                    End If
                Next i
            End If
        Case Else
            For i = 0 To lstClient.ListCount - 1
                If InStr(1, txtClient.Text, "*") = Len(txtClient.Text) Then
                    iLen = Len(txtClient.Text) - 1
                Else
                    iLen = Len(txtClient.Text)
                End If
                If UCase(Left(lstClient.List(i), iLen)) = UCase(Left(txtClient.Text, iLen)) Then
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                End If
            Next i
    End Select
    flx1.Visible = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCuno_Click()
    Dim i As Integer, iLen As Integer

    flx1.Visible = False
    flx1.Rows = 0
    txtCuno.Text = CheckStars(txtCuno.Text)
    If Left(txtCuno.Text, 1) = "*" Then
        For i = 0 To lstClient.ListCount - 1
            If Right(txtCuno.Text, 1) = "*" Then
                If InStr(1, CStr(lstClient.ItemData(i)), Mid(txtCuno.Text, 2, Len(txtCuno.Text) - 2)) > 0 Then
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                End If
            Else
                If Right(CStr(lstClient.ItemData(i)), Len(txtCuno.Text) - 1) = Mid(txtCuno.Text, 2) Then
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                End If
            End If
        Next i
        
        
        
'''        For i = flx1.Rows - 1 To 0 Step -1
'''            If Right(txtCuno.Text, 1) = "*" Then
'''                If InStr(1, flx1.TextMatrix(i, 0), Mid(txtCuno.Text, 2, Len(txtCuno.Text) - 2)) = 0 Then
'''                    If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
'''                End If
'''            Else
'''                If Right(flx1.TextMatrix(i, 0), Len(txtCuno.Text) - 1) <> Mid(txtCuno.Text, 2) Then
'''                    If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
'''                End If
'''            End If
'''        Next i
                
    Else
        For i = 0 To lstClient.ListCount - 1
            If Right(txtCuno.Text, 1) = "*" Then
                iLen = Len(txtCuno.Text) - 1
                If Left(CStr(lstClient.ItemData(i)), iLen) = Left(txtCuno.Text, iLen) Then
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                End If
            Else
                iLen = Len(txtCuno.Text)
                If CStr(lstClient.ItemData(i)) = txtCuno.Text Then
                    flx1.Rows = flx1.Rows + 1
                    flx1.TextMatrix(flx1.Rows - 1, 0) = lstClient.ItemData(i)
                    flx1.TextMatrix(flx1.Rows - 1, 1) = lstClient.List(i)
                End If
            End If
        Next i
        
        
        
'''        For i = flx1.Rows - 1 To 0 Step -1
'''            If Right(txtCuno.Text, 1) = "*" Then
'''                iLen = Len(txtCuno.Text) - 1
'''                If Left(flx1.TextMatrix(i, 0), iLen) <> txtCuno.Text Then
'''                    If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
'''                End If
'''            Else
'''                iLen = Len(txtCuno.Text)
'''                If flx1.TextMatrix(i, 0) <> txtCuno.Text Then
'''                    If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
'''                End If
'''            End If
'''        Next i
    End If
    flx1.Visible = True
        
End Sub

Private Sub cmdFilter_Click(Index As Integer)
    Dim iStr As Integer, iEnd As Integer, i As Integer
    If flx1.RowSel > flx1.Row Then
        iStr = flx1.Row: iEnd = flx1.RowSel
    Else
        iEnd = flx1.Row: iStr = flx1.RowSel
    End If
    Select Case Index
        Case 0 ''REMOVE SELECTED ROWS''
            For i = iEnd To iStr Step -1
                If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
            Next i
        Case 1 ''REMOVE UN-SELECTED ROWS''
            For i = flx1.Rows - 1 To 0 Step -1
                If i < iStr Or i > iEnd Then
                    If flx1.Rows > 1 Then flx1.RemoveItem (i) Else flx1.Rows = 0
                End If
            Next i
    End Select
'''    MsgBox "Start = " & iStr & ", End = " & iEnd
End Sub

Private Sub cmdKeywordSearch_Click()
    Dim strSQL As String
    Dim i As Integer
    
    If lstKeyApply.ListCount = 0 Then
        MsgBox "No Search Criteria has been stipulated", vbExclamation, "Hey..."
        Exit Sub
    End If
    
    Select Case sAndOr
        Case "AND"
            For i = 0 To lstKeyApply.ListCount - 1
                If strSQL = "" Then
                    strSQL = "SELECT GM.GID FROM ANNOTATOR.GFX_METADATA_R GMX, ANNOTATOR.GFX_MASTER GM "
                Else
                    strSQL = strSQL & "INTERSECT SELECT GM.GID FROM ANNOTATOR.GFX_METADATA_R GMX, ANNOTATOR.GFX_MASTER GM "
                End If
                strSQL = strSQL & "WHERE GMX.GFX_METADATA_ID = " & lstKeyApply.ItemData(i) & " "
                strSQL = strSQL & "AND GMX.GID = GM.GID "
                If tFrom = "DIL" Then strSQL = strSQL & "AND GM.AN8_CUNO = 40579 "
            Next i
            
        Case "OR"
            For i = 0 To lstKeyApply.ListCount - 1
                If strSQL = "" Then
                    strSQL = "WHERE (GMX.GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
                Else
                    strSQL = strSQL & " OR GMX.GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
                End If
            Next i
            If strSQL = "" Then
                Exit Sub
            Else
                strSQL = strSQL & ")"
                strSQL = "SELECT GM.GID FROM ANNOTATOR.GFX_METADATA_R GMX, ANNOTATOR.GFX_MASTER GM " & strSQL & " AND GMX.GID = GM.GID"
                If tFrom = "DIL" Then strSQL = strSQL & " AND GM.AN8_CUNO = 40579"
            End If
            
    End Select
    
    If strSQL = "" Then Exit Sub
    
    frmSearchResults.PassFrom = tFrom
    frmSearchResults.PassSQL = strSQL
'''    frmSearchResults.PassTBL = "GFX_METADATA_R"
    frmSearchResults.Show 1, Me
'''    Call Me.ExecuteSearchSQL("GFX_METADATA_R", strSQL)
End Sub


Private Sub cmdSQL_Click(Index As Integer)
    Dim strHdr As String, strSQL As String, sCUNO As String, sGType As String, _
                sGStatus As String, sReplace As String, _
                sGDESC As String, sD1 As String, sD2 As String, sDate As String, _
                sNames As String, sGFormat As String, sFormat(0 To 2) As String
    Dim i As Integer
    Dim lCnt As Long
    Dim sIN As String
    
    lCnt = 0
    strHdr = ""
    strSQL = ""
    
'''    If chkFilter(7).value <> 0 And lstKeyApply.ListCount > 0 Then
'''        For i = 0 To lstKeyApply.ListCount - 1
'''            If strSQL = "" Then
'''                strSQL = "WHERE (GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
'''            Else
'''                strSQL = strSQL & " " & sAndOr & " GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
'''            End If
'''        Next i
'''        strSQL = strSQL & ")"
'''
'''        strHdr = "SELECT GM.GID FROM GFX_MASTER GM, GFX_METADATA_R GMR "
'''        strSQL = strSQL & " AND GMR.GID = GM.GID"
'''    Else
'''        strHdr = "SELECT GM.GID FROM GFX_MASTER GM "
'''    End If
    
    If chkFilter(7).Value <> 0 And lstKeyApply.ListCount > 0 Then
        Select Case sAndOr
            Case "AND"
                For i = 0 To lstKeyApply.ListCount - 1
                    If strSQL = "" Then
                        strHdr = "SELECT GM.GID FROM ANNOTATOR.GFX_MASTER GM WHERE GM.GID IN ("
                        strSQL = "SELECT GM.GID FROM ANNOTATOR.GFX_METADATA_R GMX, ANNOTATOR.GFX_MASTER GM "
                    Else
                        strSQL = strSQL & "INTERSECT SELECT GM.GID FROM ANNOTATOR.GFX_METADATA_R GMX, ANNOTATOR.GFX_MASTER GM "
                    End If
                    strSQL = strSQL & "WHERE GMX.GFX_METADATA_ID = " & lstKeyApply.ItemData(i) & " "
                    strSQL = strSQL & "AND GMX.GID = GM.GID "
                    If tFrom = "DIL" Then strSQL = strSQL & "AND GM.AN8_CUNO = 40579 "
                Next i
                If strSQL <> "" Then strSQL = strSQL & ") "
            Case "OR"
                For i = 0 To lstKeyApply.ListCount - 1
                    If strSQL = "" Then
                        strSQL = "WHERE (GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
                    Else
                        strSQL = strSQL & " " & sAndOr & " GFX_METADATA_ID = " & lstKeyApply.ItemData(i)
                    End If
                Next i
                strSQL = strSQL & ")"
                
                strHdr = "SELECT GM.GID FROM ANNOTATOR.GFX_MASTER GM, ANNOTATOR.GFX_METADATA_R GMR "
                strSQL = strSQL & " AND GMR.GID = GM.GID"
        End Select
    Else
        strHdr = "SELECT GM.GID FROM ANNOTATOR.GFX_MASTER GM "
    End If
    
    
    
    
    
    
    If chkFilter(0).Value = 1 Then
        If flx1.Rows < lstClient.ListCount Or tFrom = "GA" Then ''ALL CLIENTS''
            sCUNO = ""
            For i = 0 To flx1.Rows - 1
                If sCUNO = "" Then
                    sCUNO = flx1.TextMatrix(i, 0)
                Else
                    sCUNO = sCUNO & ", " & flx1.TextMatrix(i, 0)
                End If
            Next i
            If sCUNO <> "" Then
                If strSQL = "" Then
                    strSQL = "WHERE GM.AN8_CUNO IN (" & sCUNO & ")"
                Else
                    strSQL = strSQL & " AND GM.AN8_CUNO IN (" & sCUNO & ")"
                End If
            Else
                MsgBox "Your Client Filter as selected is filtering out all files", vbExclamation, "No Files will be available..."
                Exit Sub
            End If
        Else
            If Not bClientAll_Enabled Then
                strSQL = "WHERE GM.AN8_CUNO IN (" & strCunoList & ")"
            End If
        End If
    Else
        Select Case tFrom
            Case "GH", "GA"
                If Not bClientAll_Enabled Then
                    If strSQL = "" Then
                        strSQL = "WHERE GM.AN8_CUNO IN (" & strCunoList & ")"
                    Else
                        strSQL = strSQL & " AND GM.AN8_CUNO IN (" & strCunoList & ")"
                    End If
                End If
            Case "DIL"
                If strSQL = "" Then
                    strSQL = "WHERE GM.AN8_CUNO = 40579"
                Else
                    strSQL = strSQL & " AND GM.AN8_CUNO = 40579"
                End If
        End Select
    End If
    
    If chkFilter(1).Value = 1 Then
        sGType = ""
        For i = 1 To 4
            If chkGTYPE(i).Value = 1 Then
                If sGType = "" Then sGType = CStr(i) Else sGType = sGType & ", " & CStr(i)
            End If
        Next i
        If sGType <> "" Then
            If strSQL = "" Then
                strSQL = "WHERE GM.GTYPE IN (" & sGType & ")"
            Else
                strSQL = strSQL & " AND GM.GTYPE IN (" & sGType & ")"
            End If
        End If
    End If
    
    If chkFilter(2).Value = 1 Then
        sGStatus = ""
        For i = 1 To 3
            If chkGSTATUS(i * 10).Value = 1 Then
                If sGStatus = "" Then sGStatus = CStr(i * 10) _
                            Else sGStatus = sGStatus & ", " & CStr(i * 10)
            End If
        Next i
        If chkGSTATUS(27).Value = 1 Then
            If sGStatus = "" Then sGStatus = CStr(27) _
                        Else sGStatus = sGStatus & ", " & CStr(27)
        End If
            
        If sGStatus <> "" Then
            If strSQL = "" Then
                strSQL = "WHERE GM.GSTATUS IN (" & sGStatus & ")"
            Else
                strSQL = strSQL & " AND GM.GSTATUS IN (" & sGStatus & ")"
            End If
        Else
            If strSQL = "" Then
                If bGPJ Then
                    If tFrom = "GA" Then sIN = "10, 20, 27" Else sIN = "10, 20, 27, 30"
                Else
                    If tFrom = "GA" Then sIN = "20, 27" Else sIN = "20, 27, 30"
                End If
                strSQL = "WHERE GM.GSTATUS IN (" & sIN & ")"
            Else
                If bGPJ Then
                    If tFrom = "GA" Then sIN = "10, 20, 27" Else sIN = "10, 20, 27, 30"
                Else
                    If tFrom = "GA" Then sIN = "20, 27" Else sIN = "20, 27, 30"
                End If
                strSQL = strSQL & " AND GM.GSTATUS IN (" & sIN & ")"
            End If
        End If
    Else
        If Not bPerm(56) Then
            If tFrom = "GA" Then sIN = "20, 27" Else sIN = "20, 27, 30"
            If strSQL = "" Then
                strSQL = "WHERE GM.GSTATUS IN (" & sIN & ")"
            Else
                strSQL = strSQL & " AND GM.GSTATUS IN (" & sIN & ")"
            End If
        Else
            If strSQL = "" Then
                If bGPJ Then
                    If tFrom = "GA" Then sIN = "10, 20, 27" Else sIN = "10, 20, 27, 30"
                Else
                    If tFrom = "GA" Then sIN = "20, 27" Else sIN = "20, 27, 30"
                End If
                strSQL = "WHERE GM.GSTATUS IN (" & sIN & ")"
            Else
                If bGPJ Then
                    If tFrom = "GA" Then sIN = "10, 20, 27" Else sIN = "10, 20, 27, 30"
                Else
                    If tFrom = "GA" Then sIN = "20, 27" Else sIN = "20, 27, 30"
                End If
                strSQL = strSQL & " AND GM.GSTATUS IN (" & sIN & ")"
            End If
        End If
    End If
    
    If chkFilter(3).Value = 1 Then
        If txtGDESC <> "" Then
            ''DEAL WITH WILDCARD LATER''
            sReplace = Replace(txtGDESC.Text, "*", "%")
            If Left(sReplace, 1) <> "%" Then sReplace = "%" & sReplace
            If Right(sReplace, 1) <> "%" Then sReplace = sReplace & "%"
            If strSQL = "" Then
                strSQL = "WHERE UPPER(GM.GDESC) LIKE '" & _
                            UCase(sReplace) & "'"
            Else
                strSQL = strSQL & " AND UPPER(GM.GDESC) LIKE '" & _
                            UCase(sReplace) & "'"
            End If
        End If
    End If
    
    If chkFilter(4).Value = 1 Then
        sD1 = "": sD2 = ""
        If txtDATE(0).Text <> "" Then _
                    sD1 = "TO_DATE('" & txtDATE(0).Text & "', 'DD-MON-YYYY')"
        If txtDATE(1).Text <> "" Then _
                    sD2 = "TO_DATE('" & txtDATE(1).Text & "', 'DD-MON-YYYY')"
        If sD1 = "" And sD2 = "" Then
            sDate = ""
        ElseIf sD1 = "" Then
            sDate = "GM.ADDDTTM <= " & sD2
        ElseIf sD2 = "" Then
            sDate = "GM.ADDDTTM >= " & sD1
        Else
            sDate = "GM.ADDDTTM BETWEEN " & sD1 & " AND " & sD2
        End If
        If strSQL = "" Then
            strSQL = "WHERE " & sDate
        Else
            strSQL = strSQL & " AND " & sDate
        End If
    End If
    
    If chkFilter(5).Value = 1 Then
        sNames = ""
        For i = 0 To lstPosters.ListCount - 1
            If lstPosters.Selected(i) = True Then
                If sNames = "" Then
                    sNames = "'" & lstPosters.List(i) & "'"
                Else
                    sNames = sNames & ", '" & lstPosters.List(i) & "'"
                End If
            End If
        Next i
        If sNames <> "" Then
            If strSQL = "" Then
                strSQL = "WHERE GM.ADDUSER IN (" & sNames & ")"
            Else
                strSQL = strSQL & " AND GM.ADDUSER IN (" & sNames & ")"
            End If
        End If
    End If
    
    If chkFilter(6).Value = 1 Then
        sGFormat = ""
        sFormat(0) = "JPG": sFormat(1) = "BMP": sFormat(2) = "PDF"
        For i = chkGFORMAT.LBound To chkGFORMAT.UBound
            If chkGFORMAT(i).Value = 1 Then
                If sGFormat = "" Then sGFormat = "'" & chkGFORMAT(i).Caption & "'" _
                            Else sGFormat = sGFormat & ", '" & chkGFORMAT(i).Caption & "'"
            End If
        Next i
        If sGFormat <> "" Then
            If strSQL = "" Then
                strSQL = "WHERE UPPER(GM.GFORMAT) IN (" & sGFormat & ")"
            Else
                strSQL = strSQL & " AND UPPER(GM.GFORMAT) IN (" & sGFormat & ")"
            End If
        End If
    End If
    
    If chkFilter(8).Value = 1 Then
        sNames = ""
        For i = 0 To lstApprovers.ListCount - 1
            If lstApprovers.Selected(i) = True Then
                If sNames = "" Then
                    sNames = lstApprovers.ItemData(i)
                Else
                    sNames = sNames & ", " & lstApprovers.ItemData(i)
                End If
            End If
        Next i
        If sNames <> "" Then
            If strSQL = "" Then
                strSQL = "WHERE GM.GAPPROVER_ID IN (" & sNames & ")"
            Else
                strSQL = strSQL & " AND GM.GAPPROVER_ID IN (" & sNames & ")"
            End If
        End If
    End If
    
    strSQL = strHdr & strSQL
    Debug.Print strSQL
    
    frmSearchResults.PassFrom = tFrom
    frmSearchResults.PassSQL = strSQL
'''    frmSearchResults.PassTBL = "GFX_MASTER"
    frmSearchResults.Show 1, Me

'''    Call ExecuteSearchSQL("GFX_MASTER", strSQL)
    
'''    strSQL = "SELECT GID FROM GFX_MASTER " & strSQL
'''
'''    strSQL = "SELECT GM.AN8_CUNO AS CUNO, AB.ABALPH AS CLIENT, " & _
'''                "GM.GID, GM.GDESC, GM.GFORMAT " & _
'''                "FROM GFX_MASTER GM, " & F0101 & " AB " & _
'''                "WHERE GM.GID IN (" & strSQL & ") " & _
'''                "AND GM.AN8_CUNO = AB.ABAN8 " & _
'''                "ORDER BY CLIENT, CUNO, UPPER(GM.GDESC)"
'''    Call PopTree(strSQL)
    
    
'    MsgBox strSQL, vbInformation, "SQL..."
    
End Sub

Private Sub Form_Load()
    Dim ConnStr As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    bOpen = False
    bPopped = False
    sAndOr = "AND"
    
    If Not bPerm(56) Then chkGSTATUS(10).Enabled = False

    Call GetAllKeywords(bGPJ)
    If lstKeyAvail.ListCount = 0 Then
        chkFilter(7).Caption = "No Keywords"
        chkFilter(7).Value = 0
        chkFilter(7).Enabled = False
    End If
    
    If tFrom <> "GA" Then
        Me.Height = Me.Height - Me.ScaleHeight + 2700 ''' 3180
    Else
        chkGSTATUS(30).Enabled = False
        Call cmdAdvanced_Click
    End If
    
End Sub

Private Sub optOR_Click(Index As Integer)
    Select Case Index
        Case 0: sAndOr = "AND"
        Case 1: sAndOr = "OR"
    End Select
End Sub

Private Sub txtClient_GotFocus()
    cmdClient.Default = True
End Sub

Private Sub txtClient_KeyPress(KeyAscii As Integer)
    txtCuno.Text = ""
End Sub

Private Sub txtClient_LostFocus()
    cmdClient.Default = False
End Sub

Private Sub txtCuno_GotFocus()
    cmdCuno.Default = True
End Sub

Private Sub txtCuno_KeyPress(KeyAscii As Integer)
    txtClient.Text = ""
End Sub

Private Sub txtCuno_LostFocus()
    cmdCuno.Default = False
End Sub

Public Function GetDateRange(d1 As String, d2 As String)
    Dim sMess As String
    If d1 = "" And d2 = "" Then
        sMess = ""
    ElseIf d1 = "" Then
        sMess = "Posting Date <= " & d2
    ElseIf d2 = "" Then
        sMess = "Posting Date >= " & d1
    Else
        sMess = "Posting Date between " & d1 & " and " & d2
    End If
    GetDateRange = sMess
End Function


Public Function CheckStars(sChk As String)
    Dim i As Integer
    If sChk <> "*" _
                And sChk <> "**" _
                And Len(sChk) > 2 Then
        If InStr(1, Mid(sChk, 2, Len(sChk) - 2), "*") <> 0 Then
            For i = Len(sChk) - 2 To 1 Step -1
                If Mid(sChk, i + 1, 1) = "*" Then
                    sChk = Left(sChk, i) & Mid(sChk, i + 2)
                End If
            Next i
        End If
    End If
    CheckStars = sChk
            
End Function

Public Sub SetFrame(Index As Integer, bFlag As Boolean)
    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To frmSearch.Controls.Count
        If frmSearch.Controls(i).Container.Name = "fraFilter" _
                    And frmSearch.Controls(i).Container.Index = Index Then
            If Err Then
                Err.Clear
                GoTo NextOne
            End If
            frmSearch.Controls(i).Visible = bFlag
        End If
NextOne:
    Next i
            
End Sub

Private Sub txtKeyword_Change()
    Dim i As Integer, iCnt As Integer, iTop As Integer
    
    iCnt = 0: iTop = 0
    lstKeyAvail.Visible = False
    For i = lstKeyAvail.ListCount - 1 To 0 Step -1
        If Left(lstKeyAvail.List(i), Len(txtKeyword.Text)) = txtKeyword.Text Then
            lstKeyAvail.Selected(i) = True
            iTop = i
            iCnt = iCnt + 1
        Else
            lstKeyAvail.Selected(i) = False
        End If
        lstKeyAvail.Selected(i) = Left(lstKeyAvail.List(i), Len(txtKeyword.Text)) = txtKeyword.Text
    Next i
    If iTop - 2 >= 0 Then iTop = iTop - 2
    lstKeyAvail.TopIndex = iTop
'    lblCnt = "Selected Files:  " & CStr(iCnt)
    lstKeyAvail.Visible = True
End Sub

Public Sub GetAdvanced()
    Dim ConnStr As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    If tFrom = "GH" Then
        If Not bGPJ Then chkGSTATUS(10).Enabled = False
        flx1.Rows = 0
        If bClientAll_Enabled Then
            strSelect = "SELECT ABAN8 AS CUNO, UPPER(ABALPH) AS CLIENT " & _
                        "FROM " & F0101 & " " & _
                        "WHERE ABAN8 IN " & _
                        "(SELECT DISTINCT AN8_CUNO " & _
                        "FROM ANNOTATOR.GFX_MASTER) " & _
                        "ORDER BY CLIENT"
        Else
            strSelect = "SELECT ABAN8 AS CUNO, UPPER(ABALPH) AS CLIENT " & _
                        "FROM " & F0101 & " " & _
                        "WHERE ABAN8 IN " & _
                        "(SELECT DISTINCT AN8_CUNO " & _
                        "FROM ANNOTATOR.GFX_MASTER " & _
                        "WHERE AN8_CUNO IN (" & strCunoList & ")) " & _
                        "ORDER BY CLIENT"
        End If
    
    
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            flx1.Rows = flx1.Rows + 1
            flx1.TextMatrix(flx1.Rows - 1, 0) = rst.Fields("CUNO")
            flx1.TextMatrix(flx1.Rows - 1, 1) = Trim(rst.Fields("CLIENT"))
            lstClient.AddItem Trim(rst.Fields("CLIENT"))
            lstClient.ItemData(lstClient.NewIndex) = rst.Fields("CUNO")
            rst.MoveNext
        Loop
        rst.Close: Set rst = Nothing
        
        flx1.ColAlignment(0) = 3
        flx1.ColWidth(0) = 1200
        flx1.ColWidth(1) = flx1.Width - 1200
    
    ElseIf tFrom = "GA" Then
        If Not bGPJ Then chkGSTATUS(10).Enabled = False
        chkFilter(0).Value = 1
        chkFilter(0).Enabled = False
        flx1.Rows = 1
        flx1.TextMatrix(flx1.Rows - 1, 1) = frmGraphics.cboCUNO(4).Text
        flx1.TextMatrix(flx1.Rows - 1, 0) = frmGraphics.cboCUNO(4).ItemData(frmGraphics.cboCUNO(4).ListIndex)
        flx1.ColAlignment(0) = 3
        flx1.ColWidth(0) = 1200
        flx1.ColWidth(1) = flx1.Width - 1200
        fraFilter(0).Enabled = False
        
        Select Case frmGraphics.iTabStatus
            Case 0
                If bGPJ Then chkGSTATUS(10).Value = 1 Else chkGSTATUS(10).Enabled = False
                chkGSTATUS(20).Value = 1
                chkGSTATUS(27).Value = 1
            Case 1
                If bGPJ Then chkGSTATUS(10).Value = 1
            Case 2
                chkGSTATUS(20).Value = 1
            Case 3
                chkGSTATUS(27).Value = 1
        End Select
        chkFilter(2).Value = 1
        
    End If
    
    Select Case tFrom
        Case "GH"
            If bClientAll_Enabled Then
                strSelect = "SELECT DISTINCT ADDUSER " & _
                            "From ANNOTATOR.GFX_MASTER " & _
                            "WHERE  GSTATUS > 0 " & _
                            "ORDER BY UPPER(ADDUSER)"
            Else
                strSelect = "SELECT DISTINCT ADDUSER " & _
                            "From ANNOTATOR.GFX_MASTER " & _
                            "WHERE AN8_CUNO IN (" & strCunoList & ") " & _
                            "AND GSTATUS > 0 " & _
                            "ORDER BY UPPER(ADDUSER)"
            End If
        Case "GA"
            strSelect = "SELECT DISTINCT ADDUSER " & _
                        "From ANNOTATOR.GFX_MASTER " & _
                        "WHERE AN8_CUNO = " & pBCC & " " & _
                        "AND GSTATUS > 0 " & _
                        "ORDER BY UPPER(ADDUSER)"
        Case "DIL"
            strSelect = "SELECT DISTINCT ADDUSER " & _
                        "From ANNOTATOR.GFX_MASTER " & _
                        "WHERE AN8_CUNO = 40579 " & _
                        "AND GSTATUS > 0 " & _
                        "ORDER BY UPPER(ADDUSER)"
    End Select
        
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If Not IsNull(rst.Fields("ADDUSER")) Then lstPosters.AddItem Trim(rst.Fields("ADDUSER"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    Select Case tFrom
        Case "GH"
            If bClientAll_Enabled Then
                strSelect = "SELECT DISTINCT GM.GAPPROVER_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM ANNOTATOR.GFX_MASTER GM, IGLPROD.IGL_USER U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND GM.GSTATUS > 0 " & _
                            "AND GM.GAPPROVER_ID > 0 " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID (+) " & _
                            "ORDER BY UPPER(APPROVER)"
            Else
                strSelect = "SELECT DISTINCT GM.GAPPROVER_ID, " & _
                            "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                            "FROM ANNOTATOR.GFX_MASTER GM, IGLPROD.IGL_USER U " & _
                            "WHERE GM.GID > 0 " & _
                            "AND AN8_CUNO IN (" & strCunoList & ") " & _
                            "AND  GM.GSTATUS > 0 " & _
                            "AND GM.GAPPROVER_ID > 0 " & _
                            "AND GM.GAPPROVER_ID = U.USER_SEQ_ID " & _
                            "ORDER BY UPPER(APPROVER)"
            End If
        Case "GA"
            strSelect = "SELECT DISTINCT GM.GAPPROVER_ID, " & _
                        "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS APPROVER " & _
                        "FROM ANNOTATOR.GFX_MASTER GM, IGLPROD.IGL_USER U " & _
                        "WHERE GM.GID > 0 " & _
                        "AND GM.AN8_CUNO = " & pBCC & " " & _
                        "AND  GM.GSTATUS > 0 " & _
                        "AND GM.GAPPROVER_ID > 0 " & _
                        "AND GM.GAPPROVER_ID = U.USER_SEQ_ID " & _
                        "ORDER BY UPPER(APPROVER)"
    End Select
    If tFrom <> "DIL" Then
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            lstApprovers.AddItem StrConv(Trim(rst.Fields("APPROVER")), vbProperCase)
            lstApprovers.ItemData(lstApprovers.NewIndex) = rst.Fields("GAPPROVER_ID")
            rst.MoveNext
        Loop
        rst.Close: Set rst = Nothing
        lstApprovers.AddItem "<No Approver>"
        lstApprovers.ItemData(lstApprovers.NewIndex) = 0
        
        If tFrom = "GA" Then
            If frmGraphics.optApproverView(0).Value = True Then
                For i = 0 To lstApprovers.ListCount - 1
                    If lstApprovers.ItemData(i) = UserID Then
                        lstApprovers.Selected(i) = True
                        chkFilter(8).Value = 1
                        Exit For
                    End If
                Next i
            ElseIf frmGraphics.optApproverView(1).Value = True Then
                For i = 0 To lstApprovers.ListCount - 1
                    lstApprovers.Selected(i) = True
                    chkFilter(8).Value = 1
                Next i
            End If
        End If
    End If
    
End Sub

Private Sub txtKeyword_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub


Public Sub GetAllKeywords(bInternal As Boolean)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    Select Case tFrom
        Case "GH"
'''            strSelect = "SELECT GFX_METADATA_ID AS KEYID, " & _
'''                        "UPPER(VALUECHAR) AS KEYWORD " & _
'''                        "FROM GFX_METADATA " & _
'''                        "WHERE TYPE_CD = 305 " & _
'''                        "ORDER BY KEYWORD"
            If bInternal Then
                strSelect = "SELECT DISTINCT " & _
                            "GMX.GFX_METADATA_ID AS KEYID, " & _
                            "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                            "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                            "Where GMX.GFX_METADATA_ID > 0 " & _
                            "AND GMX.TYPE_CD = 305 " & _
                            "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                            "AND GMR.GID = GM.GID " & _
                            "AND GM.AN8_CUNO <> 40579 " & _
                            "ORDER BY KEYWORD"
            Else
                strSelect = "SELECT DISTINCT " & _
                            "GMX.GFX_METADATA_ID AS KEYID, " & _
                            "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                            "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                            "Where GMX.GFX_METADATA_ID > 0 " & _
                            "AND GMX.TYPE_CD = 305 " & _
                            "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                            "AND GMR.GID = GM.GID " & _
                            "AND GM.AN8_CUNO IN (" & strCunoList & ") " & _
                            "ORDER BY KEYWORD"
            End If
        Case "GA"
            strSelect = "SELECT DISTINCT " & _
                        "GMX.GFX_METADATA_ID AS KEYID, " & _
                        "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                        "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GMX.GFX_METADATA_ID > 0 " & _
                        "AND GMX.TYPE_CD = 305 " & _
                        "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                        "AND GMR.GID = GM.GID " & _
                        "AND GM.AN8_CUNO = " & pBCC & " " & _
                        "ORDER BY KEYWORD"
        Case "DIL"
            strSelect = "SELECT DISTINCT " & _
                        "GMX.GFX_METADATA_ID AS KEYID, " & _
                        "UPPER(GMX.VALUECHAR) As KEYWORD " & _
                        "FROM ANNOTATOR.GFX_METADATA GMX, ANNOTATOR.GFX_METADATA_R GMR, ANNOTATOR.GFX_MASTER GM " & _
                        "Where GMX.GFX_METADATA_ID > 0 " & _
                        "AND GMX.TYPE_CD = 305 " & _
                        "AND GMX.GFX_METADATA_ID = GMR.GFX_METADATA_ID " & _
                        "AND GMR.GID = GM.GID " & _
                        "AND GM.AN8_CUNO = 40579 " & _
                        "ORDER BY KEYWORD"
    End Select

    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstKeyAvail.AddItem Trim(rst.Fields("KEYWORD"))
        lstKeyAvail.ItemData(lstKeyAvail.NewIndex) = rst.Fields("KEYID")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub


