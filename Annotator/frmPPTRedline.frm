VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPPTRedline 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PowerPoint Redliner..."
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "frmPPTRedline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPPTRedline.frx":08CA
   ScaleHeight     =   6480
   ScaleWidth      =   9345
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picReds 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   0
      Picture         =   "frmPPTRedline.frx":5874E
      ScaleHeight     =   4860
      ScaleWidth      =   1980
      TabIndex        =   4
      Top             =   1460
      Width           =   1980
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   4635
         Left            =   60
         ScaleHeight     =   4575
         ScaleWidth      =   1575
         TabIndex        =   6
         Top             =   120
         Width           =   1635
      End
      Begin VB.VScrollBar vsc1 
         Height          =   1395
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   4
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   4755
      Left            =   2040
      ScaleHeight     =   4755
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1440
      Width           =   3675
   End
   Begin MSComctlLib.ImageList imlSkins 
      Left            =   120
      Top             =   5280
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
            Picture         =   "frmPPTRedline.frx":B9272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPPTRedline.frx":111106
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H000E5838&
      Height          =   195
      Left            =   8580
      MouseIcon       =   "frmPPTRedline.frx":168F9A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   540
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
      ForeColor       =   &H000E5838&
      Height          =   240
      Left            =   8565
      MouseIcon       =   "frmPPTRedline.frx":1692A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   180
      Width           =   510
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
      ForeColor       =   &H000E5838&
      Height          =   240
      Left            =   300
      MouseIcon       =   "frmPPTRedline.frx":1695AE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
      Width           =   465
   End
   Begin VB.Image img 
      Height          =   375
      Left            =   7980
      Top             =   6540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgMenu 
      Height          =   570
      Left            =   0
      Picture         =   "frmPPTRedline.frx":1698B8
      Top             =   900
      Width           =   1080
   End
   Begin VB.Image imgClose 
      Height          =   945
      Left            =   8280
      Picture         =   "frmPPTRedline.frx":169E89
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "frmPPTRedline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim maxX As Long, maxY As Long
Dim bRedding As Boolean
Dim xStr As Single, yStr As Single
Dim rPic As Double, rImg As Double
Dim lBackColor(0 To 1) As Long
Dim iBC As Integer

Dim pLeft As Long, pTop As Long, pWidth As Long, pHgt As Long



Public Property Get PassLeft() As Long
    PassLeft = pLeft
End Property
Public Property Let PassLeft(ByVal vNewValue As Long)
    pLeft = vNewValue
End Property

Public Property Get PassTop() As Long
    PassTop = pTop
End Property
Public Property Let PassTop(ByVal vNewValue As Long)
    pTop = vNewValue
End Property

Public Property Get PassWidth() As Long
    PassWidth = pWidth
End Property
Public Property Let PassWidth(ByVal vNewValue As Long)
    pWidth = vNewValue
End Property

Public Property Get PassHeight() As Long
    PassHeight = pHgt
End Property
Public Property Let PassHeight(ByVal vNewValue As Long)
    pHgt = vNewValue
End Property



Private Sub Form_Click()
    iBC = Abs(iBC - 1)
    Set Me.Picture = imlSkins.ListImages(iBC + 1).Picture
    Me.BackColor = lBackColor(iBC)
'    pic.BackColor = lBackColor(iBC)
End Sub

Private Sub Form_Load()
    lBackColor(0) = vbBlack
    lBackColor(1) = vbWhite
    iBC = 0
    pic.Top = 1440 ''1500 ''pTop ''60
    pic.Left = 2100 ''60 ''pLeft ''60
    pic.Width = pWidth
    pic.Height = pHgt
    
    Me.WindowState = frmDIL.WindowState
    Me.Width = frmDIL.Width
    Me.Height = frmDIL.Height
    
    img.Picture = Clipboard.GetData(): rImg = img.Width / img.Height
    pic.PaintPicture img.Picture, 0, 0, pic.Width, pic.Height
End Sub

Private Sub Form_Paint()
'''    pic.PaintPicture Clipboard.GetData(), 0, 0, pic.Width, pic.Height
    
    
'''    Set img.Picture = LoadPicture("")
End Sub

Private Sub Form_Resize()
    imgClose.Left = Me.ScaleWidth - imgClose.Width
    lblHelp.Left = imgClose.Left + (imgClose.Width / 2) - (lblHelp.Width / 2)
    lblClose.Left = imgClose.Left + (imgClose.Width / 2) - (lblClose.Width / 2)
    
    picReds.Height = Me.ScaleHeight - picReds.Top
    
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = lGeo_Back '' vbWhite
    lblClose.ForeColor = lGeo_Back '' vbWhite
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblClose.ForeColor = vbWhite
End Sub

Private Sub lblHelp_Click()
    lblHelp.ForeColor = vbWhite '' lColor
    frmHelp.Show 1
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.ForeColor = vbWhite
End Sub

'Private Sub Form_Resize()
'    Dim x1 As Long, y1 As Long
'
'
'    pic.Width = Me.ScaleWidth - 120
'    pic.Height = Me.ScaleHeight - 120
'    maxX = pic.Width
'    maxY = pic.Height
'
'    x1 = pic.Width: y1 = pic.Height: rPic = x1 / y1
'
'    If img.Width > x1 Or img.Height > y1 Then
'        If rPic < rImg Then ''WIDTH DRIVES ASPECT RATIO''
'            pic.Left = 60
'            pic.Width = maxX
'            pic.Height = img.Height * (x1 / img.Width)
'            pic.Top = (Me.ScaleHeight - pic.Height) / 2
'        ElseIf rPic > rImg Then ''HEIGHT DRIVES ASPECT RATIO''
'            pic.Top = 60
'            pic.Height = maxY
'            pic.Width = img.Width * (y1 / img.Height)
'            pic.Left = (Me.ScaleWidth - pic.Width) / 2
'        End If
'    Else
'        If rPic < rImg Then ''WIDTH DRIVES ASPECT RATIO''
'            pic.Width = img.Width
'            pic.Left = (Me.ScaleWidth - pic.Width) / 2
'            pic.Height = img.Height
'            pic.Top = (Me.ScaleHeight - pic.Height) / 2
'        ElseIf rPic > rImg Then ''HEIGHT DRIVES ASPECT RATIO''
'            pic.Height = img.Height
'            pic.Top = (Me.ScaleHeight - pic.Height) / 2
'            pic.Width = img.Width
'            pic.Left = (Me.ScaleWidth - pic.Width) / 2
'        End If
'    End If
'End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bRedding = True
    xStr = X: yStr = Y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRedding Then
        pic.Line (xStr, yStr)-(X, Y)
        xStr = X: yStr = Y
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bRedding = False
End Sub
