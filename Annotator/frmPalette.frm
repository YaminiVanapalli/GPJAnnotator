VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   870
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPalette 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   240
      Width           =   870
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000F&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   15
         Left            =   60
         MouseIcon       =   "frmPalette.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   16
         Top             =   2820
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   14
         Left            =   60
         MouseIcon       =   "frmPalette.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   15
         Top             =   2640
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   13
         Left            =   60
         MouseIcon       =   "frmPalette.frx":0614
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   14
         Top             =   2460
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   12
         Left            =   60
         MouseIcon       =   "frmPalette.frx":091E
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   13
         Top             =   2280
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   11
         Left            =   60
         MouseIcon       =   "frmPalette.frx":0C28
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   12
         Top             =   2100
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   10
         Left            =   60
         MouseIcon       =   "frmPalette.frx":0F32
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   11
         Top             =   1920
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   9
         Left            =   60
         MouseIcon       =   "frmPalette.frx":123C
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   10
         Top             =   1740
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   8
         Left            =   60
         MouseIcon       =   "frmPalette.frx":1546
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   9
         Top             =   1560
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   7
         Left            =   60
         MouseIcon       =   "frmPalette.frx":1850
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   8
         Top             =   1380
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   6
         Left            =   60
         MouseIcon       =   "frmPalette.frx":1B5A
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   7
         Top             =   1200
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   5
         Left            =   60
         MouseIcon       =   "frmPalette.frx":1E64
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   6
         Top             =   1020
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   4
         Left            =   60
         MouseIcon       =   "frmPalette.frx":216E
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   5
         Top             =   840
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   60
         MouseIcon       =   "frmPalette.frx":2478
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   4
         Top             =   660
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   60
         MouseIcon       =   "frmPalette.frx":2782
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   3
         Top             =   480
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   60
         MouseIcon       =   "frmPalette.frx":2A8C
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   2
         Top             =   300
         Width           =   810
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   60
         MouseIcon       =   "frmPalette.frx":2D96
         MousePointer    =   99  'Custom
         ScaleHeight     =   120
         ScaleWidth      =   810
         TabIndex        =   1
         Top             =   60
         Width           =   810
      End
      Begin VB.Shape shpHL 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   180
         Left            =   0
         Top             =   0
         Width           =   870
      End
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   195
      Left            =   60
      MouseIcon       =   "frmPalette.frx":30A0
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim pLeft As Long, pTop As Long
Dim pColor As Integer

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

Public Property Get PassColor() As Integer
    PassColor = pColor
End Property
Public Property Let PassColor(ByVal vNewValue As Integer)
    pColor = vNewValue
End Property




Private Sub Form_Load()
    Dim i As Integer
    
    For i = picColor.LBound To picColor.UBound
        picColor(i).Top = 30 + ((picColor(0).Height + 60) * i)
        picColor(i).Left = 30
        picColor(i).Width = picPalette.ScaleWidth - 60
        picColor(i).Height = 120
        picColor(i).BackColor = QBColor(i)
    Next i
    picPalette.Height = 16 * 180
    shpHL.Top = pColor * 180
    
    Me.Height = (Me.Height - Me.ScaleHeight) + picPalette.Top + picPalette.Height
    lblCancel.Left = (Me.ScaleWidth - lblCancel.Width) / 2
    
    Me.Left = pLeft
    Me.Top = pTop - Me.Height
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub picColor_Click(Index As Integer)
    frmGraphics.iAnnoColor = Index
    Unload Me
End Sub
