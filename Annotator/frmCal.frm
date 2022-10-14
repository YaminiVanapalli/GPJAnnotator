VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Date..."
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   3090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      Height          =   375
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Page Forward a Month"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Page Back a Month"
      Top             =   60
      Width           =   315
   End
   Begin MSACAL.Calendar cal1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      _Version        =   524288
      _ExtentX        =   5424
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2001
      Month           =   6
      Day             =   8
      DayLength       =   0
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tLeft As Long, tTop As Long

Public Property Get PassLeft() As Long
    PassLeft = tLeft
End Property
Public Property Let PassLeft(ByVal vNewValue As Long)
    tLeft = vNewValue
End Property

Public Property Get PassTop() As Long
    PassTop = tTop
End Property
Public Property Let PassTop(ByVal vNewValue As Long)
    tTop = vNewValue
End Property



Private Sub cal1_Click()
'''    tForm.tControl.Text = cal1.Value
    PassDate = cal1.value
    Unload Me
End Sub

Private Sub cmdBack_Click()
    cal1.value = DateAdd("m", -1, cal1.value)
End Sub

Private Sub cmdForward_Click()
    cal1.value = DateAdd("m", 1, cal1.value)
End Sub

Private Sub Form_Load()
'    cal1.Value = PassDate
    Me.Top = tTop - Me.Height
    Me.Left = tLeft
    If PassDate = Empty Then cal1.value = Now Else cal1.value = PassDate
End Sub
