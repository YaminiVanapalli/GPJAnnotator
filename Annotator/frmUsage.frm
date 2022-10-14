VERSION 5.00
Begin VB.Form frmUsage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   7380
      Width           =   1095
   End
   Begin VB.TextBox txtList 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   1380
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   8175
   End
   Begin VB.Label lblQty 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   1140
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   427
      Picture         =   "frmUsage.frx":08CA
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMess As String, sTitle As String
Dim pQTY As Integer

Public Property Get PassMess() As String
    PassMess = sMess
End Property
Public Property Let PassMess(ByVal vNewValue As String)
    sMess = vNewValue
End Property

Public Property Get PassTitle() As String
    PassTitle = sTitle
End Property
Public Property Let PassTitle(ByVal vNewValue As String)
    sTitle = vNewValue
End Property


Public Property Get PassQty() As Integer
    PassQty = pQTY
End Property
Public Property Let PassQty(ByVal vNewValue As Integer)
    pQTY = vNewValue
End Property



Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblQty.Caption = pQTY
    txtList.Text = sMess
    Me.Caption = sTitle
End Sub
