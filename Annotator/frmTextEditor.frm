VERSION 5.00
Begin VB.Form frmTextEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Editor..."
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1920
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   5775
   End
End
Attribute VB_Name = "frmTextEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pText As String
Dim pIndex As Integer

Public Property Get PassText() As String
    PassText = pText
End Property
Public Property Let PassText(ByVal vNewValue As String)
    pText = vNewValue
End Property

Public Property Get PassIndex() As Integer
    PassIndex = pIndex
End Property
Public Property Let PassIndex(ByVal vNewValue As Integer)
    pIndex = vNewValue
End Property


Private Sub cmdOK_Click()
    frmGraphics.lblRed(pIndex).Caption = Text1.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = pText
End Sub
