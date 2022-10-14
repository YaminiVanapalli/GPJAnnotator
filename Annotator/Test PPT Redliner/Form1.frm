VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   180
      Width           =   1635
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawWidth       =   4
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   5655
      Left            =   60
      ScaleHeight     =   5655
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   60
      Width           =   7395
   End
   Begin VB.Image img 
      Height          =   375
      Left            =   7980
      Top             =   6540
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim maxX As Long, maxY As Long
Dim bRedding As Boolean
Dim xStr As Single, yStr As Single


Private Sub Command1_Click()
    Dim x1 As Long, y1 As Long
    Dim rPic As Double, rImg As Double
    x1 = pic.Width: y1 = pic.Height: rPic = x1 / y1
    
    img.Picture = Clipboard.GetData(): rImg = img.Width / img.Height
    
    If img.Width > x1 Or img.Height > y1 Then
        If rPic < rImg Then ''WIDTH DRIVES ASPECT RATIO''
            pic.Left = 60
            pic.Width = maxX
            pic.Height = img.Height * (x1 / img.Width)
            pic.Top = (Me.ScaleHeight - pic.Height) / 2
        ElseIf rPic > rImg Then ''HEIGHT DRIVES ASPECT RATIO''
            pic.Top = 60
            pic.Height = maxY
            pic.Width = img.Width * (y1 / img.Height)
            pic.Left = (Me.ScaleWidth - pic.Width) / 2
        End If
    Else
        If rPic < rImg Then ''WIDTH DRIVES ASPECT RATIO''
            pic.Width = img.Width
            pic.Left = (Me.ScaleWidth - pic.Width) / 2
            pic.Height = img.Height
            pic.Top = (Me.ScaleHeight - pic.Height) / 2
        ElseIf rPic > rImg Then ''HEIGHT DRIVES ASPECT RATIO''
            pic.Height = img.Height
            pic.Top = (Me.ScaleHeight - pic.Height) / 2
            pic.Width = img.Width
            pic.Left = (Me.ScaleWidth - pic.Width) / 2
        End If
    End If
        
        
        
'''    If img.Width < x1 And img.Height < y1 Then
'''
'''
'''
'''    Else
'''        pic.Width = img.Width
'''        pic.Height = img.Height
'''    End If
    pic.PaintPicture img.Picture, 0, 0, pic.Width, pic.Height
    
'    If Me.WindowState = 0 Then
'        Me.Width = (Me.Width - Me.ScaleWidth) + (pic.Left * 2) + pic.Width
'        Me.Height = (Me.Height - Me.ScaleHeight) + (pic.Top * 2) + pic.Height
'    End If
End Sub

Private Sub Form_Load()
    pic.Top = 60
    pic.Left = 60
End Sub

Private Sub Form_Resize()
    pic.Width = Me.ScaleWidth - 120
    pic.Height = Me.ScaleHeight - 120
    maxX = pic.Width
    maxY = pic.Height
End Sub

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
