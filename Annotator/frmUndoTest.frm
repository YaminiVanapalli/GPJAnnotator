VERSION 5.00
Begin VB.Form frmUndoTest 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmUndoTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Height          =   5520
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   19
      Left            =   8700
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   18
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   17
      Left            =   4740
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   16
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   15
      Left            =   780
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   14
      Left            =   8700
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   13
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   12
      Left            =   4740
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   11
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   10
      Left            =   780
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   9
      Left            =   8700
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   8
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   7
      Left            =   4740
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   6
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   5
      Left            =   780
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   4
      Left            =   8700
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   3
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   2
      Left            =   4740
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   1
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   1335
      Index           =   0
      Left            =   780
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "frmUndoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    
    lst.Clear
    For i = frmGraphics.imgUndo.LBound To frmGraphics.imgUndo.UBound
        lst.AddItem i
        img(i).Picture = frmGraphics.imgUndo(i).Picture
    Next i
End Sub


