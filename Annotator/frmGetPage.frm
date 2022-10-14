VERSION 5.00
Begin VB.Form frmGetPage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   LinkTopic       =   "Form1"
   ScaleHeight     =   315
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   60
   End
   Begin VB.ComboBox cboPage 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmGetPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pVal As Integer, pMax As Integer

Public Property Get PassVal() As Integer
    PassVal = pVal
End Property
Public Property Let PassVal(ByVal vNewValue As Integer)
    pVal = vNewValue
End Property

Public Property Get PassMax() As Integer
    PassMax = pMax
End Property
Public Property Let PassMax(ByVal vNewValue As Integer)
    pMax = vNewValue
End Property

'''Private Sub cmdOK_Click()
'''    iPDFPage = CInt(txtPage.Text)
'''    Unload Me
'''End Sub



Private Sub cboPage_Click()
    iPDFPage = CInt(cboPage.Text)
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Me.Top = frmGraphics.Top + (frmGraphics.Height - frmGraphics.ScaleHeight) _
                - ((frmGraphics.Width - frmGraphics.ScaleWidth) / 2) _
                + frmGraphics.picTools.Top + 30
    Me.Left = frmGraphics.Left + ((frmGraphics.Width - frmGraphics.ScaleWidth) / 2) _
                + frmGraphics.picPDFTools.Left + frmGraphics.picPage.Left _
                + frmGraphics.lblPage.Left + 30
    For i = 1 To pMax
        cboPage.AddItem i
    Next i
    
'''    txtPage.Text = pVal
'''    txtPage.SelLength = Len(pVal)
'''    txtPage.SelStart = 1
'''    txtPage.MaxLength = Len(CStr(pMax))
End Sub

'''Private Sub txtPage_Change()
'''    If Val(txtPage.Text) > pMax Or Val(txtPage.Text) < 1 Then
'''        txtPage.Text = pVal
'''        txtPage.SelLength = Len(pVal)
'''        txtPage.SelStart = 1
'''    Else
'''        cmdOK.Default = True
'''    End If
'''End Sub
'''
'''Private Sub txtPage_KeyPress(KeyAscii As Integer)
'''    Call CheckInteger(KeyAscii)
'''End Sub
'''
'''Public Sub CheckInteger(KeyAscii As Integer)
'''    If Not IsNumeric(Chr(KeyAscii)) Then
'''        If Not KeyAscii = vbKeyBack Then
'''            KeyAscii = 0
'''        End If
'''    End If
'''End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
End Sub
