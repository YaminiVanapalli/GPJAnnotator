VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmGfxList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graphics List..."
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7155
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
   ScaleHeight     =   7425
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":1228
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":1542
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":1E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGfxList.frx":26F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
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
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open File..."
      End
      Begin VB.Menu mnuDisplayGroup 
         Caption         =   "Display Page Group..."
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmGfxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lColor(0 To 1) As Long
Dim bOpen As Boolean
Dim lX As Long, lY As Long

Dim pSQL As String, pFrom As String
Dim pSize As Integer

Public Property Get PassSQL() As String
    PassSQL = pSQL
End Property
Public Property Let PassSQL(ByVal vNewValue As String)
    pSQL = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property

Public Property Get PassSize() As Integer
    PassSize = pSize
End Property
Public Property Let PassSize(ByVal vNewValue As Integer)
    pSize = vNewValue
End Property




Private Sub flx1_Click()
    If bOpen Then
        Me.PopupMenu mnuPopup, , lX, lY
    End If
End Sub

Private Sub flx1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > flx1.RowHeight(0) Then
        bOpen = True
        lX = flx1.Left + X: lY = flx1.Top + Y
    Else
        bOpen = False
    End If
End Sub

Private Sub Form_Load()
    
    If pSize = 0 Then pSize = 20
    
    lColor(0) = vbWindowBackground
    lColor(1) = vbButtonShadow
    
    flx1.ColWidth(0) = 0
    flx1.ColWidth(1) = 480: flx1.ColAlignment(1) = 4: flx1.TextMatrix(0, 1) = "Index"
    flx1.ColWidth(2) = 600: flx1.ColAlignment(2) = 4: flx1.TextMatrix(0, 2) = "Format"
    flx1.ColAlignment(3) = 1: flx1.FixedAlignment(3) = 4: flx1.TextMatrix(0, 3) = "Description"
    flx1.ColWidth(4) = 0
    flx1.ColWidth(5) = 0
    
    Call PopList
End Sub

Private Sub Form_Resize()
    flx1.Width = Me.ScaleWidth - (flx1.Left * 2)
    flx1.Height = Me.ScaleHeight - (flx1.Top * 2)
    
    If flx1.Rows * 480 < flx1.Height Then
        flx1.ColWidth(3) = flx1.Width - 1080
    Else
        flx1.ColWidth(3) = flx1.Width - 1080 - 250
    End If
    
End Sub

Public Sub PopList()
    Dim rst As ADODB.Recordset
    Dim i As Integer, iFormat As Integer, iColor As Integer, iCol As Integer
    
    i = 0
    Set rst = Conn.Execute(pSQL)
    Do While Not rst.EOF
        i = i + 1
        flx1.Rows = i + 1
        flx1.TextMatrix(i, 0) = rst.Fields("GID")
        flx1.TextMatrix(i, 1) = i
        flx1.TextMatrix(i, 3) = Trim(rst.Fields("GDESC"))
        flx1.TextMatrix(i, 4) = Trim(rst.Fields("GPATH"))
        Select Case UCase(Trim(rst.Fields("GFORMAT")))
            Case "JPG": iFormat = 1
            Case "BMP": iFormat = 2
            Case "PDF": iFormat = 3
            Case "PPT": iFormat = 4
            Case "PPS": iFormat = 5
            Case "AVI": iFormat = 6
            Case "MPG": iFormat = 7
            Case "MOV": iFormat = 8
        End Select
        flx1.Row = i: flx1.Col = 2
        Set flx1.CellPicture = ImageList1.ListImages(iFormat).Picture
        flx1.CellPictureAlignment = 4
        flx1.RowHeight(i) = 480
        
        For iCol = 1 To 3
            flx1.Col = iCol: flx1.CellBackColor = lColor(Int((i - 1) / pSize) Mod 2)
        Next iCol
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
            
            
    
End Sub

Private Sub mnuDisplayGroup_Click()
    Dim iStart As Integer, i As Integer
    
    Screen.MousePointer = 11
    
    i = CInt(flx1.TextMatrix(flx1.RowSel, 1))
    iStart = (Int((i - 1) / pSize) * pSize) + 1
    
    Select Case UCase(pFrom)
        Case "FRMDIL"
            frmDIL.picWait.Visible = True
            frmDIL.picWait.Refresh
            frmDIL.picInner(0).Visible = False
            Call frmDIL.GetGraphics(0, frmDIL.CurrSelect, iStart, frmDIL.tvwGraphics.SelectedItem.key)
            frmDIL.picInner(0).Visible = True
    
            frmDIL.picWait.Visible = False
    
            Call frmDIL.ResetBatch(iStart, flx1.Rows - 1, 0)
        
        Case "FRMGRAPHICS"
            frmGraphics.picWait.Visible = True
            frmGraphics.picWait.Refresh
            frmGraphics.picInner(frmGraphics.sst1.Tab).Visible = False
            Call frmGraphics.GetGraphics(frmGraphics.sst1.Tab, _
                        frmGraphics.sst1.Tab, _
                        CurrSelect(frmGraphics.sst1.Tab), _
                        iStart, _
                        frmGraphics.tvwGraphics(frmGraphics.sst1.Tab).SelectedItem.key)
            frmGraphics.picInner(frmGraphics.sst1.Tab).Visible = True
    
            frmGraphics.picWait.Visible = False
    
            Call frmGraphics.ResetBatch(iStart, flx1.Rows - 1, frmGraphics.sst1.Tab)
    End Select
    
    Screen.MousePointer = 0
    
    Unload Me
End Sub

Private Sub mnuOpenFile_Click()
    Select Case UCase(pFrom)
        Case "FRMDIL"
            frmDIL.picDirs.Visible = False
            frmDIL.imgDirs.ToolTipText = "Click to Open File Index..."
            frmDIL.bDirsOpen = False
            Set frmDIL.imgDirs.Picture = frmDIL.imlDirs.ListImages(1).Picture
            '///// TIME TO LOAD THE GRAPHIC \\\\\
            Call frmDIL.LoadGraphic(0, flx1.TextMatrix(flx1.RowSel, 0), _
                        flx1.TextMatrix(flx1.RowSel, 3), frmDIL.tvwGraphics.SelectedItem.Text)
            Call frmDIL.WhatSupDoc(CLng(flx1.TextMatrix(flx1.RowSel, 0)))
        Case "FRMGRAPHICS"
            If frmGraphics.chkClose(frmGraphics.sst1.Tab).value = 1 Then
                frmGraphics.sst1.Visible = False
                frmGraphics.imgDirs.ToolTipText = "Click to Open File Index..."
            End If
            '///// TIME TO LOAD THE GRAPHIC \\\\\
            Call frmGraphics.LoadGraphic(frmGraphics.sst1.Tab, _
                        "g" & CStr(flx1.TextMatrix(flx1.RowSel, 0)), _
                        flx1.TextMatrix(flx1.RowSel, 3), _
                        frmGraphics.tvwGraphics(frmGraphics.sst1.Tab).SelectedItem.key, frmGraphics.tvwGraphics(frmGraphics.sst1.Tab).SelectedItem.Text)
            
    End Select
    
    Unload Me
End Sub
