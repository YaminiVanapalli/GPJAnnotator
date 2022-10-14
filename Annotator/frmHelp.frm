VERSION 5.00
Object = "{23319180-2253-11D7-BD2E-08004608C318}#3.0#0"; "XpdfViewerCtrl.ocx"
Object = "{A0369ABE-B6D8-11D3-901D-00207816FA15}#3.0#0"; "aghypertext.ocx"
Begin VB.Form frmHelp 
   Caption         =   "GPJ Annotator Help"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFiles 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      Begin VB.ListBox lstHelp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5580
         ItemData        =   "frmHelp.frx":08CA
         Left            =   180
         List            =   "frmHelp.frx":08CC
         MouseIcon       =   "frmHelp.frx":08CE
         TabIndex        =   1
         Top             =   180
         Width           =   5475
      End
      Begin VB.CommandButton cmd 
         Height          =   6075
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5835
      End
   End
   Begin VB.CommandButton cmd1 
      Height          =   270
      Left            =   1620
      TabIndex        =   4
      Top             =   7260
      Width           =   250
   End
   Begin AgHyperText.AgHyperTxt lnkPrint 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HyperLink       =   "Print..."
   End
   Begin XpdfViewerCtl.XpdfViewer xpdf1 
      Height          =   2115
      Left            =   4920
      TabIndex        =   8
      Top             =   2820
      Visible         =   0   'False
      Width           =   3855
      showScrollbars  =   -1  'True
      showBorder      =   -1  'True
      showPasswordDialog=   -1  'True
   End
   Begin VB.Image imgDirs 
      Height          =   480
      Left            =   60
      MouseIcon       =   "frmHelp.frx":0A18
      MousePointer    =   99  'Custom
      Picture         =   "frmHelp.frx":0D22
      ToolTipText     =   "Click to Close File Index"
      Top             =   60
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblFileDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   10650
      TabIndex        =   3
      Top             =   660
      Width           =   45
   End
   Begin VB.Label lblFile 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   90
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8565
      MouseIcon       =   "frmHelp.frx":186C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   180
      Width           =   510
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sHelpPath As String
Dim bPicked As Boolean


Private Sub imgDirs_Click()
    If picFiles.Visible = False Then
        xpdf1.Visible = False
        picFiles.Visible = True
        imgDirs.ToolTipText = "Click to Close File Index"
    Else
        picFiles.Visible = False
        xpdf1.Visible = True
        imgDirs.ToolTipText = "Click to Open File Index..."
    End If
    
'''    Select Case Trim(cmdFiles.Caption)
'''        Case "Help Files"
'''            picFiles.Visible = True
'''            cmdFiles.Caption = "          Close Files"
'''        Case "Close Files"
'''            picFiles.Visible = False
'''            cmdFiles.Caption = "          Help Files"
'''    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    bPicked = False
    
'    lstHelp.ForeColor = RGB(150, 150, 102)
    
    '///// GET HELPFILES \\\\\
    lstHelp.Clear
    strSelect = "SELECT HELPFILE, PERMNODE " & _
                "FROM " & ANOHelp & " " & _
                "WHERE APP_ID = 1002 " & _
                "AND HSTATUS > 0 " & _
                "ORDER BY SORT_ID"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If rst.Fields("PERMNODE") = -1 Then
            lstHelp.AddItem Left(Trim(rst.Fields("HELPFILE")), Len(Trim(rst.Fields("HELPFILE"))) - 4)
        ElseIf bPerm(rst.Fields("PERMNODE")) Then
            lstHelp.AddItem Left(Trim(rst.Fields("HELPFILE")), Len(Trim(rst.Fields("HELPFILE"))) - 4)
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(False)
'''    If bENABLE_PRINTERS Then
'''        cmd1.Enabled = True
'''        cmd1.Width = 1320
'''    Else
'''        cmd1.Enabled = False
'''        cmd1.Width = 1760
'''    End If
    lnkPrint.Enabled = bENABLE_PRINTERS
    
    ''SHOW CMD1 BE RESIZED ???''
    
    '\\\\\ -------------------------------------------------------- /////'
    
    'sHelpPath = "\\DETMSFS01\GPJAnnotator\HelpFiles\"
    sHelpPath = "C:\Users\SrilathaS\Documents\Annotator-master\HelpFiles\"
    
    xpdf1.Top = shpHDR.Height + 360 ''240
    xpdf1.ZOrder 1
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        If Me.Width > 2000 And Me.Height > 2000 Then
            xpdf1.Left = 120
            xpdf1.Width = Me.ScaleWidth - (xpdf1.Left * 2)
            xpdf1.Height = Me.ScaleHeight - xpdf1.Top - xpdf1.Left
            cmd1.Top = xpdf1.Top + 15
            cmd1.Left = xpdf1.Left + xpdf1.Width - cmd1.Width
            
            lblFileDate.Left = xpdf1.Left + xpdf1.Width - lblFileDate.Width
            lblFileDate.Top = xpdf1.Top - lblFileDate.Height - 30
            
            lblClose.Left = Me.ScaleWidth - 300 - lblClose.Width
            
            shpHDR.Width = Me.ScaleWidth
            
            picFiles.Width = Me.ScaleWidth
            picFiles.Height = Me.ScaleHeight - picFiles.Top
            cmd.Width = picFiles.ScaleWidth
            cmd.Height = picFiles.ScaleHeight
            lstHelp.Width = picFiles.ScaleWidth - (lstHelp.Left * 2)
            lstHelp.Height = picFiles.ScaleHeight - (lstHelp.Top * 2)
            
        End If
    End If
End Sub

Private Sub Image1_Click()

End Sub

'''Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.ForeColor = lGeo_Back '' vbWhite
'''End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lnkPrint_Click()
    xpdf1.printWithDialog
End Sub

'''Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    lblClose.ForeColor = vbWhite
'''End Sub

Private Sub lstHelp_Click()
    Dim sFile As String, sChk As String
    
    If bPicked Then
        Me.MousePointer = 11
'''        picHelp.Visible = False
        picFiles.Visible = False
        
        
        sFile = sHelpPath & lstHelp.List(lstHelp.ListIndex) & ".pdf"
        sChk = Dir(sFile, vbNormal)
        If sChk <> "" Then
            lblFileDate = "File Last Edited " & Format(FileDateTime(sFile), "dd-mmm-yyyy") & "."
'            pdf1.setShowToolbar (False)
            xpdf1.loadFile (sFile)
''            If Pdf1.Visible = False Then
''                Pdf1.Visible = True
            If xpdf1.Visible = False Then
                xpdf1.Visible = True
'                cmd1.Visible = True
                lnkPrint.Visible = True
            End If
            xpdf1.Zoom = xpdf1.zoomWidth
'            cmd1.Visible = True
            lblFile = lstHelp.List(lstHelp.ListIndex)
            imgDirs.Visible = True
        Else
            MsgBox "File could not be found.", vbExclamation, "File Access Error..."
            lblFile = ""
            lblFileDate = ""
        End If
        bPicked = False
        Me.MousePointer = 0
    End If
End Sub

Private Sub lstHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bPicked = True
End Sub

Private Sub lstHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iL As Integer
    iL = Int((Y + (lstHelp.TopIndex * 240)) / 240)
    If iL < lstHelp.ListCount Then lstHelp.Selected(iL) = True
End Sub


