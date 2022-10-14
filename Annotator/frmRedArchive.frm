VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Begin VB.Form frmRedArchive 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   13230
   Icon            =   "frmRedArchive.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOuter 
      BackColor       =   &H00404040&
      Height          =   7875
      Left            =   180
      ScaleHeight     =   7815
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   180
      Width           =   12855
      Begin VB.VScrollBar vsc1 
         Height          =   1455
         Left            =   12540
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   250
      End
      Begin VB.PictureBox picInner 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7755
         Left            =   0
         ScaleHeight     =   7755
         ScaleWidth      =   12555
         TabIndex        =   1
         Top             =   0
         Width           =   12555
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            ForeColor       =   &H80000008&
            Height          =   9900
            Index           =   0
            Left            =   0
            ScaleHeight     =   9870
            ScaleWidth      =   12405
            TabIndex        =   3
            Top             =   0
            Width           =   12435
            Begin VOLOVIEWXLibCtl.AvViewX vol 
               Height          =   4500
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   600
               Width           =   6000
               _cx             =   10583
               _cy             =   7937
               Appearance      =   0
               BorderStyle     =   0
               BackgroundColor =   "DefaultColors"
               Enabled         =   -1  'True
               UserMode        =   "Zoom"
               HighlightLinks  =   0   'False
               src             =   ""
               LayersOn        =   ""
               LayersOff       =   ""
               SrcTemp         =   ""
               SupportPath     =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
               FontPath        =   ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"
               NamedView       =   ""
               GeometryColor   =   "DefaultColors"
               PrintBackgroundColor=   "16777215"
               PrintGeometryColor=   "0"
               ShadingMode     =   "Gouraud"
               ProjectionMode  =   "Parallel"
               EnableUIMode    =   "DisableRightClickMenu"
               Layout          =   ""
               DisplayMode     =   -1
            End
            Begin VB.Label lblArchiveBy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ArchiveBy:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   6
               Top             =   300
               Width           =   900
            End
            Begin VB.Label lblAddBy 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "AddBy:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000009&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   60
               Width           =   585
            End
         End
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuRCPan 
         Caption         =   "Pan"
      End
      Begin VB.Menu mnuRCZoom 
         Caption         =   "Dynamic Zoom"
      End
      Begin VB.Menu mnuRCZoomW 
         Caption         =   "Zoom Window"
      End
      Begin VB.Menu mnuRCFullView 
         Caption         =   "Full View"
      End
      Begin VB.Menu mnuRCDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuRCDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmRedArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iD As Integer
Dim dLeft() As Double, dRight() As Double, dTop() As Double, dBottom() As Double
Dim bViewSet() As Boolean

Dim pDWGID As Long

Public Property Get PassDWGID() As Long
    PassDWGID = pDWGID
End Property
Public Property Let PassDWGID(ByVal vNewValue As Long)
    pDWGID = vNewValue
End Property


'''Private Sub Form_Activate()
'''    MsgBox "Activate Event"
'''    Dim i As Integer
'''
'''    For i = 0 To vol.Count - 1
'''        vol(i).GetCurrentView dLeft(i), dRight(i), dBottom(i), dTop(i)
'''    Next i
'''End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    
    picOuter.Top = 180
    picOuter.Left = 180
    picInner.Top = 0
    pic(0).Top = 0
    vsc1.Top = 0
    
    Call GetArchive(pDWGID)
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    picOuter.Width = Me.ScaleWidth - (picOuter.Left * 2)
    picOuter.Height = Me.ScaleHeight - picOuter.Top - picOuter.Left
    
    vsc1.Left = picOuter.ScaleWidth - vsc1.Width
    vsc1.Height = picOuter.ScaleHeight
    
    picInner.Width = picOuter.ScaleWidth - vsc1.Width
    
    For i = 0 To pic.Count - 1
        pic(i).Width = picInner.Width
    Next i
    
End Sub


Public Sub GetArchive(pID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Long
    
    i = -1
    strSelect = "SELECT DWFID, DWFPATH, DWFSTATUS, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM " & _
                "FROM ANNOTATOR.DWG_DWF " & _
                "WHERE DWGID = " & pID & " " & _
                "AND DWFTYPE = 9 " & _
                "AND DWFSTATUS >= 0 " & _
                "ORDER BY ADDDTTM DESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        i = i + 1
        
        ReDim Preserve dLeft(i)
        ReDim Preserve dRight(i)
        ReDim Preserve dTop(i)
        ReDim Preserve dBottom(i)
        
        ReDim Preserve bViewSet(i)
        bViewSet(i) = False
        
        If i >= pic.Count Then
            Load pic(i): Set pic(i).Container = picInner
            Load lblAddBy(i): Set lblAddBy(i).Container = pic(i)
            Load lblArchiveBy(i): Set lblArchiveBy(i).Container = pic(i)
            Load vol(i): Set vol(i).Container = pic(i)
        End If
        
        pic(i).Top = i * 9900 ''7000 ''9900
        pic(i).Width = picInner.Width
        pic(i).Height = 9900 ''7000 ''9900
        pic(i).Visible = True
        
        
        lblAddBy(i).Top = 60: lblAddBy(i).Left = 120
        lblArchiveBy(i).Top = 300: lblArchiveBy(i).Left = 120
        Select Case rst.Fields("DWFSTATUS")
            Case 0
                lblAddBy(i).ForeColor = vbWindowBackground
                lblArchiveBy(i).ForeColor = vbWindowBackground
                lblAddBy(i).Caption = "ARCHIVED Redline originally posted by " & UCase(Trim(rst.Fields("ADDUSER"))) & _
                            " on " & Format(rst.Fields("ADDDTTM"), "DDDD, MMMM D, YYYY (hh:nn AMPM)")
                lblArchiveBy(i).Caption = "Redline archived by " & UCase(Trim(rst.Fields("UPDUSER"))) & _
                            " on " & Format(rst.Fields("UPDDTTM"), "DDDD, MMMM D, YYYY (h:nn AMPM)")
'''                pic(i).BackColor = &H404040
            Case Else
                lblAddBy(i).ForeColor = vbGreen
                lblArchiveBy(i).ForeColor = vbGreen
                lblAddBy(i).Caption = "ACTIVE Redline posted by " & UCase(Trim(rst.Fields("ADDUSER"))) & _
                            " on " & Format(rst.Fields("ADDDTTM"), "DDDD, MMMM D, YYYY (h:nn AMPM)")
                If rst.Fields("ADDDTTM") <> rst.Fields("UPDDTTM") Then
                    lblArchiveBy(i).Caption = "Redline last updated by " & UCase(Trim(rst.Fields("UPDUSER"))) & _
                                " on " & Format(rst.Fields("UPDDTTM"), "DDDD, MMMM D, YYYY (h:nn AMPM)")
                Else
                    lblArchiveBy(i).Caption = ""
                End If
                
'''                pic(i).BackColor = vbGreen
        End Select
        lblAddBy(i).Visible = True
        
        
        
        lblArchiveBy(i).Visible = True
        
        vol(i).Top = 600: vol(i).Left = 120
        vol(i).Width = 12000: vol(i).Height = 9000
'        bViewSet = False
        vol(i).src = Trim(rst.Fields("DWFPATH"))
        vol(i).Tag = CStr(rst.Fields("DWFID"))
        vol(i).Visible = True
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    picInner.Height = pic.Count * pic(0).Height
    
    ''SIZE VSC HERE''
    Call SetScroll
    vsc1.Visible = (picInner.Height > picOuter.ScaleHeight)
        
    If pic.Count = 1 Then
        Me.Caption = "1 Archived Redline File..."
    Else
        Me.Caption = pic.Count & " Archived Redline Files..."
    End If
    
End Sub

Public Sub SetScroll()
    If picInner.Height > picOuter.ScaleHeight Then
        vsc1.Max = (picInner.Height / 100) - (picOuter.ScaleHeight / 100)
        vsc1.Visible = True
        vsc1.Value = 0
        vsc1.SmallChange = 1200 / 100
        vsc1.LargeChange = picOuter.ScaleHeight / 100
    Else
        picInner.Top = 0
        vsc1.Visible = False
    End If
End Sub

Private Sub mnuRCFullView_Click()
'    Dim b1 As Boolean, b2 As Boolean
'    vol(iD).GetDrawingExtents dLeft(iD), dRight(iD), dTop(iD), dBottom(iD)
    vol(iD).SetCurrentView dLeft(iD), dRight(iD), dBottom(iD), dTop(iD)
End Sub

Private Sub mnuRCPan_Click()
    Dim i As Integer
    mnuRCPan.Checked = True
    mnuRCZoom.Checked = False
    mnuRCZoomW.Checked = False
    For i = 0 To vol.Count - 1
        vol(iD).UserMode = "Pan"
    Next i
End Sub

Private Sub mnuRCPrint_Click()
'    Printer.Orientation = 2
    vol(iD).ShowPrintDialog
End Sub

Private Sub mnuRCZoom_Click()
    Dim i As Integer
    mnuRCPan.Checked = False
    mnuRCZoomW.Checked = False
    mnuRCZoom.Checked = True
    For i = 0 To vol.Count - 1
        vol(i).UserMode = "Zoom"
    Next i
End Sub

Private Sub mnuRCZoomW_Click()
    Dim i As Integer
    mnuRCPan.Checked = False
    mnuRCZoom.Checked = False
    mnuRCZoomW.Checked = True
    For i = 0 To vol.Count - 1
        vol(i).UserMode = "ZoomToRect"
    Next i
End Sub

Private Sub vol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Double, Y As Double)
    If Not bViewSet(Index) Then
        vol(Index).GetCurrentView dLeft(Index), dRight(Index), dBottom(Index), dTop(Index)
        bViewSet(Index) = True
    End If
    If Button = vbRightButton Then
        iD = Index
        Me.PopupMenu mnuRightClick
    End If
End Sub

'Private Sub vol_OnProgress(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusCode As Long, ByVal StatusText As String, bAbort As Boolean)
'    If bViewSet = False Then
'        If StatusCode = 42 Then
'            Call InitialView(Index)
'            bViewSet = True
'        End If
'    End If
'End Sub

Private Sub vsc1_Change()
    picInner.Top = CLng(vsc1.Value) * (-100)
End Sub

Private Sub vsc1_Scroll()
    picInner.Top = CLng(vsc1.Value) * (-100)
End Sub

Public Function InitialView(pIndex As Integer)
    vol(pIndex).GetCurrentView dLeft(pIndex), dRight(pIndex), dBottom(pIndex), dTop(pIndex)
End Function
