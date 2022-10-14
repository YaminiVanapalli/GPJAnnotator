VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmOthers 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Drawing to view:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1695
   End
   Begin VB.Image imgClose 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3540
      MouseIcon       =   "frmOthers.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmOthers.frx":030A
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "frmOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pFCCD As Long

Public Property Get PassFCCD() As Long
    PassFCCD = pFCCD
End Property
Public Property Let PassFCCD(ByVal vNewValue As Long)
    pFCCD = vNewValue
End Property



Private Sub flx1_Click()
    Dim iRow As Integer
    iRow = flx1.RowSel
    With frmFacil
        .vol1.src = flx1.TextMatrix(iRow, 1)
        .vol1.Tag = flx1.TextMatrix(iRow, 0)
        .vol1.Update
        .lblReds.Caption = flx1.TextMatrix(iRow, 2)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    flx1.ColWidth(0) = 0
    flx1.ColWidth(1) = 0
    flx1.ColWidth(2) = flx1.Width
    
    Call GetOthers(pFCCD)
    flx1.Height = flx1.Rows * flx1.RowHeight(0)
    Me.Height = (Me.Height - Me.ScaleHeight) + flx1.Top + flx1.Height + flx1.Left
    
    Me.Top = frmFacil.Top + (frmFacil.Height - frmFacil.ScaleHeight - 30) _
                + frmFacil.imgOthers.Top + frmFacil.imgOthers.Height
    Me.Left = frmFacil.Left + ((frmFacil.Width - frmFacil.ScaleWidth) / 2) _
                + frmFacil.imgOthers.Left
                
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Public Sub GetOthers(pFID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iRow As Integer
    
    iRow = 0
    strSelect = "SELECT DF.DWFID, DF.DWFDESC, DF.DWFPATH, UPPER(DF.DWFDESC) AS UDESC, DF.DWFSTATUS " & _
                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHEET DS, ANNOTATOR.DWG_DWF DF " & _
                "Where DM.AN8_CUNO = " & pFID & " " & _
                "AND DM.DWGID = DS.DWGID " & _
                "AND DS.DWGID = DF.DWGID " & _
                "AND DS.SHTID = DF.SHTID " & _
                "ORDER BY DF.DWFSTATUS DESC, UDESC ASC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        flx1.Rows = iRow + 1
        flx1.TextMatrix(iRow, 0) = rst.Fields("DWFID")
        flx1.TextMatrix(iRow, 1) = Trim(rst.Fields("DWFPATH"))
        flx1.TextMatrix(iRow, 2) = Trim(rst.Fields("DWFDESC"))
        
        iRow = iRow + 1
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub
