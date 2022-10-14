VERSION 5.00
Begin VB.Form frmOtherViews 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Drawing to View..."
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
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
   ScaleHeight     =   2970
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDWFPath 
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.ListBox lstOthers 
      Height          =   2010
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1395
   End
End
Attribute VB_Name = "frmOtherViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tSHTID As Long

Public Property Get PassSHTID() As Long
    PassSHTID = tSHTID
End Property
Public Property Let PassSHTID(ByVal vNewValue As Long)
    tSHTID = vNewValue
End Property


Private Sub cmdCancel_Click()
    Unload Me
End Sub

'''Private Sub cmdOK_Click()
'''    With frmAnnotator
'''        .lblReds.Caption = lstOthers.List(lstOthers.ListIndex)
'''        .volFrame.src = lstDWFPath.List(lstOthers.ListIndex)
'''    End With
'''    Unload Me
'''End Sub

Private Sub Form_Load()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    '///// POP LIST OF OTHER VIEWS (DWFTYPE=8) IN DWG_DWFS, WHERE DWG_DWFS.SHTID=tSHTID \\\\\
    strSelect = "SELECT DWF.DWFDESC, DWF.DWFPATH " & _
                "FROM " & DWGShow & " SHO, " & DWGDwf & " DWF " & _
                "WHERE SHO.AN8_CUNO = " & CLng(BCC) & " " & _
                "AND SHO.SHYR = " & SHYR & " " & _
                "AND SHO.AN8_SHCD = " & SHCD & " " & _
                "AND SHO.DWGID = DWF.DWGID " & _
                "AND DWF.DWFTYPE = 8 " & _
                "AND DWF.DWFSTATUS > 0 " & _
                "ORDER BY DWF.DWFDESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstOthers.AddItem UCase(Trim(rst.Fields("DWFDESC")))
        lstDWFPath.AddItem Trim(rst.Fields("DWFPATH"))
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Private Sub lstOthers_Click()
    Dim sCheck As String
    
    With frmAnnotator
        .lblReds.Caption = lstOthers.List(lstOthers.ListIndex)
        .volFrame.src = lstDWFPath.List(lstOthers.ListIndex)
    End With
    
    ''CHECK FOR PDF''
    sCheck = Dir(Left(frmAnnotator.volFrame.src, Len(frmAnnotator.volFrame.src) - 3) & "pdf")
    If sCheck <> "" Then
        frmAnnotator.mnuDownloadPDF.Enabled = True
        frmAnnotator.mnuEmailPDF.Enabled = True
    Else
        frmAnnotator.mnuDownloadPDF.Enabled = False
        frmAnnotator.mnuEmailPDF.Enabled = False
    End If
    
    
    Unload Me
'''    cmdOK.Enabled = True
'''    cmdOK.Default = True
End Sub
