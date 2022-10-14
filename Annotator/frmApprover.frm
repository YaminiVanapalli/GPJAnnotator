VERSION 5.00
Begin VB.Form frmApprover 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graphic Approver..."
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApprover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   780
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Approver"
      Height          =   615
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save New Approver"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1935
   End
   Begin VB.ListBox lstApprover 
      Height          =   1230
      ItemData        =   "frmApprover.frx":0442
      Left            =   120
      List            =   "frmApprover.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmApprover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lApproverID As Long, lCurrentID As Long

Dim pBCC As Long, pGID As Long
Dim pName As String, pHdr As String
Dim X1 As Long, Y1 As Long

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property

Public Property Get PassName() As String
    PassName = pName
End Property
Public Property Let PassName(ByVal vNewValue As String)
    pName = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = pHdr
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    pHdr = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property

Public Property Get PassX() As Long
    PassX = X1
End Property
Public Property Let PassX(ByVal vNewValue As Long)
    X1 = vNewValue
End Property

Public Property Get PassY() As Long
    PassY = Y1
End Property
Public Property Let PassY(ByVal vNewValue As Long)
    Y1 = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    Dim strUpdate As String, sNote As String
    Dim iErr As Integer
    
    strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                "GAPPROVER_ID = 0, " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE GID = " & pGID
    Conn.Execute (strUpdate)
    
    frmGraphics.flxApprove.TextMatrix(frmGraphics.flxApprove.RowSel, 6) = ""
    frmGraphics.flxApprove.Row = frmGraphics.flxApprove.RowSel
    frmGraphics.flxApprove.Col = 6
    frmGraphics.flxApprove.CellPictureAlignment = 4
    Set frmGraphics.flxApprove.CellPicture = frmGraphics.imgApprovers.Picture
    
    sNote = "Graphic Approver removed by " & StrConv(LogName, vbProperCase)
    iErr = InsertGfxComment(pGID, sNote)
    
End Sub

Private Sub cmdSave_Click()
    Dim iErr As Integer
    Dim strUpdate As String, sNote As String
    
    If lApproverID = lCurrentID Then Exit Sub
    
    strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                "GAPPROVER_ID = " & lApproverID & ", " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE GID = " & pGID
    Conn.Execute (strUpdate)
    
    frmGraphics.flxApprove.Row = frmGraphics.flxApprove.RowSel
    frmGraphics.flxApprove.Col = 6
    Set frmGraphics.flxApprove.CellPicture = LoadPicture("")
    frmGraphics.flxApprove.TextMatrix(frmGraphics.flxApprove.RowSel, 6) _
                = StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase)
    
    sNote = "Graphic Approver set to " & StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase) & _
                " by " & StrConv(LogName, vbProperCase)
    iErr = InsertGfxComment(pGID, sNote)
    
    Unload Me
        
End Sub

Private Sub Form_Load()
    Me.Left = X1
    If Y1 + Me.Height > Screen.Height Then Y1 = Y1 - Me.Height
    Me.Top = Y1
    Me.Caption = pHdr
    
    lApproverID = 0
    lCurrentID = 0
    Call GetApproverList(pBCC, pName)
End Sub

Public Sub GetApproverList(tBCC As Long, tName As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT (TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS MEMBER, " & _
                "U.USER_SEQ_ID AS MEMBERID, U.EMPLOYER " & _
                "FROM ANNOTATOR.ANO_EMAIL_TEAM ET, ANNOTATOR.ANO_EMAIL_TEAM_USER_R ETU, IGLPROD.IGL_USER U " & _
                "Where ET.AN8_CUNO = " & tBCC & " " & _
                "AND ET.TEAM_ID = ETU.TEAM_ID " & _
                "AND ETU.USER_SEQ_ID = U.USER_SEQ_ID"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstApprover.AddItem Trim(rst.Fields("MEMBER"))
        lstApprover.ItemData(lstApprover.NewIndex) = rst.Fields("MEMBERID")
        If lstApprover.List(lstApprover.NewIndex) = UCase(tName) Then
            lstApprover.Selected(lstApprover.NewIndex) = True
            lCurrentID = rst.Fields("MEMBERID")
        End If
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Private Sub lstApprover_Click()
    lApproverID = lstApprover.ItemData(lstApprover.ListIndex)
    If lstApprover.ItemData(lstApprover.ListIndex) <> lCurrentID Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    
End Sub



