VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmUserLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GPJ Annotator User Log"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmUserLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccess 
      Caption         =   "View Client Access Rights"
      Height          =   375
      Left            =   9180
      TabIndex        =   4
      Top             =   60
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid flxFloorplans 
      Height          =   3795
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxGraphics 
      Height          =   3795
      Left            =   180
      TabIndex        =   2
      Top             =   4620
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Graphics Accessed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   4380
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Floorplans Accessed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "frmUserLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tUSERID As Long
Dim tUser As String

Public Property Get PassUserID() As Long
    PassUserID = tUSERID
End Property
Public Property Let PassUserID(ByVal vNewValue As Long)
    tUSERID = vNewValue
End Property

Public Property Get PassUser() As String
    PassUser = tUser
End Property
Public Property Let PassUser(ByVal vNewValue As String)
    tUser = vNewValue
End Property



Private Sub cmdAccess_Click()
    Dim sList As String, sMess As String, strSelect As String
    Dim rst As ADODB.Recordset
    
    sList = GetClientList(tUSERID)
    Select Case sList
        Case "ALL"
            sMess = tUser & " has access rights to all Clients."
        Case Else
            strSelect = "SELECT ABALPH FROM " & F0101 & " " & _
                        "WHERE ABAN8 IN (" & sList & ") " & _
                        "ORDER BY ABALPH"
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sMess = "" Then
                    sMess = Trim(rst.Fields("ABALPH"))
                Else
                    sMess = sMess & ", " & Trim(rst.Fields("ABALPH"))
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            sMess = tUser & " has access rights to the following Clients:" & _
                        vbNewLine & vbNewLine & sMess
    End Select
    MsgBox sMess, vbInformation, tUser
End Sub

Private Sub Form_Load()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iRow As Integer, iCol As Integer
    Dim GfxType(1 To 4) As String
    
    GfxType(1) = "Digital Photo"
    GfxType(2) = "Graphic Files"
    GfxType(3) = "Graphic Layout"
    GfxType(4) = "Presentation"
    
    Me.Caption = Me.Caption & "   [ " & tUser & " ]"
    Call SetGrid
    
    flxFloorplans.Rows = 1
    strSelect = "SELECT " & _
                "TO_CHAR(LL.LOCKOPENDTTM, 'MON DD YYYY')OPENDATE, " & _
                "TO_CHAR(LL.LOCKOPENDTTM, 'HH:MI AM')OPENTIME, " & _
                "DS.SHYR, S.ABALPH AS SHOW, C.ABALPH AS CLIENT, " & _
                "LL.LOCKSTATUS " & _
                "FROM " & ANOLockLog & " LL, " & DWGShow & " DS, " & _
                F0101 & " C, " & F0101 & " S " & _
                "WHERE LL.USER_SEQ_ID = " & tUSERID & " " & _
                "AND LL.LOCKREFSOURCE = 'DWG_MASTER' " & _
                "AND LL.LOCKREFID = DS.DWGID " & _
                "AND DS.AN8_CUNO = C.ABAN8 " & _
                "AND DS.AN8_SHCD = S.ABAN8 " & _
                "ORDER BY LL.LOCKOPENDTTM DESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        iRow = flxFloorplans.Rows
        flxFloorplans.Rows = flxFloorplans.Rows + 1
        flxFloorplans.TextMatrix(iRow, 0) = Trim(rst.Fields("OPENDATE"))
        flxFloorplans.TextMatrix(iRow, 1) = Trim(rst.Fields("OPENTIME"))
        flxFloorplans.TextMatrix(iRow, 2) = UCase(Trim(rst.Fields("CLIENT")))
        flxFloorplans.TextMatrix(iRow, 3) = CStr(rst.Fields("SHYR")) & " - " & _
                    UCase(Trim(rst.Fields("SHOW")))
        If Abs(rst.Fields("LOCKSTATUS")) = 1 Then
            flxFloorplans.TextMatrix(iRow, 4) = "Read Only"
        ElseIf Abs(rst.Fields("LOCKSTATUS")) = 2 Then
            flxFloorplans.TextMatrix(iRow, 4) = "Read Write"
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    Label1(0).Caption = CStr(flxFloorplans.Rows - 1) & " - Floorplans Accessed:"
    
        
    flxGraphics.Rows = 1
'    strSelect = "SELECT " & _
'                "TO_CHAR(LL.LOCKOPENDTTM, 'MON DD YYYY')OPENDATE, " & _
'                "TO_CHAR(LL.LOCKOPENDTTM, 'HH:MI AM')OPENTIME, " & _
'                "C.ABALPH AS CLIENT, GM.GDESC, GM.GTYPE, " & _
'                "LL.LOCKSTATUS " & _
'                "FROM " & ANOLockLog & " LL, " & GFXMas & " GM, " & _
'                F0101 & " C " & _
'                "WHERE LL.USER_SEQ_ID = " & tUSERID & " " & _
'                "AND LL.LOCKREFSOURCE = 'GFX_MASTER' " & _
'                "AND LL.LOCKREFID = GM.GID " & _
'                "AND GM.AN8_CUNO = C.ABAN8 " & _
'                "ORDER BY LL.LOCKOPENDTTM DESC"
'    Set rst = Conn.Execute(strSelect)
'    Do While Not rst.EOF
'        iRow = flxGraphics.Rows
'        flxGraphics.Rows = flxGraphics.Rows + 1
'        flxGraphics.TextMatrix(iRow, 0) = Trim(rst.Fields("OPENDATE"))
'        flxGraphics.TextMatrix(iRow, 1) = Trim(rst.Fields("OPENTIME"))
'        flxGraphics.TextMatrix(iRow, 2) = UCase(Trim(rst.Fields("CLIENT")))
'        flxGraphics.TextMatrix(iRow, 3) = UCase(Trim(rst.Fields("GDESC")))
'        flxGraphics.TextMatrix(iRow, 4) = GfxType(rst.Fields("GTYPE"))
'        If Abs(rst.Fields("LOCKSTATUS")) = 1 Then
'            flxGraphics.TextMatrix(iRow, 5) = "Read Only"
'        ElseIf Abs(rst.Fields("LOCKSTATUS")) = 2 Then
'            flxGraphics.TextMatrix(iRow, 5) = "Read Write"
'        End If
'        rst.MoveNext
'    Loop
'    rst.Close: Set rst = Nothing
'
'    Label1(1).Caption = CStr(flxGraphics.Rows - 1) & " - Graphics Accessed:"
End Sub

Public Sub SetGrid()
    Dim iCol As Integer
    
    
    flxFloorplans.ColWidth(0) = 1200
    flxFloorplans.ColAlignment(0) = 4
    flxFloorplans.TextMatrix(0, 0) = "Date"
    
    flxFloorplans.ColWidth(1) = 1200
    flxFloorplans.ColAlignment(1) = 4
    flxFloorplans.TextMatrix(0, 1) = "Time"
    
    flxFloorplans.Row = 0: flxFloorplans.Col = 2
    flxFloorplans.ColWidth(2) = (flxFloorplans.Width - 3900) * 0.4
    flxFloorplans.ColAlignment(2) = 1
    flxFloorplans.CellAlignment = 4
    flxFloorplans.TextMatrix(0, 2) = "Client"
    
    flxFloorplans.Row = 0: flxFloorplans.Col = 3
    flxFloorplans.ColWidth(3) = (flxFloorplans.Width - 3900) * 0.6
    flxFloorplans.ColAlignment(3) = 1
    flxFloorplans.CellAlignment = 4
    flxFloorplans.TextMatrix(0, 3) = "Show"
    
    flxFloorplans.ColWidth(4) = 1200
    flxFloorplans.ColAlignment(4) = 4
    flxFloorplans.TextMatrix(0, 4) = "Access"
    
    
    
    flxGraphics.ColWidth(0) = 1200
    flxGraphics.ColAlignment(0) = 4
    flxGraphics.TextMatrix(0, 0) = "Date"
    
    flxGraphics.ColWidth(1) = 1000
    flxGraphics.ColAlignment(1) = 4
    flxGraphics.TextMatrix(0, 1) = "Time"
    
    flxGraphics.Row = 0: flxGraphics.Col = 2
    flxGraphics.ColWidth(2) = (flxGraphics.Width - 4900) * 0.4
    flxGraphics.ColAlignment(2) = 1
    flxGraphics.CellAlignment = 4
    flxGraphics.TextMatrix(0, 2) = "Client"
    
    flxGraphics.Row = 0: flxGraphics.Col = 3
    flxGraphics.ColWidth(3) = (flxGraphics.Width - 4900) * 0.6
    flxGraphics.ColAlignment(3) = 1
    flxGraphics.CellAlignment = 4
    flxGraphics.TextMatrix(0, 3) = "Graphic"
    
    flxGraphics.ColWidth(4) = 1200
    flxGraphics.ColAlignment(4) = 4
    flxGraphics.TextMatrix(0, 4) = "File Type"
    
    flxGraphics.ColWidth(5) = 1200
    flxGraphics.ColAlignment(5) = 4
    flxGraphics.TextMatrix(0, 5) = "Access"
End Sub

Public Function GetClientList(UID As Long) As String
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim bClientAll As Boolean
    Dim sCList As String
    
    '///// DETERMINE IF USER HAS ACCESS TO ALL CLIENTS \\\\\
    strSelect = "SELECT CUNO_GROUP_ID " & _
                "FROM " & IGLUserCR & " " & _
                "WHERE USER_SEQ_ID = " & UID
    Set rst = Conn.Execute(strSelect)
    bClientAll = False
    If Not rst.EOF Then
        If CInt(rst.Fields("CUNO_GROUP_ID")) = -1 Then bClientAll = True
    End If
    rst.Close: Set rst = Nothing
    
    If bClientAll Then
        GetClientList = "ALL"
        Exit Function
    End If
    
    '///// IF NOT ALL, GET CLIENT LIST \\\\\
    strSelect = "SELECT AN8_CUNO FROM " & IGLUserCR & " " & _
                "WHERE USER_SEQ_ID = " & UID & " " & _
                "AND CUNO_GROUP_ID = 0 " & _
                "UNION " & _
                "SELECT GR.AN8_CUNO " & _
                "FROM " & IGLUserCR & " CR, " & IGLCGR & " GR " & _
                "WHERE CR.USER_SEQ_ID = " & UID & " " & _
                "AND CR.CUNO_GROUP_ID = GR.CUNO_GROUP_ID " & _
                "ORDER BY AN8_CUNO"
    Set rst = Conn.Execute(strSelect)
    sCList = ""
    If Not rst.EOF Then
        sCList = CStr(rst.Fields("AN8_CUNO"))
        rst.MoveNext
        Do While Not rst.EOF
            sCList = sCList & ", " & CStr(rst.Fields("AN8_CUNO"))
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    GetClientList = sCList
End Function


