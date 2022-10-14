VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmUsageGannt 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   Icon            =   "frmUsageGannt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOpt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4860
      TabIndex        =   3
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkGridlines 
         Alignment       =   1  'Right Justify
         Caption         =   "Gridlines"
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
         Index           =   1
         Left            =   2820
         TabIndex        =   7
         Top             =   420
         Width           =   1035
      End
      Begin VB.CheckBox chkGridlines 
         Alignment       =   1  'Right Justify
         Caption         =   "Gridlines"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   420
         Width           =   1035
      End
      Begin VB.OptionButton opt1 
         Alignment       =   1  'Right Justify
         Caption         =   "Daily User Count View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   5
         Top             =   180
         Width           =   1935
      End
      Begin VB.OptionButton opt1 
         Alignment       =   1  'Right Justify
         Caption         =   "Daily Gantt View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx2 
      Height          =   4815
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   50
      Cols            =   10
      FixedCols       =   2
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      MergeCells      =   2
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
      MouseIcon       =   "frmUsageGannt.frx":08CA
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4815
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   10
      FixedCols       =   4
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      MergeCells      =   2
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
   Begin VB.CommandButton cmdSetDate 
      Caption         =   "Set Date..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3180
   End
   Begin VB.CheckBox chkIncludeDev 
      Caption         =   "include Developer"
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
      Left            =   3420
      TabIndex        =   8
      Top             =   420
      Width           =   1695
   End
End
Attribute VB_Name = "frmUsageGannt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iType As Integer, iIncludeDev As Integer
Public CurrDate As Date
Dim bPopped As Boolean


Private Sub chkGridlines_Click(Index As Integer)
    Select Case Index
        Case 0
            If chkGridLines(0).Value = 0 Then
                flx1.GridLines = flexGridNone
            Else
                flx1.GridLines = flexGridFlat
            End If
        Case 1
            If chkGridLines(1).Value = 0 Then
                flx2.GridLines = flexGridNone
            Else
                flx2.GridLines = flexGridFlat
            End If
    End Select

End Sub

Private Sub chkIncludeDev_Click()
    iIncludeDev = chkIncludeDev.Value
    Call PopGrid(DateValue(cmdSetDate.Tag))
End Sub

'Private Sub chkGridlines_Click()
'End Sub

Private Sub cmdSetDate_Click()
    PassDate = DateValue(cmdSetDate.Tag)
    With frmCal
        .PassLeft = Me.Left + ((Me.Width - Me.ScaleWidth) / 2) + cmdSetDate.Left
        .PassTop = Me.Top + (Me.Height - Me.ScaleHeight - 75) + cmdSetDate.Top _
                    + cmdSetDate.Height + 2730 '''+ frmCal.Height
        .Show 1
    End With
    
    Select Case iType
        Case 0
            If PassDate = Empty Then
                Me.Caption = "Usage View:  " & Format(Date, "dddd, mmmm d, yyyy")
                cmdSetDate.Tag = Format(Date, "mmmm d, yyyy")
            Else
                Me.Caption = "Usage View:  " & Format(PassDate, "dddd, mmmm d, yyyy")
                cmdSetDate.Tag = Format(PassDate, "mmmm d, yyyy")
            End If
            
            Call PopGrid(DateValue(cmdSetDate.Tag))
            
        Case 1
            If PassDate >= Date Then
                MsgBox "Date must be prior to today", vbExclamation, "Sorry..."
            Else
                Call Me.SetCountView(PassDate)
            End If
    End Select
End Sub

Private Sub flx2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < flx2.ColWidth(0) Then
        flx2.ToolTipText = "Click to view User List"
        flx2.MousePointer = flexCustom
    Else
        flx2.ToolTipText = ""
        flx2.MousePointer = flexDefault
    End If
End Sub

Private Sub flx2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < flx2.ColWidth(0) Then
        Call GetUserList(DateValue(Left(flx2.TextMatrix(flx2.RowSel, 0), 12)))
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, iHour As Integer, iCol As Integer, iCnt As Integer
    Dim sAMPM As String
    
    Screen.MousePointer = 11
    
    iType = 0
    CurrDate = Date - 60
    bPopped = False
    iIncludeDev = 0
    
    flx1.Cols = 4 + (60 * 24)
    flx1.ColWidth(0) = 0
    flx1.ColWidth(1) = 2400
    flx1.ColWidth(2) = 800
    flx1.ColWidth(3) = 800
    
    For i = 4 To flx1.Cols - 1
        flx1.ColWidth(i) = 15 ''30
        flx1.ColAlignment(i) = 4
    Next i
    
    Call PopGrid(Date)
    
    sAMPM = " AM"
    For i = 4 To flx1.Cols - 1
        iHour = Int((i - 4) / 60) ''+ 6
        If iHour > 12 Then sAMPM = " PM"
        If sAMPM = " PM" Then iHour = iHour - 12
        flx1.TextMatrix(0, i) = CStr(iHour) & ":00" & sAMPM
    Next i
    flx1.MergeRow(0) = True
    
    flx1.TextMatrix(0, 1) = "User": flx1.FixedAlignment(1) = 4
    flx1.TextMatrix(0, 2) = "IN": flx1.FixedAlignment(2) = 4
    flx1.TextMatrix(0, 3) = "OUT": flx1.FixedAlignment(3) = 4
    
    cmdSetDate.Tag = Format(Date, "mmmm d, yyyy")
    Me.Caption = "Usage View:  " & Format(Date, "dddd, mmmm d, yyyy")
    
    flx1.LeftCol = 484
    
    
'''    Call SetCountView(CurrDate)
    
'''    flx2.Cols = 52
'''    flx2.ColWidth(0) = 2400
'''    flx2.ColWidth(1) = 450
''''''    For i = 2 To flx2.Cols - 1
''''''        flx2.ColWidth(i) = 300
''''''        flx2.ColAlignment(i) = 4
''''''        flx2.TextMatrix(0, i) = i - 1
''''''    Next i
'''    flx2.Rows = DateValue(Date) - DateValue("01-AUG-03") + 1 + 1
'''    For i = 1 To flx2.Rows - 1
'''        flx2.TextMatrix(i, 0) = format(DateValue(Date - (i - 1)), "MMM D, YYYY (DDDD)")
'''        iCnt = GetColCount(format(DateValue(Date - (i - 1)), "dd-mmm-yy"))
'''        If iCnt + 2 > flx2.Cols Then flx2.Cols = iCnt + 2
'''        flx2.TextMatrix(i, 1) = iCnt
'''        flx2.Row = i
'''        For iCol = 2 To iCnt + 1
'''            flx2.Col = iCol
'''            flx2.CellBackColor = RGB(209, 225, 115) ''(255, 177, 0)
'''        Next iCol
'''        If iCnt > 0 Then
'''            flx2.Col = iCnt + 1: flx2.CellForeColor = vbBlue: flx2.Text = iCnt
'''        End If
'''    Next i
'''
    flx2.FixedAlignment(0) = 4: flx2.TextMatrix(0, 0) = "Date"
    flx2.FixedAlignment(1) = 4: flx2.TextMatrix(0, 1) = "Cnt"
'''    For i = 2 To flx2.Cols - 1
'''        flx2.ColWidth(i) = 300
'''        flx2.ColAlignment(i) = 4
'''        flx2.TextMatrix(0, i) = i - 1
'''    Next i
    
    Screen.MousePointer = 0
    
    
End Sub

Public Sub PopGrid(pDate As Date)
    Dim strSelect As String, sIncludeDev As String
    Dim rst As ADODB.Recordset
    Dim iRow As Integer, iCol As Integer
    Dim l1 As Long, l2 As Long
    
    flx1.Visible = False
    
    Select Case iIncludeDev
        Case 1: sIncludeDev = ""
        Case 0: sIncludeDev = "AND LL.ADDUSER NOT LIKE 'Steve Westerholm' "
    End Select
    
    iRow = 0
    flx1.Rows = 1
    strSelect = "SELECT LL.LOCKID, LL.LOCKOPENDTTM, LL.LOCKCLOSEDTTM, " & _
                "TO_CHAR(LL.LOCKOPENDTTM, 'SSSSS') AS OPEN, " & _
                "TO_CHAR(LL.LOCKCLOSEDTTM, 'SSSSS') AS CLOSE, " & _
                "(TRIM(U.NAME_FIRST)||' '||TRIM(NAME_LAST)) AS USERNAME, U.EMPLOYER " & _
                "FROM ANNOTATOR.ANO_LOCKLOG LL, IGLPROD.IGL_USER U " & _
                "WHERE LL.LOCKREFSOURCE = 'ANNO_OPEN' " & _
                "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YY') = '" & UCase(Format(pDate, "dd-mmm-yy")) & "' " & _
                sIncludeDev & _
                "AND LL.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "ORDER BY LL.LOCKOPENDTTM"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        iRow = iRow + 1
        flx1.Rows = iRow + 1
        flx1.TextMatrix(iRow, 0) = rst.Fields("LOCKID")
        flx1.TextMatrix(iRow, 1) = Trim(rst.Fields("USERNAME"))
        flx1.TextMatrix(iRow, 2) = Format(rst.Fields("LOCKOPENDTTM"), "HH:NN AMPM")
        flx1.TextMatrix(iRow, 3) = Format(rst.Fields("LOCKCLOSEDTTM"), "HH:NN AMPM")
        
        ''DO GANTT HERE''
        l1 = (rst.Fields("OPEN") / 60) + 4 '' - (6 * 60) + 4
        If l1 < 4 Then l1 = 4
        If Format(rst.Fields("LOCKOPENDTTM"), "DD-MMM-YY") = Format(rst.Fields("LOCKCLOSEDTTM"), "DD-MMM-YY") Then
            l2 = (rst.Fields("CLOSE") / 60) + 4 '' - (6 * 60) + 4
            If l2 >= flx1.Cols Then l2 = flx1.Cols - 1
        Else
            If Format(rst.Fields("LOCKOPENDTTM"), "DD-MMM-YY") = Format(Now, "DD-MMM-YY") Then
                l2 = (GetMinutes(Now)) + 4 '' - (6 * 60) + 4
                If l2 >= flx1.Cols Then l2 = flx1.Cols - 1
            Else
                l2 = l1 + 60
                If l2 >= flx1.Cols Then l2 = flx1.Cols - 1
            End If
        End If
        flx1.Row = iRow
        For iCol = l1 To l2
            flx1.Col = iCol
            If UCase(Left(rst.Fields("EMPLOYER"), 3)) = "GPJ" Then
                flx1.CellBackColor = RGB(209, 225, 115) ''RGB(255, 177, 0)
            Else
                flx1.CellBackColor = RGB(0, 113, 255)
            End If
        Next iCol
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    flx1.Visible = True
End Sub

Private Sub Form_Resize()
'''    Dim i As Integer
    If Me.WindowState = 1 Then Exit Sub
    flx1.Width = Me.ScaleWidth - (flx1.Left * 2)
    flx1.Height = Me.ScaleHeight - flx1.Top - flx1.Left
    
    flx2.Width = flx1.Width
    flx2.Height = flx1.Height
    
    fraOpt.Left = flx1.Left + flx1.Width - fraOpt.Width
End Sub

Public Function GetMinutes(pDate As Date) As Long
    Dim lHr As Long, lMin As Long
    lHr = CLng(Format(pDate, "h"))
    lMin = CLng(Format(pDate, "n"))
    GetMinutes = (lHr * 60) + lMin
End Function

Public Function GetColCount(pDate As String)
    Dim strSelect As String, sDate As String
    Dim rst As ADODB.Recordset
    
    strSelect = "select COUNT(distinct adduser) AS CNT " & _
                "From ANNOTATOR.ANO_LOCKLOG " & _
                "WHERE LOCKREFSOURCE = 'ANNO_OPEN' " & _
                "AND UPPER(to_char(lockopendttm, 'DD-MON-YY')) = '" & UCase(pDate) & "'"
    Set rst = Conn.Execute(strSelect)
    GetColCount = rst.Fields("CNT")
    rst.Close: Set rst = Nothing
End Function

Public Sub GetUserList(pDate As Date)
    Dim strSelect As String, sDate As String, sMess As String
    Dim rst As ADODB.Recordset
    Dim iCnt As Integer
    
    sDate = Format(pDate, "dd-mmm-yy")
    sMess = ""
    iCnt = 0
    strSelect = "select distinct adduser " & _
                "From ANNOTATOR.ANO_LOCKLOG " & _
                "WHERE LOCKREFSOURCE = 'ANNO_OPEN' " & _
                "AND UPPER(to_char(lockopendttm, 'DD-MON-YY')) = '" & UCase(sDate) & "'"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        iCnt = iCnt + 1
        sMess = sMess & Trim(rst.Fields("ADDUSER")) & vbNewLine
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    MsgBox sMess, vbInformation, "(" & iCnt & ") Users on " & Format(pDate, "dddd, mmmm d, yyyy")
End Sub

Private Sub opt1_Click(Index As Integer)
    iType = Index
    Select Case iType
        Case 0
            flx2.Visible = False
'            cmdSetDate.Enabled = True
        Case 1
            If Not bPopped Then Call SetCountView(CurrDate)
            flx2.Visible = True
'            cmdSetDate.Enabled = False
    End Select
End Sub

Public Sub SetCountView(pDate As Date)
    Dim i As Integer, iCol As Integer, iCnt As Integer
    
    Me.MousePointer = 11
    
    flx2.Visible = False
    flx2.Rows = 1
    flx2.Cols = 102 ''52
    flx2.ColWidth(0) = 2400
    flx2.ColWidth(1) = 450
'''    For i = 2 To flx2.Cols - 1
'''        flx2.ColWidth(i) = 300
'''        flx2.ColAlignment(i) = 4
'''        flx2.TextMatrix(0, i) = i - 1
'''    Next i
    flx2.Rows = DateValue(Date) - pDate + 1 + 1
    For i = 1 To flx2.Rows - 1
        flx2.TextMatrix(i, 0) = Format(DateValue(Date - (i - 1)), "MMM D, YYYY (DDDD)")
        iCnt = GetColCount(Format(DateValue(Date - (i - 1)), "dd-mmm-yy"))
        If iCnt + 2 > flx2.Cols Then flx2.Cols = iCnt + 2
        flx2.TextMatrix(i, 1) = iCnt
        flx2.Row = i
        For iCol = 2 To iCnt + 1
            flx2.Col = iCol
            flx2.CellBackColor = RGB(209, 225, 115) ''(255, 177, 0)
        Next iCol
        If iCnt > 0 Then
            flx2.Col = iCnt + 1: flx2.CellForeColor = vbBlue: flx2.Text = iCnt
        End If
    Next i
    
    
    For i = 2 To flx2.Cols - 1
        flx2.ColWidth(i) = 300
        flx2.ColAlignment(i) = 4
        flx2.TextMatrix(0, i) = i - 1
    Next i
    
    bPopped = True
    If iType = 1 Then flx2.Visible = True
    Me.MousePointer = 0
    
End Sub
