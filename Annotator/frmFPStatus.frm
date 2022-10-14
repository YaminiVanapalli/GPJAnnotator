VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmFPStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Floorplan Status Screen"
   ClientHeight    =   8625
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
   Icon            =   "frmFPStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optClients 
      Caption         =   "Show All Clients"
      Height          =   375
      Index           =   1
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   2115
   End
   Begin VB.OptionButton optClients 
      Caption         =   "Show Current Client Only"
      Height          =   375
      Index           =   0
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Value           =   -1  'True
      Width           =   2115
   End
   Begin VB.CheckBox chkShowReleased 
      Caption         =   "<Used in background>"
      Height          =   255
      Left            =   1980
      TabIndex        =   5
      Top             =   8250
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkGridLines 
      Caption         =   "Show Grid Lines"
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   7575
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13361
      _Version        =   393216
      Rows            =   1000
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   2
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
   Begin VB.Label lblPrint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hard Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      MouseIcon       =   "frmFPStatus.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Click to Export to Excel"
      Top             =   8220
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Floorplan Status Screen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9060
      MouseIcon       =   "frmFPStatus.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to Return to Annotator"
      Top             =   8220
      Width           =   2580
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   60
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4020
      TabIndex        =   8
      Top             =   8100
      Width           =   7875
   End
End
Attribute VB_Name = "frmFPStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public d_start As String
Public d_end As String
'''''Public ConnStr As String
''''''''Public LogInit As String
'''''Public BCC As String
'''''Public BCN As String
'''''Public FBCN As String
'''''Public CSHYR As Integer
'''''Public SHYR As Integer
'''''Public SHCD As Integer
'''''Public SHNM As String
Public SelClause As String
Public WhereClause As String
Public AndClause As String
Dim sCunoList As String
Dim tBCC As String
Dim tFBCN As String


Public Property Get PassBCC() As String
    PassBCC = tBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As String)
    tBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = tFBCN
End Property
Public Property Let PassFBCN(ByVal vNewValue As String)
    tFBCN = vNewValue
End Property




Private Sub chkShowReleased_Click()
    If chkShowReleased.Value = 0 Then
        AndClause = "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = KU.AN8_SHCD " & _
                    "AND SM.SHY56SHYR = KU.SHYR " & _
                    "AND CS.CSY56CUNO = KU.AN8_CUNO " & _
                    "AND CS.CSY56CUNO = CU.ABAN8 " & _
                    "AND KU.FPSTATUS > 0 " & _
                    "AND KU.FPSTATUS < 7 " & _
                    "ORDER BY SM.SHY56BEGDT, UPPER(SM.SHY56NAMA), UPPER(CU.ABALPH)"
    Else
        AndClause = "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                    "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                    "AND SM.SHY56SHCD = KU.AN8_SHCD " & _
                    "AND SM.SHY56SHYR = KU.SHYR " & _
                    "AND CS.CSY56CUNO = KU.AN8_CUNO " & _
                    "AND CS.CSY56CUNO = CU.ABAN8 " & _
                    "AND KU.FPSTATUS > 0 " & _
                    "ORDER BY SM.SHY56BEGDT, UPPER(SM.SHY56NAMA), UPPER(CU.ABALPH)"
    End If
End Sub



Private Sub lblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPrint.FontBold = False
    lblClose.FontBold = False
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LabelBorders lblClose
End Sub

Private Sub lblPrint_Click()
    lblPrint.MousePointer = 11
    Screen.MousePointer = 11
    StatusPrintOut
    Screen.MousePointer = 0
    lblPrint.MousePointer = 99
End Sub

Private Sub chkGridlines_Click()
    If chkGridLines.Value = 1 Then
        Flex1.GridLines = flexGridFlat
    Else
        Flex1.GridLines = flexGridNone
    End If
End Sub

Private Sub Form_Load()
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim i As Integer
    
    Screen.MousePointer = 11
    sCunoList = CLng(tBCC)
    For i = 0 To 9
        Select Case i
            Case 1: Flex1.ColAlignment(i) = 1
            Case Else: Flex1.ColAlignment(i) = 4
        End Select
    Next i
    FillHeader
    d_start = Format(Now, "DD-MMM-YYYY")
    d_end = Format(DateAdd("m", 3, Now), "DD-MMM-YYYY")
    lblPeriod.Caption = "Report Period:  " & Format(d_start, "d-mmm-yy") & " to " & Format(d_end, "d-mmm-yy")
       
    SelClause = "SELECT SM.SHY56SHCD, SM.SHY56NAMA, CS.CSY56CUNO, CU.ABALPH, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'DD-MON-YYYY')BEG_DATE, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'DD-MON-YYYY')END_DATE, " & _
                "FPS.FPSTATUS, FPS.FPSTATDT " & _
                "FROM " & F5601 & " SM, " & F5611 & " CS, " & _
                "" & F0101 & " CU, AQUA.AQUA_FLOORPLAN_STATUS FPS "
    WhereClause = "WHERE CS.CSY56CUNO = " & CLng(tBCC) & " " & _
                "AND SM.SHY56BEGDT > " & IGLToJDEDate(d_start) & " " & _
                "AND SM.SHY56BEGDT < " & IGLToJDEDate(d_end) & " "
    AndClause = "AND SM.SHY56SHCD = CS.CSY56SHCD " & _
                "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                "AND SM.SHY56SHCD = FPS.AN8_SHCD " & _
                "AND SM.SHY56SHYR = FPS.SHYR " & _
                "AND CS.CSY56CUNO = FPS.AN8_CUNO " & _
                "AND CS.CSY56CUNO = CU.ABAN8 " & _
                "AND FPS.FPSTATUS > 0 " & _
                "ORDER BY SM.SHY56BEGDT, UPPER(SM.SHY56NAMA), UPPER(CU.ABALPH)"
    FillGrid
    
    

    Screen.MousePointer = 0
End Sub


Public Function FillGrid()
    'Dim Conn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim iRow As Integer, nrow As Integer, ictr As Integer
    Dim ShoCode As Long
    Dim clnt As String, cuno As String, fpsd As String
    Dim fpst As Integer, iCol As Integer
    Dim t_shcd As String, t_nama As String, d_begd As String, d_endd As String
    
    Flex1.Visible = False
    Flex1.Rows = 2
    strSelect = SelClause & WhereClause & AndClause
    Debug.Print strSelect
    Set rst = Conn.Execute(strSelect)
    iRow = 1
    nrow = 1
    ictr = 0
    ShoCode = 0
    With rst
        Do While Not .EOF
            Flex1.Rows = Flex1.Rows + 1
            Flex1.Row = iRow + ictr
            If .Fields("SHY56SHCD") = ShoCode Then
                clnt = UCase(Trim(.Fields("ABALPH")))
                cuno = Right("00000000" & CStr(.Fields("CSY56CUNO")), 8)
                fpst = .Fields("FPSTATUS") + 1
                fpsd = Format(.Fields("FPSTATDT"), "d-mmm-yy")
                Flex1.Col = 1
                Flex1.Text = CLng(cuno) & " - " & clnt
                For iCol = 2 To fpst
                    Flex1.Col = iCol
                    Flex1.CellBackColor = vbActiveTitleBar
                Next
'''                Flex1.CellAlignment = 4
                Flex1.CellForeColor = vbTitleBarText
                Flex1.Text = fpsd
                .MoveNext
            Else
                ShoCode = .Fields("SHY56SHCD")
'''                t_shcd = Right("0000" & CStr(ShoCode), 4)
                t_nama = UCase(Trim(.Fields("SHY56NAMA")))
                d_begd = Format(DateValue(.Fields("BEG_DATE")), "d-mmm-yy")
'                d_endd = Format(.Fields("END_DATE"), "m/d/yy")
                With Flex1
                    .Col = 0
'''                    .CellAlignment = 4
                    .CellFontBold = True
                    .Text = CStr(ShoCode)
                    .Col = 1
                    .CellFontBold = True
                    .Text = t_nama
                    .Col = 2
'''                    .CellAlignment = 4
                    .Text = "Start date"
                    .Col = 3
'''                    .CellAlignment = 4
                    .Text = d_begd
'                    .Col = 4
'                    .CellAlignment = 4
'                    .Text = d_endd
                End With
            End If
            ictr = ictr + 1
        Loop
        .Close
    End With
    If ictr = 0 Then Flex1.Rows = 1 Else Flex1.Rows = ictr + 1
    Flex1.Visible = True
End Function


Public Function FillHeader()
    With Flex1
        If chkShowReleased.Value = 1 Then
            .ColWidth(0) = (.Width - 300) * 0.06 '0.07
            .ColWidth(1) = (.Width - 300) * 0.34  '0.39
            For i = 2 To 9
                .ColWidth(i) = (.Width - 300) * 0.075 '0.09
            Next i
        Else
            .ColWidth(0) = (.Width - 300) * 0.07 '0.07
            .ColWidth(1) = (.Width - 300) * 0.39 '0.39
            For i = 2 To 7
                .ColWidth(i) = (.Width - 300) * 0.09 '0.09
            Next i
            For i = 8 To 9
                .ColWidth(i) = 0
            Next i
        End If
        .RowHeight(0) = 500
        For i = 1 To 499
            .RowHeight(i) = 250
        Next i
        .Col = 0
        .Row = 0
'''        .CellAlignment = 4
        .Text = "Show Code"
        .Col = 2
'''        .CellAlignment = 4
        .Text = "Plan Req'd"
        .Col = 3
'''        .CellAlignment = 4
        .Text = "DWG Setup"
        .Col = 4
'''        .CellAlignment = 4
        .Text = "Bkgrd Drawn"
        .Col = 5
'''        .CellAlignment = 4
        .Text = "Prelim Layout"
        .Col = 6
'''        .CellAlignment = 4
        .Text = "A/E Apprvd"
        .Col = 7
'''        .CellAlignment = 4
        .Text = "DWG Comp"
        .Col = 8
'''        .CellAlignment = 4
        .Text = "DWG Release"
        .Col = 9
'''        .CellAlignment = 4
        .Text = "Revised Release"
    End With
End Function

Private Sub lblClose_Click()
    Unload Me
End Sub

Public Function StatusPrintOut()
'''''    'Dim Conn As New ADODB.Connection
'''''    Dim rst As ADODB.Recordset
'''''    Dim strSelect As String
'''''    Dim objApp As New Excel.Application
'''''    Dim objBook As Excel.Workbook
'''''    Dim objSheet As Excel.Worksheet
'''''    Dim CurrSC As Long
'''''    Dim iRow As Integer
'''''    Dim intSheet As Integer
'''''    Dim EndCell As String
'''''    Dim ExcelOpen As Boolean
'''''
'''''    On Error GoTo ErrorTrap
'''''    '''Set objApp = New Excel.Application
'''''    Set objApp = CreateObject("Excel.Application")
'''''    ExcelOpen = True
''''''    objApp.Visible = True
'''''    Set objBook = objApp.Workbooks.Add
'''''
'''''    objApp.DisplayAlerts = False
'''''    For intSheet = objBook.Worksheets.Count To 2 Step -1
'''''        objBook.Worksheets(intSheet).Delete
'''''    Next
'''''    objApp.DisplayAlerts = True
'''''
'''''    Set objSheet = objBook.Worksheets(1)
'''''    objSheet.Name = "FPStatus"
'''''
'''''    With objSheet
'''''        .Columns("A:A").ColumnWidth = 7
'''''        .Columns("B:B").ColumnWidth = 42
'''''        If chkShowReleased.value = 0 Then
'''''            .Columns("C:H").ColumnWidth = 13.3
'''''        Else
'''''            .Columns("C:J").ColumnWidth = 10
'''''        End If
'''''    End With
'''''    If chkShowReleased.value = 0 Then
'''''        objSheet.Range("A1:F1").Merge
'''''        With objSheet.Range("A1:F1")
''''''''            .merge
'''''            .Font.Size = 20
'''''            .Font.Bold = True
'''''            .Font.Shadow = True
'''''            .HorizontalAlignment = xlLeft
'''''            .FormulaR1C1 = "Floor Plan Status Report: " & UCase(Format(d_start, "d-mmm-yy")) & " to " & _
'''''                        UCase(Format(d_end, "d-mmm-yy"))
'''''        End With
'''''    Else
'''''        objSheet.Range("A1:G1").Merge
'''''        With objSheet.Range("A1:G1")
''''''''            .merge
'''''            .Font.Size = 20
'''''            .Font.Bold = True
'''''            .Font.Shadow = True
'''''            .HorizontalAlignment = xlLeft
'''''            .FormulaR1C1 = "Floor Plan Status Report: " & UCase(Format(d_start, "d-mmm-yy")) & " to " & _
'''''                        UCase(Format(d_end, "d-mmm-yy"))
'''''        End With
'''''    End If
'''''    If chkShowReleased.value = 0 Then
'''''        With objSheet.Range("G1:H1")
'''''            .Merge
'''''            .Font.Size = 10
'''''            .Font.Bold = True
'''''            .Font.Shadow = True
'''''            .HorizontalAlignment = xlCenter
'''''            .WrapText = True
'''''            .FormulaR1C1 = "List Displays Current Status of all non-Released Floor Plans"
'''''        End With
'''''    Else
'''''        With objSheet.Range("H1:J1")
'''''            .Merge
'''''            .Font.Size = 10
'''''            .Font.Bold = True
'''''            .Font.Shadow = True
'''''            .HorizontalAlignment = xlCenter
'''''            .WrapText = True
'''''            .FormulaR1C1 = "List Displays Current Status of all Requested Floor Plans"
'''''        End With
'''''    End If
'''''    With objSheet
'''''        .Cells(1, 1).Font.Bold = True
'''''        .Cells(1, 1).Font.Size = 20
'''''        .Rows("2:2").RowHeight = 25.5
'''''        .Rows("2:2").WrapText = True
'''''        .Rows("2:2").HorizontalAlignment = xlCenter
'''''        .Cells(2, 1).Formula = "Show Code"
'''''        .Cells(2, 2).Formula = "Show / Client"
'''''        .Cells(2, 3).Formula = "Plan Request"
'''''        .Cells(2, 4).Formula = "DWG Setup"
'''''        .Cells(2, 5).Formula = "Backgrd Drawn"
'''''        .Cells(2, 6).Formula = "Prelim Layout"
'''''        .Cells(2, 7).Formula = "A/E Approved"
'''''        .Cells(2, 8).Formula = "DWG Complete"
'''''        If chkShowReleased.value = 1 Then
'''''            .Cells(2, 9).Formula = "DWG Release"
'''''            .Cells(2, 10).Formula = "Rev'd Release"
'''''        End If
'''''        For i = 1 To 10
'''''            .Cells(2, i).Font.Bold = True
'''''        Next
'''''        If chkShowReleased.value = 0 Then
'''''            .Range("A2:H2").Interior.ColorIndex = 15
'''''            .Range("A2:H2").Interior.Pattern = xlSolid
'''''        Else
'''''            .Range("A2:J2").Interior.ColorIndex = 15
'''''            .Range("A2:J2").Interior.Pattern = xlSolid
'''''        End If
'''''        .Columns("A:A").NumberFormat = "@"
'''''    End With
'''''
'''''    'Conn.Open (ConnStr)
'''''    'ConnOpen = True
'''''    strSelect = SelClause & WhereClause & AndClause
'''''    Set rst = Conn.Execute(strSelect)
'''''    CurrSC = 0
'''''    iRow = 3
'''''    With rst
'''''        Do While Not .EOF
'''''            If .Fields("SHY56SHCD") <> CurrSC Then
'''''                objSheet.Cells(iRow, 1).Formula = CStr(.Fields("SHY56SHCD"))
'''''                objSheet.Cells(iRow, 1).Font.Bold = True
'''''                CurrSC = .Fields("SHY56SHCD")
'''''                objSheet.Cells(iRow, 2).Formula = UCase(Trim(.Fields("SHY56NAMA")))
'''''                objSheet.Cells(iRow, 2).Font.Bold = True
'''''                objSheet.Cells(iRow, 3).Formula = "Start Date"
'''''                objSheet.Cells(iRow, 4).Formula = Format(DateValue(.Fields("BEG_DATE")), "d-mmm-yy")
'''''                iRow = iRow + 1
'''''            End If
'''''            objSheet.Cells(iRow, 2).Formula = CStr(.Fields("CSY56CUNO")) & " - " & UCase(Trim(.Fields("ABALPH")))
'''''            objSheet.Cells(iRow, (2 + .Fields("FPSTATUS"))) = Format(.Fields("FPSTATDT"), "d-mmm-yy")
'''''            Select Case .Fields("FPSTATUS")
'''''            Case 1
'''''                EndCell = "C"
'''''            Case 2
'''''                EndCell = "D"
'''''            Case 3
'''''                EndCell = "E"
'''''            Case 4
'''''                EndCell = "F"
'''''            Case 5
'''''                EndCell = "G"
'''''            Case 6
'''''                EndCell = "H"
'''''            Case 7
'''''                EndCell = "I"
'''''            Case 8
'''''                EndCell = "J"
'''''            End Select
'''''            objSheet.Range("C" & CStr(iRow) & ":" & EndCell & CStr(iRow)).Interior.ColorIndex = 15
'''''            objSheet.Range("C" & CStr(iRow) & ":" & EndCell & CStr(iRow)).Interior.Pattern = xlSolid
'''''            iRow = iRow + 1
'''''            .MoveNext
'''''        Loop
'''''    End With
'''''    rst.Close
'''''
'''''    strRange = "C2:H" & CStr(iRow - 1)
'''''    objSheet.Range(strRange).HorizontalAlignment = xlCenter
'''''    objSheet.Range("A1").Select
'''''
'''''    With objSheet.PageSetup
'''''        .PrintTitleRows = "$2:$2"
'''''        .PrintTitleColumns = ""
'''''    End With
'''''    objSheet.PageSetup.PrintArea = ""
'''''    With objSheet.PageSetup
'''''        .LeftFooter = "&D"
'''''        .RightFooter = "Page &P of &N"
'''''        .LeftMargin = objApp.InchesToPoints(0.25)
'''''        .RightMargin = objApp.InchesToPoints(0.25)
'''''        .TopMargin = objApp.InchesToPoints(0.7)
'''''        .BottomMargin = objApp.InchesToPoints(0.7)
'''''        .HeaderMargin = objApp.InchesToPoints(0.5)
'''''        .FooterMargin = objApp.InchesToPoints(0.5)
'''''        .PrintGridlines = True
'''''        .CenterHorizontally = True
'''''        .Orientation = xlLandscape
'''''        .PaperSize = xlPaperLetter
'''''        .Zoom = 100
'''''    End With
'''''    objApp.Visible = True
'''''Exit Function
'''''
'''''ErrorTrap:
'''''    MsgBox "Error encountered:" & vbCr & vbCr & Err.Description, vbCritical, "Oh,oh..."
'''''    If ExcelOpen = True Then
'''''        objBook.Close SaveChanges:=False
'''''        objApp.Quit
'''''        Set objApp = Nothing
'''''        Set objBook = Nothing
'''''        Set objSheet = Nothing
'''''    End If
End Function

Public Sub CheckInteger(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = vbKeyBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub lblPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LabelBorders lblPrint
End Sub

Private Sub optClients_Click(Index As Integer)
    Screen.MousePointer = 11
    Select Case Index
        Case 0
            If InStr(1, sCunoList, ",") > 0 Then
                sCunoList = CStr(tBCC)
                WhereClause = "WHERE CS.CSY56CUNO = " & CLng(tBCC) & " " & _
                            "AND SM.SHY56BEGDT > " & IGLToJDEDate(d_start) & " " & _
                            "AND SM.SHY56BEGDT < " & IGLToJDEDate(d_end) & " "
                FillGrid
            End If
        Case 1
            If InStr(1, sCunoList, ",") = 0 Then
                sCunoList = CunoList
                WhereClause = "WHERE CS.CSY56CUNO IN (" & sCunoList & ") " & _
                            "AND SM.SHY56BEGDT > " & IGLToJDEDate(d_start) & " " & _
                            "AND SM.SHY56BEGDT < " & IGLToJDEDate(d_end) & " "
                FillGrid
            End If
    End Select
    Screen.MousePointer = 0
End Sub

Public Sub LabelBorders(lbl1 As Label)
    lblPrint.FontBold = False
    lblClose.FontBold = False
    lbl1.FontBold = True
End Sub
