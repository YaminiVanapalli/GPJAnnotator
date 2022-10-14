VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Space Plan Drawing Log"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHours 
      Caption         =   "Drawing Hours"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid flxTime 
         Height          =   2595
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4577
         _Version        =   393216
         Rows            =   8
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Entries:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   300
         UseMnemonic     =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCheck 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCheck 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Eng"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2820
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   345
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entry"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3900
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   450
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraLog 
      Caption         =   "Drawing Log"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid flxLog 
         Height          =   2595
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4577
         _Version        =   393216
         Rows            =   8
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
      End
      Begin VB.Label lblCheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entry"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3900
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   450
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCheck 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Eng"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   345
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCheck 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   405
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Log Entries:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   300
         UseMnemonic     =   0   'False
         Width           =   855
      End
   End
   Begin VB.Menu mnuLogMenu 
      Caption         =   "Log Menu"
      Begin VB.Menu mnuDataRefresh 
         Caption         =   "Refresh with Active DWG"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintLog 
         Caption         =   "Printable View..."
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseLog 
         Caption         =   "Close Log Interface"
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tDWGID As Long
Dim sHdr As String
Dim bChecking As Boolean
Dim iPriority As Integer

Public Property Get PassDWGID() As Long
    PassDWGID = tDWGID
End Property
Public Property Let PassDWGID(ByVal vNewValue As Long)
    tDWGID = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = sHdr
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    sHdr = vNewValue
End Property





Private Sub Form_Load()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim iRow As Integer, iCol As Integer
    
    lblShow.Caption = "LOG: " & sHdr
    
    With flxLog
        .Row = 0
        .ColWidth(0) = 0
        .Col = 1: .ColWidth(1) = (.Width - 215) * 0.15: .ColAlignment(1) = 3: .Text = "Date"
        lblCheck(1).Left = .Left + .CellLeft: lblCheck(1).Width = .ColWidth(1)
        .Col = 2: .ColWidth(2) = (.Width - 215) * 0.1: .ColAlignment(2) = 3: .Text = "FPS"
        lblCheck(2).Left = .Left + .CellLeft: lblCheck(2).Width = .ColWidth(2)
        .Col = 3: .ColWidth(3) = (.Width - 215) * 0.75: .ColAlignment(3) = 0: .CellAlignment = 4: .Text = "Log Entry"
        lblCheck(3).Left = .Left + .CellLeft: lblCheck(3).Width = .ColWidth(3)
        .ColWidth(4) = 250
    End With
    
'''    strSelect = "SELECT EP.PRGNAME, EE.ELEMNAME " & _
'''                "FROM " & ENGElt & " EE, " & ENGProg & " EP " & _
'''                "WHERE EE.ELEMID = " & tElemID & " " & _
'''                "AND EE.PRGID = EP.PRGID"
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        Me.Caption = "LOG: " & Trim(rst.Fields("PRGNAME")) & " -- " & Trim(rst.Fields("ELEMNAME"))
'''    Else
'''        Me.Caption = "LOG: " & CStr(tElemID)
'''    End If
'''    rst.Close: Set rst = Nothing
    
    ''GET EXISTING LOG ENTRIES''
    Call GetLog(tDWGID)
    
End Sub

Private Sub mnuCloseLog_Click()
    Unload Me
End Sub

'''Private Sub mnuDataRefresh_Click()
'''    Dim sType As String
'''    Dim AcadApp As AcadApplication
'''    Dim AcadDoc As AcadDocument
'''    Dim TrackingDictionary As AcadDictionary, TrackingXRecord As AcadXRecord
'''    Dim XRecordDataType As Variant, XRecordData As Variant
'''    Dim tElemID As Long, tShtID As Long
'''    Dim strSelect As String
'''    Dim rst As ADODB.Recordset
'''
'''    On Error Resume Next
'''    Set AcadApp = GetObject(, "AutoCAD.Application")
'''    If Err Then
'''        MsgBox "AutoCAD not currently running.", vbExclamation, "Sorry..."
'''        Exit Sub
'''    End If
'''
'''    Set AcadDoc = AcadApp.ActiveDocument
'''    If Err Then
'''        MsgBox "There is no active drawing.", vbExclamation, "Sorry..."
'''        Exit Sub
'''    End If
'''
'''
'''
'''    Set TrackingDictionary = AcadDoc.Dictionaries("GPJ")
'''    If Err Then
'''        MsgBox "The GPJ Dictionary has not been activated in this drawing.", vbExclamation, "Sorry..."
'''        Exit Sub
'''    End If
'''
'''    Set TrackingXRecord = TrackingDictionary.GetObject("DB_DWG_TYPE")
'''    If Err Then
'''        MsgBox "This drawing has not been setup using the Engineering App.", vbExclamation, "Sorry..."
'''        Exit Sub
'''    End If
'''
'''    ' Get current XRecordData
'''    TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
'''    sType = CStr(XRecordData(0))
'''
'''    Select Case sType
'''        Case "FAB"
'''            MsgBox "The Log Entry Interface is Element Level.  You are currently in a FAB drawing, " & _
'''                        "which is Program Level, not Element Level.", vbExclamation, "Sorry..."
'''            Exit Sub
'''        Case "SHT", "DTL", "MOD"
'''            Set TrackingXRecord = TrackingDictionary.GetObject("DB_PROGID")
'''            TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
'''            tProgID = CLng(XRecordData(0))
'''            Set TrackingXRecord = TrackingDictionary.GetObject("DB_ELEMID")
'''            TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
'''            tElemID = CLng(XRecordData(0))
'''            Set TrackingXRecord = TrackingDictionary.GetObject("DB_SHTID")
'''            TrackingXRecord.GetXRecordData XRecordDataType, XRecordData
'''            tShtID = CLng(XRecordData(0))
'''
'''            strSelect = "SELECT EP.PRGNAME, EE.ELEMNAME " & _
'''                        "FROM " & ENGElt & " EE, " & ENGProg & " EP " & _
'''                        "WHERE EE.ELEMID = " & tElemID & " " & _
'''                        "AND EE.PRGID = EP.PRGID"
'''            Set rst = Conn.Execute(strSelect)
'''            If Not rst.EOF Then
'''                Me.Caption = "LOG: " & Trim(rst.Fields("PRGNAME")) & " -- " & Trim(rst.Fields("ELEMNAME"))
'''            Else
'''                Me.Caption = "LOG: " & CStr(tElemID)
'''            End If
'''            rst.Close: Set rst = Nothing
'''
'''            Call GetLog(tElemID)
'''
'''    End Select
'''End Sub


'''Private Sub mnuSelect_Click()
'''    With frmPickElement
'''        .PassELEMID = tElemID
'''        .PassFORM = Me.Name
'''        .Show 1
'''    End With
'''End Sub

Public Sub GetLog(tDID As Long)
    Dim iRow As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    iRow = 1
    flxLog.Rows = iRow
    strSelect = "SELECT * FROM " & EngLog & " " & _
                "WHERE DWGID = " & tDID & " " & _
                "ORDER BY LOGID DESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        flxLog.Rows = iRow + 1
        flxLog.TextMatrix(iRow, 0) = rst.Fields("LOGID")
        flxLog.TextMatrix(iRow, 1) = Format(rst.Fields("ADDDTTM"), "M/D/YY H:MM AMPM")
        flxLog.TextMatrix(iRow, 2) = Trim(rst.Fields("ADDUSER"))
        lblCheck(2).Caption = Trim(rst.Fields("ADDUSER"))
        flxLog.TextMatrix(iRow, 3) = Trim(rst.Fields("LOGENTRY"))
        lblCheck(3).Caption = Trim(rst.Fields("LOGENTRY"))
        If lblCheck(3).Height > lblCheck(2).Height Then
            flxLog.RowHeight(iRow) = lblCheck(3).Height
        Else
            flxLog.RowHeight(iRow) = lblCheck(2).Height
        End If
        iRow = iRow + 1
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub

Private Sub mnuPrintLog_Click()
    frmReport.PassHDR = sHdr
    frmReport.Show 1
End Sub

Public Function ConvertToHTML() As String
    Dim iRow As Integer, iCol As Integer
    
    Dim sHTML As String, strHTMLPath As String, tFile1 As String
    Dim i As Integer
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f3O As String, f3C As String, f2O As String, f2C As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tbl2O As String, tbl3O As String, tblC As String
    Dim trO_A As String, trO_B As String, trCO As String, trC As String, trMO As String, trO As String
    Dim hr As String, br As String, sp As String
    
    Dim th_A As String, th_B As String, th_C As String
    Dim td_A As String, td_B As String, td_C As String
    Dim td_A2 As String, td_B2 As String, td_C2 As String
    Dim thcO_A As String, thcO_B As String, thcO_C As String
    Dim sTab As String, sColor As String, sClient As String
    
    strHTMLPath = App.Path & "\" & Me.Caption & " - " & Format(Now, "dd-mmm-yy") & ".htm"
    sColor = "#FFFFFF"
    
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""middle"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trO_A = "<TR VALIGN=""top"" BGCOLOR=": trO_B = ">": trC = "</TR>"
    hr = "<HR>": br = "<BR>": sp = "&nbsp;"
    
    
    th_A = "<TH WIDTH="""
    th_B = "><FONT SIZE=2 FACE=""Times New Roman""><I><B>"
    th_C = "</B></I></FONT></TH>"
    thcO_A = "<TH WIDTH=""": thcO_B = """ ALIGN=""right"" COLSPAN=": thcO_C = "><FONT SIZE=2 FACE=""Times New Roman""><B><I>"

    td_A = "<TD WIDTH="""
    td_B = "><FONT SIZE=2 FACE=""Tahoma"">"
    td_C = "</FONT></TD>"
    
    td_A2 = "<TD WIDTH="""
    td_B2 = "><FONT SIZE=2 FACE=""Tahoma"">"
    td_C2 = "</FONT></TD>"
    
    thcO_A = "<TH WIDTH=""": thcO_B = """ ALIGN=""center"" COLSPAN=": thcO_C = "><FONT SIZE=2 FACE=""Times New Roman""><B><I>"
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY BGCOLOR=""#FFFFFF"">": bodC = "</BODY>"
'''    f4O = "<FONT SIZE=4 FACE=""Times New Roman""><B><I>": f4C = "</I></B></FONT>"
    f3O = "<FONT SIZE=3 FACE=""Tahoma""><B>": f3C = "</B></FONT>"
    f2O = "<FONT SIZE=2 FACE=""Tahoma"">": f2C = "</FONT>"
    
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
'''    trO_A = "<TR VALIGN=""top"" BGCOLOR=": trO_B = ">": trC = "</TR>"
'''    trO = "<TR VALIGN=""top"" height=""30"">": trC = "</TR>"
    
    hr = "<HR>": br = "<BR>": sp = "&nbsp;"
    
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & Me.Caption & " - " & Format(Now, "dd-mmm-yyyy") & tiC & hdC & vbNewLine

    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & Me.Caption & br & f3C & vbNewLine
    sHTML = sHTML & f2O & "Print Date:  " & Format(Now, "dd-mmm-yyyy") & f2C & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & trO_A & sColor & trO_B & vbNewLine
    sHTML = sHTML & th_A & "20%"" ALIGN=CENTER" & th_B & "Date" & th_C & vbNewLine
    sHTML = sHTML & th_A & "10%"" ALIGN=CENTER" & th_B & "FPS" & th_C & vbNewLine
    sHTML = sHTML & th_A & "70%"" ALIGN=CENTER" & th_B & "Comment" & th_C & vbNewLine
    
    sHTML = sHTML & trC & vbNewLine
    sHTML = sHTML & tblC & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & tblO & vbNewLine
    
    
    sClient = ""
    For iRow = frmLog.flxLog.Rows - 1 To 1 Step -1
        sHTML = sHTML & trO & vbNewLine
        For iCol = 1 To 3
            Select Case iCol
                Case 1
                    sHTML = sHTML & td_A & "20%"" ALIGN=CENTER" & td_B & frmLog.flxLog.TextMatrix(iRow, iCol) & td_C & vbNewLine
                Case 2
                    sHTML = sHTML & td_A & "10%"" ALIGN=CENTER" & td_B & frmLog.flxLog.TextMatrix(iRow, iCol) & td_C & vbNewLine
                Case 3
                    sHTML = sHTML & td_A & "70%"" ALIGN=LEFT" & td_B & frmLog.flxLog.TextMatrix(iRow, iCol) & td_C & vbNewLine
            End Select
'''            MsgBox "Row " & iRow & " : Col " & iCol & vbNewLine & vbNewLine & flx1.TextMatrix(iRow, iCol)
        Next iCol
        sHTML = sHTML & trC & vbNewLine
    Next iRow
    
    
    sHTML = sHTML & tblC & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    ConvertToHTML = tFile1
''    frmLog.flxLog.Rows = 2
    
''    web1.Navigate tFile1
''    web1.Visible = True
End Function

