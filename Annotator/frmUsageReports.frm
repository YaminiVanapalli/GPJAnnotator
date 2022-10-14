VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmUsageReports 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Usage Reports..."
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsageReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHdr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      Picture         =   "frmUsageReports.frx":020A
      ScaleHeight     =   1440
      ScaleWidth      =   11370
      TabIndex        =   6
      Top             =   0
      Width           =   11370
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Resort by User Type"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.CommandButton cmdUserReport 
      Caption         =   "View User Report..."
      Height          =   435
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdUsageGantt 
      Caption         =   "View Usage Gantt..."
      Height          =   435
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   3735
      Left            =   180
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   10995
      ExtentX         =   19394
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ComboBox cboViewUsage 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1583
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report:"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   1523
      Width           =   795
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpHDR 
      BackColor       =   &H00666666&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00666666&
      Height          =   600
      Left            =   0
      Top             =   1440
      Width           =   11370
   End
End
Attribute VB_Name = "frmUsageReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMinWidth As Long
Dim iUserReportView As Integer
Dim bDoneLoading As Boolean


Private Sub cboViewUsage_Click()
    Dim strSelect As String
    Dim sSource As String, sName As String, sMess As String, sBody As String, sQty As String, _
                sShow As String, sNewShow As String, sType As String, sTitle As String
    Dim rst As ADODB.Recordset
    Dim iOff As Integer, iCase As Integer, iLen As Integer, iQty As Integer
    Dim sUType As String, sHTML As String, tFile1 As String ''', sMess As String
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String, f2bO As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, _
                tdOa As String, tdObl As String, tdOb As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    Dim dl As String, dlC As String, dt As String, dtC As String
    Dim divO As String, divC As String
    Dim iUserCnt As Integer
    
    
    
    
    If cboViewUsage.Text = "" Then Exit Sub
    
    Me.MousePointer = 11
    
    
    
    iQty = 0
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY LINK=""black"" VLINK=""black"" ALINK=""blue"">": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    f2bO = "<FONT SIZE=2 COLOR=""000080"" FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 ALIGN=""CENTER"" VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=2><DIV align=left><FONT SIZE=2 FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdObl = "%"" ALIGN=left VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    tdOa = "<TD WIDTH=""": tdOb = "%"" VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    dl = "<DL>": dlC = "</DL>": dt = "<DT>": dtC = "</DT>"
    divO = "<DIV ALIGN=""RIGHT"">": divC = "</DIV>"
    
    
    
    iCase = cboViewUsage.ItemData(cboViewUsage.ListIndex)
'''    MsgBox iCase
    Select Case iCase
        Case 10: sTitle = "USAGE LOG: Floorplan Users during past 7 Days"
        Case 11: sTitle = "USAGE LOG: Graphics Users during past 7 Days"
        Case 16: sTitle = "POSTING LOG: Quantity of Graphics posted during past 7 Days"
        Case 17: sTitle = "POSTING LOG: Quantity of Graphics posted during past 30 Days"
        Case 22, 23, 24, 25
            iOff = iCase - 22
            sTitle = "DIGITAL IMAGE LIBRARY USAGE LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
        Case 26
            iOff = 7
            sTitle = "DIGITAL IMAGE LIBRARY  USAGE LOG: Users accessing DIL during past 7 Days"
        Case 27
            iOff = 30
            If iCase = 26 Then iOff = 7 Else iOff = 30
            sTitle = "DIGITAL IMAGE LIBRARY  USAGE LOG: Users accessing DIL during past 30 Days"
'            iOff = 365
'            If iCase = 26 Then iOff = 7 Else iOff = 365
'            sTitle = "ANNOTATOR USAGE LOG: Users accessing Annotator during past 365 Days"
        Case 18, 19, 20, 21
            'If iCase < 21 Then
                iOff = iCase - 18
                sTitle = "ANNOTATOR USER LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
            'Else
                'iOff = 365
                'sTitle = "ANNOTATOR USER LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
            'End If
        Case 4, 5, 6, 7
            iOff = iCase - 4
            sTitle = "GRAPHICS USER LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
        Case 12, 13, 14, 15
            iOff = iCase - 12
            sTitle = "GRAPHIC POSTING LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
        Case Else
            iOff = iCase
            sTitle = "FLOORPLAN USER LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
    End Select
    
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & sTitle & tiC & hdC & vbNewLine
    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & bolO & sTitle & bolC & fC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    
    
    
    
    
    
    Select Case iCase
        Case 10
            strSelect = "SELECT " & _
                        "LL.UPDUSER, C.ABALPH AS CLIENT, DS.SHYR, S.ABALPH AS SHOW " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, ANNOTATOR.DWG_SHOW DS, " & F0101 & " C, " & F0101 & " S " & _
                        "WHERE LOCKREFSOURCE = 'DWG_MASTER' " & _
                        "AND LL.LOCKOPENDTTM >= SYSDATE-7 " & _
                        "AND LL.LOCKREFID = DS.DWGID " & _
                        "AND DS.AN8_CUNO = C.ABAN8 " & _
                        "AND DS.AN8_SHCD = S.ABAN8 " & _
                        "ORDER BY UPDUSER, CLIENT, SHYR, SHOW"
            sName = "": sMess = "": sShow = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sName <> UCase(Trim(rst.Fields("UPDUSER"))) Then
                    iQty = iQty + 1
                    sName = UCase(Trim(rst.Fields("UPDUSER")))
                    sBody = sBody & trO & tdNO & f3O & bolO & sName & bolC & fC & tdNC & trC & vbNewLine
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  " & _
                            CStr(rst.Fields("SHYR")) & " - " & _
                            Trim(rst.Fields("SHOW"))
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sBody = sBody & trO & vbNewLine
                    sBody = sBody & tdOa & "5" & tdOb & tdC & vbNewLine
                    sBody = sBody & tdOa & "95" & tdOb & sShow & tdC & vbNewLine
                    sBody = sBody & trC & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            sQty = trO & tdNO & f3O & bolO & "Total Floorplan Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
        Case 11
            strSelect = "SELECT " & _
                        "LL.UPDUSER, C.ABALPH AS CLIENT, GM.GDESC " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, " & F0101 & " C, ANNOTATOR.GFX_MASTER GM " & _
                        "WHERE LOCKREFSOURCE = 'GFX_MASTER' " & _
                        "AND LL.LOCKOPENDTTM >= SYSDATE-7 " & _
                        "AND LL.LOCKREFID = GM.GID " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY UPDUSER, CLIENT, GDESC"
            sName = "": sMess = "": sShow = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sName <> UCase(Trim(rst.Fields("UPDUSER"))) Then
                    iQty = iQty + 1
                    sName = UCase(Trim(rst.Fields("UPDUSER")))
                    sBody = sBody & trO & tdNO & f3O & bolO & sName & bolC & fC & tdNC & trC & vbNewLine
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  [ " & _
                            Trim(rst.Fields("GDESC")) & " ]"
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sBody = sBody & trO & vbNewLine
                    sBody = sBody & tdOa & "5" & tdOb & tdC & vbNewLine
                    sBody = sBody & tdOa & "95" & tdOb & sShow & tdC & vbNewLine
                    sBody = sBody & trC & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            sQty = trO & tdNO & f3O & bolO & "Total Graphics Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
        Case 12, 13, 14, 15
'            iOff = iCase - 12
            strSelect = "SELECT ADDUSER, COUNT(GID) AS GCOUNT " & _
                        "From ANNOTATOR.GFX_MASTER " & _
                        "WHERE TO_CHAR(ADDDTTM, 'DD-MON-YY') = '" & _
                        UCase(Format(DateAdd("d", (iCase - 12) * -1, Date), "DD-MMM-YY")) & "' " & _
                        "GROUP BY ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                iQty = iQty + rst.Fields("GCOUNT")
                sBody = sBody & trO & vbNewLine
                sBody = sBody & tdOa & "8" & tdOb & rst.Fields("GCOUNT") & tdC & vbNewLine
                sBody = sBody & tdOa & "92" & tdOb & Trim(rst.Fields("ADDUSER")) & tdC & vbNewLine
                sBody = sBody & trC & vbNewLine
                sMess = sMess & vbTab & rst.Fields("GCOUNT") & vbTab & Trim(rst.Fields("ADDUSER")) & vbNewLine
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            If sMess = "" Then
                sMess = "THERE WERE NO GRAPHICS POSTED"
            Else
                sMess = "THE FOLLOWING QUANTITIES OF GRAPHICS WERE POSTED:" & vbNewLine & sMess
            End If
            
        Case 16, 17
            If iCase = 16 Then iLen = -7 Else iLen = -30
            strSelect = "SELECT ADDUSER, COUNT(GID) AS GCOUNT " & _
                        "From ANNOTATOR.GFX_MASTER " & _
                        "WHERE ADDDTTM BETWEEN " & _
                        "TO_DATE('" & UCase(Format(DateAdd("d", iLen, Date), "DD-MMM-YY")) & "', 'DD-MON-YY') " & _
                        "AND TO_DATE('" & UCase(Format(Date, "DD-MMM-YY")) & "', 'DD-MON-YY') " & _
                        "GROUP BY ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                iQty = iQty + rst.Fields("GCOUNT")
                sBody = sBody & trO & vbNewLine
                sBody = sBody & tdOa & "8" & tdOb & rst.Fields("GCOUNT") & tdC & vbNewLine
                sBody = sBody & tdOa & "92" & tdOb & Trim(rst.Fields("ADDUSER")) & tdC & vbNewLine
                sBody = sBody & trC & vbNewLine
                sMess = sMess & vbTab & rst.Fields("GCOUNT") & vbTab & Trim(rst.Fields("ADDUSER")) & vbNewLine
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            If sMess = "" Then
                sMess = "THERE WERE NO GRAPHICS POSTED"
            Else
                sMess = "THE FOLLOWING QUANTITIES OF GRAPHICS WERE POSTED:" & vbNewLine & sMess
            End If
                
            
        Case 18, 19, 20, 21
'            iOff = iCase - 18
            strSelect = "SELECT DISTINCT LL.UPDUSER, AU.USERTYPE " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, IGLPROD.IGL_USER_APP_R US, ANNOTATOR.ANO_USERTYPE AU " & _
                        "WHERE TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "AND LL.LOCKREFSOURCE = 'ANNO_OPEN' " & _
                        "AND LL.USER_SEQ_ID = US.USER_SEQ_ID " & _
                        "AND US.APP_ID = 1002 " & _
                        "AND US.USER_PERMISSION_ID = AU.USERTYPEID " & _
                        "ORDER BY AU.USERTYPE, LL.UPDUSER"
            sMess = "": sType = "": sBody = "": sQty = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sType <> UCase(Trim(rst.Fields("USERTYPE"))) Then
                    sType = UCase(Trim(rst.Fields("USERTYPE")))
                    sBody = sBody & trO & tdNO & f3O & bolO & sType & bolC & fC & tdNC & trC & vbNewLine
'''                    If sMess = "" Then
'''                        sMess = sType & vbNewLine
'''                    Else
'''                        sMess = sMess & vbNewLine & sType & vbNewLine
'''                    End If
                    sShow = ""
                End If
                iQty = iQty + 1
'''                sMess = sMess & vbTab & UCase(Trim(rst.Fields("UPDUSER"))) & vbNewLine
                sBody = sBody & trO & vbNewLine
                sBody = sBody & tdOa & "5" & tdOb & tdC & vbNewLine
                sBody = sBody & tdOa & "95" & tdOb & UCase(Trim(rst.Fields("UPDUSER"))) & tdC & vbNewLine
                sBody = sBody & trC & vbNewLine
                
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            sQty = trO & tdNO & f3O & bolO & "Total Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
        
        Case 22, 23, 24, 25
'            iOff = iCase - 22
            strSelect = "SELECT DISTINCT LL.ADDUSER " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL " & _
                        "WHERE LOCKREFSOURCE = 'DIL_OPEN' " & _
                        "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "ORDER BY LL.ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
'                sBody = sBody & trO & tdNO & f3O & bolO & UCase(Trim(rst.Fields("ADDUSER"))) & bolC & fC & tdNC & trC & vbNewLine
                sBody = sBody & trO & tdNO & f3O & UCase(Trim(rst.Fields("ADDUSER"))) & fC & tdNC & trC & vbNewLine
                sMess = sMess & UCase(Trim(rst.Fields("ADDUSER"))) & vbNewLine
                iQty = iQty + 1
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            sQty = trO & tdNO & f3O & bolO & "Total DIL Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
        Case 26, 27
'            If iCase = 26 Then iOff = 7 Else iOff = 30
            strSelect = "SELECT DISTINCT LL.ADDUSER " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL " & _
                        "WHERE LL.LOCKREFSOURCE = 'DIL_OPEN' " & _
                        "AND LL.LOCKOPENDTTM >= SYSDATE-" & iOff & " " & _
                        "ORDER BY LL.ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sMess = sMess & UCase(Trim(rst.Fields("ADDUSER"))) & vbNewLine
                sBody = sBody & trO & tdNO & f3O & UCase(Trim(rst.Fields("ADDUSER"))) & fC & tdNC & trC & vbNewLine
                iQty = iQty + 1
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            sQty = trO & tdNO & f3O & bolO & "Total DIL Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
'        Case 27
''            If iCase = 26 Then iOff = 7 Else iOff = 30
'            strSelect = "SELECT DISTINCT LL.ADDUSER " & _
'                        "FROM ANO_LOCKLOG LL " & _
'                        "WHERE LL.LOCKREFSOURCE = 'ANNO_OPEN' " & _
'                        "AND LL.LOCKOPENDTTM >= SYSDATE-" & iOff & " " & _
'                        "ORDER BY LL.ADDUSER"
'            sMess = ""
'            Set rst = Conn.Execute(strSelect)
'            Do While Not rst.EOF
'                sMess = sMess & UCase(Trim(rst.Fields("ADDUSER"))) & vbNewLine
'                sBody = sBody & trO & tdNO & f3O & UCase(Trim(rst.Fields("ADDUSER"))) & fC & tdNC & trC & vbNewLine
'                iQty = iQty + 1
'                rst.MoveNext
'            Loop
'            rst.Close: Set rst = Nothing
'
'            sQty = trO & tdNO & f3O & bolO & "Total DIL Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
        Case 4, 5, 6, 7
'            iOff = iCase - 4
            strSelect = "SELECT " & _
                        "LL.UPDUSER, C.ABALPH AS CLIENT, GM.GDESC " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, " & F0101 & " C, ANNOTATOR.GFX_MASTER GM " & _
                        "WHERE LOCKREFSOURCE = 'GFX_MASTER' " & _
                        "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "AND LL.LOCKREFID = GM.GID " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY UPDUSER, CLIENT, GDESC"
            sName = "": sMess = "": sShow = "": sBody = "": sQty = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sName <> UCase(Trim(rst.Fields("UPDUSER"))) Then
                    sName = UCase(Trim(rst.Fields("UPDUSER")))
                    iQty = iQty + 1
                    sBody = sBody & trO & tdNO & f3O & bolO & sName & bolC & fC & tdNC & trC & vbNewLine
'''                    If sMess = "" Then
'''                        sMess = sName & vbNewLine
'''                    Else
'''                        sMess = sMess & vbNewLine & sName & vbNewLine
'''                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  [ " & _
                            Trim(rst.Fields("GDESC")) & " ]"
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    
                    sBody = sBody & trO & vbNewLine
                    sBody = sBody & tdOa & "5" & tdOb & tdC & vbNewLine
                    sBody = sBody & tdOa & "95" & tdOb & sShow & tdC & vbNewLine
                    sBody = sBody & trC & vbNewLine
'                    sMess = sMess & vbTab & sShow & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
            sQty = trO & tdNO & f3O & bolO & "Total Graphic Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
        Case Is < 4
'            iOff = iCase
            strSelect = "SELECT " & _
                        "LL.UPDUSER, C.ABALPH AS CLIENT, DS.SHYR, S.ABALPH AS SHOW " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, ANNOTATOR.DWG_SHOW DS, " & F0101 & " C, " & F0101 & " S " & _
                        "WHERE LOCKREFSOURCE = 'DWG_MASTER' " & _
                        "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "AND LL.LOCKREFID = DS.DWGID " & _
                        "AND DS.AN8_CUNO = C.ABAN8 " & _
                        "AND DS.AN8_SHCD = S.ABAN8 " & _
                        "ORDER BY UPDUSER, CLIENT, SHYR, SHOW"
            sName = "": sMess = "": sShow = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sName <> UCase(Trim(rst.Fields("UPDUSER"))) Then
                    sName = UCase(Trim(rst.Fields("UPDUSER")))
                    iQty = iQty + 1
                    sBody = sBody & trO & tdNO & f3O & bolO & sName & bolC & fC & tdNC & trC & vbNewLine
'''                    If sMess = "" Then
'''                        sMess = sName & vbNewLine
'''                    Else
'''                        sMess = sMess & vbNewLine & sName & vbNewLine
'''                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  " & _
                            CStr(rst.Fields("SHYR")) & " - " & _
                            Trim(rst.Fields("SHOW"))
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sBody = sBody & trO & vbNewLine
                    sBody = sBody & tdOa & "5" & tdOb & tdC & vbNewLine
                    sBody = sBody & tdOa & "95" & tdOb & sShow & tdC & vbNewLine
                    sBody = sBody & trC & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing

            sQty = trO & tdNO & f3O & bolO & "Total Floorplan Users: " & iQty & bolC & fC & tdNC & trC & vbNewLine
            
    End Select
    
    If sQty <> "" Then
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & sQty & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        sHTML = sHTML & hr & vbNewLine
    End If
    
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & sBody & vbNewLine
    sHTML = sHTML & tblC & vbNewLine
    
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath & "Usage.htm"
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    web1.Navigate tFile1
    web1.Visible = True
    
    Me.MousePointer = 0
    
    cmdOptions.Visible = False
    
    
    
''        .PassQty = iQty
    
'''    Call CreateHTML(sTitle, sMess)
    
'''    MsgBox sMess, vbInformation, "LOG: " & format(Now, "MMMM D, YYYY")
    
End Sub


Private Sub cmdOptions_Click()
    Dim sFile As String
    
    Me.MousePointer = 11
    
    iUserReportView = Abs(iUserReportView - 1)
    Select Case iUserReportView
        Case 0
            sFile = GetUserList ''frmSecurity.GetUserList
            cmdOptions.Caption = "Re-sort by User Type"
        Case 1
            sFile = GetUserTypeList ''frmSecurity.GetUserTypeList
            cmdOptions.Caption = "Re-sort by User Name"
    End Select
    
    web1.Navigate sFile
    web1.Visible = True
    
    Me.MousePointer = 0

End Sub

Private Sub cmdUsageGantt_Click()
    frmUsageGannt.Show 1, Me
End Sub

Private Sub cmdUserReport_Click()
    Dim sFile As String
    
    Me.MousePointer = 11
    sFile = GetUserList ''frmSecurity.GetUserList
    web1.Navigate sFile
    web1.Visible = True
    cmdOptions.Caption = "Resort by User Type"
    iUserReportView = 0
    cmdOptions.Visible = True
    
'    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
'    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(True)
'    If Not bENABLE_PRINTERS Then lblPrint.Visible = True
'    '\\\\\ -------------------------------------------------------- /////
    
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()
    
'''    strHTMLPath = "\\DETMSFS01\GPJAnnotator\Support\HTML\"
    Me.Width = frmStartUp.Width
    Me.Height = frmStartUp.Height
    
    If bPerm(63) Then cmdUsageGantt.Visible = True Else cmdUsageGantt.Visible = False
    If bPerm(64) Then cmdUserReport.Visible = True Else cmdUserReport.Visible = False
    If bPerm(64) Then
        lMinWidth = Me.Width '' (Me.Width - Me.ScaleWidth) + cmdUserReport.Left + _
                    cmdUserReport.Width + web1.Left
    Else
        lMinWidth = Me.Width '' (Me.Width - Me.ScaleWidth) + cmdUsageGantt.Left + _
                    cmdUsageGantt.Width + web1.Left
    End If
    cmdOptions.Top = web1.Top + 120
    
    With cboViewUsage
        If bPerm(65) Then
            .AddItem "AnnoUsers - Today  (" & Format(Date, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 18
            .AddItem "AnnoUsers - Yesterday  (" & Format(Date - 1, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 19
            .AddItem "AnnoUsers - 2 Days Ago  (" & Format(Date - 2, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 20
            .AddItem "AnnoUsers - 3 Days Ago  (" & Format(Date - 3, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 21
        End If
        
        If bPerm(66) Then
            .AddItem "Floorplans - Today  (" & Format(Date, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 0
            .AddItem "Floorplans - Yesterday  (" & Format(Date - 1, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 1
            .AddItem "Floorplans - 2 Days Ago  (" & Format(Date - 2, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 2
            .AddItem "Floorplans - 3 Days Ago  (" & Format(Date - 3, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 3
            .AddItem "Floorplans - For Week"
            .ItemData(.NewIndex) = 10
        End If
        
        If bPerm(67) Then
            .AddItem "Graphics - Today  (" & Format(Date, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 4
            .AddItem "Graphics - Yesterday  (" & Format(Date - 1, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 5
            .AddItem "Graphics - 2 Days Ago  (" & Format(Date - 2, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 6
            .AddItem "Graphics - 3 Days Ago  (" & Format(Date - 3, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 7
            .AddItem "Graphics - For Week"
            .ItemData(.NewIndex) = 11
        End If
        
        If bPerm(68) Then
            .AddItem "Posted Gfx - Today  (" & Format(Date, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 12
            .AddItem "Posted Gfx - Yesterday  (" & Format(Date - 1, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 13
            .AddItem "Posted Gfx - 2 Days Ago  (" & Format(Date - 2, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 14
            .AddItem "Posted Gfx - 3 Days Ago  (" & Format(Date - 3, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 15
            .AddItem "Posted Gfx - For Week"
            .ItemData(.NewIndex) = 16
            .AddItem "Posted Gfx - For Month"
            .ItemData(.NewIndex) = 17
        End If
        
        If bPerm(69) Then
            .AddItem "DIL Users - Today  (" & Format(Date, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 22
            .AddItem "DIL Users - Yesterday  (" & Format(Date - 1, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 23
            .AddItem "DIL Users - 2 Days Ago  (" & Format(Date - 2, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 24
            .AddItem "DIL Users - 3 Days Ago  (" & Format(Date - 3, "dddd, mmm d") & ")"
            .ItemData(.NewIndex) = 25
            .AddItem "DIL Users - For Week"
            .ItemData(.NewIndex) = 26
            .AddItem "DIL Users - For Month"
            .ItemData(.NewIndex) = 27
        End If
    End With
    
    web1.Visible = False
    bDoneLoading = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < lMinWidth Then Me.Width = lMinWidth
    picHdr.Width = Me.ScaleWidth
    shpHDR.Width = picHdr.Width
    web1.Width = Me.ScaleWidth - (web1.Left * 2)
    web1.Height = Me.ScaleHeight - web1.Top - web1.Left
    cmdOptions.Left = web1.Left + web1.Width - 360 - cmdOptions.Width
    cmdUserReport.Left = Me.ScaleWidth - web1.Left - cmdUserReport.Width
    cmdUsageGantt.Left = cmdUserReport.Left - web1.Left - cmdUsageGantt.Width
End Sub

Private Sub web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim strSelect As String, sList As String, sMess As String, sName As String
    Dim rst As ADODB.Recordset
    Dim iChk As Integer, iDash As Integer, iHtm As Integer
    Dim lUser As Long
    
    If bDoneLoading = False Then
        web1.Visible = False
        Exit Sub
    End If
    
    Debug.Print URL
    sMess = ""
'''    iChk = InStr(1, URL, "?Rights=")
    If InStr(1, URL, "?Rights=") > 0 Then
        iChk = InStr(1, URL, "?Rights=")
        iChk = iChk + 8
        iDash = InStr(iChk, URL, "-")
        lUser = CLng(Mid(URL, iChk, iDash - iChk))
        sName = Mid(URL, iDash + 1)
        sName = Replace(sName, "%20", " ")
        iChk = InStr(1, sName, ",")
        If iChk > 0 Then sName = Mid(sName, iChk + 2) & " " & Left(sName, iChk - 1)
        
        
        frmUserLog.PassUser = sName
        frmUserLog.PassUserID = lUser
        frmUserLog.Show 1
        
        
'''        sList = GetClientList(lUser)
'''        Select Case sList
'''            Case "ALL"
'''                sMess = sName & " has access rights to all Clients."
'''            Case Else
'''                strSelect = "SELECT ABALPH FROM " & F0101 & " " & _
'''                            "WHERE ABAN8 IN (" & sList & ") " & _
'''                            "ORDER BY ABALPH"
'''                Set rst = Conn.Execute(strSelect)
'''                Do While Not rst.EOF
'''                    If sMess = "" Then
'''                        sMess = Trim(rst.Fields("ABALPH"))
'''                    Else
'''                        sMess = sMess & ", " & Trim(rst.Fields("ABALPH"))
'''                    End If
'''                    rst.MoveNext
'''                Loop
'''                rst.Close: Set rst = Nothing
'''                sMess = sName & " has access rights to the following Clients:" & _
'''                            vbNewLine & vbNewLine & sMess
'''        End Select
'''        MsgBox sMess, vbInformation, sName
        Cancel = True
    ElseIf InStr(1, URL, "?Desc=") > 0 Then
        iChk = InStr(1, URL, "?Desc=")
        iChk = iChk + 6
        iDash = InStr(iChk, URL, "-")
        lUser = CLng(Mid(URL, iChk, iDash - iChk))
        sName = Mid(URL, iDash + 1)
        sName = Replace(sName, "%20", " ")
        iChk = InStr(1, sName, ",")
        If iChk > 0 Then sName = Mid(sName, iChk + 2) & " " & Left(sName, iChk - 1)
        strSelect = "SELECT USERTYPEDESC FROM ANNOTATOR.ANO_USERTYPE " & _
                    "WHERE USERTYPEID = " & lUser
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sMess = sMess & Trim(rst.Fields("USERTYPEDESC"))
        Else
            sMess = "User Type could not be found."
        End If
        rst.Close: Set rst = Nothing
        MsgBox sMess, vbInformation, sName
        Cancel = True
    End If
    
'''    web1.Visible = True
End Sub

''Public Sub CreateHTML(tTitle As String, tMess As String)
''
''
''
''
''
''    sHTML = sHTML & sMess & vbNewLine
''
''
''
''End Sub

Public Function GetUserList() As String
    Dim strSelect As String, sUType As String, sHTML As String
    Dim tFile1 As String
    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
    Dim lUser As Long
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String, f2bO As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, _
                tdOa As String, tdObl As String, tdObc As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    Dim dl As String, dlC As String, dt As String, dtC As String
    Dim divO As String, divC As String
    Dim iUserCnt As Integer
    
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY LINK=""black"" VLINK=""black"" ALINK=""blue"">": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    f2bO = "<FONT SIZE=2 COLOR=""000080"" FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 ALIGN=""CENTER"" VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=3><DIV align=left><FONT SIZE=2 FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdObl = "%"" ALIGN=left VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    tdOa = "<TD WIDTH=""": tdObc = "%"" ALIGN=center VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    dl = "<DL>": dlC = "</DL>": dt = "<DT>": dtC = "</DT>"
    divO = "<DIV ALIGN=""RIGHT"">": divC = "</DIV>"
    
    
    strSelect = "SELECT COUNT(*) AS USERCNT " & _
                "FROM IGLPROD.IGL_USER_APP_R R, IGLPROD.IGL_USER U " & _
                "WHERE R.APP_ID = 1002 " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then iUserCnt = rst.Fields("USERCNT")
    rst.Close
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & "GPJ Annotator Users (by User Name)" & tiC & hdC & vbNewLine
    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & bolO & "GPJ Annotator Users (by User Name)" & bolC & fC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & tdOa & "25" & tdObl & "Total Number of Users:  " & bolO & iUserCnt & bolC & tdC & vbNewLine
    sHTML = sHTML & tdOa & "35" & tdObc & "User Type" & tdC & vbNewLine
    sHTML = sHTML & tdOa & "20" & tdObc & "Setup Date" & tdC & vbNewLine
    sHTML = sHTML & tdOa & "10" & tdObc & "Floorplans Accessed" & tdC & vbNewLine
    sHTML = sHTML & tdOa & "10" & tdObc & "Graphics Accessed" & tdC & vbNewLine
    sHTML = sHTML & trC & vbNewLine
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & tblO & vbNewLine
    
    strSelect = "SELECT UT.USERTYPE, TRIM(U.NAME_LAST) || ', ' || TRIM(U.NAME_FIRST) FULLNAME, " & _
                "TO_CHAR(UR.ADDDTTM, 'MON DD, YYYY') AS SETUP_DATE, U.USER_SEQ_ID, UT.USERTYPEID " & _
                "FROM IGLPROD.IGL_USER_APP_R UR, ANNOTATOR.ANO_USERTYPE UT, IGLPROD.IGL_USER U " & _
                "WHERE UR.APP_ID = 1002 " & _
                "AND UR.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "and UR.USER_PERMISSION_ID = UT.USERTYPEID " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
'''''        If rst.FIELDS("USERTYPE") <> sUType Then
'''''            If sUType <> "" Then
'''''                sHTML = sHTML & tblC & vbNewLine
'''''                sHTML = sHTML & divC & vbNewLine
'''''            End If
'''''            sUType = rst.FIELDS("USERTYPE")
'''''            sHTML = sHTML & dt & f2O & bolO & UCase(sUType) & bolC & fC & dtC & vbNewLine
'''''            sHTML = sHTML & divO & vbNewLine
'''''            sHTML = sHTML & tblO & vbNewLine
'''''            sHTML = sHTML & trO & vbNewLine
'''''        End If
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "25" & tdObl & "<A HREF=""" & strHTMLPath & "Pass.htm?Rights=" & _
                    rst.Fields("USER_SEQ_ID") & "-" & Trim(rst.Fields("FULLNAME")) & _
                    """ TITLE=""Click to View User Log & Client Access Rights"">" & bolO & _
                    UCase(Trim(rst.Fields("FULLNAME"))) & bolC & "</A>" & tdC & vbNewLine
        sHTML = sHTML & tdOa & "35" & tdObl & "<A HREF=""" & strHTMLPath & "Pass.htm?Desc=" & _
                    rst.Fields("USERTYPEID") & "-" & Trim(rst.Fields("USERTYPE")) & _
                    """ TITLE=""Click to View UserType Description"">" & UCase(Trim(rst.Fields("USERTYPE"))) & _
                    "</A>" & tdC & vbNewLine
        
'''        sHTML = sHTML & dt & "<A HREF=""Pass.htm?Desc=" & _
'''                    rst.FIELDS("USERTYPEID") & "-" & Trim(rst.FIELDS("USERTYPE")) & _
'''                    """ TITLE=""Click to View UserType Description"">" & f2O & bolO & _
'''                    UCase(sUType) & bolC & fC & "</A>" & dtC & vbNewLine
        
        sHTML = sHTML & tdOa & "20" & tdObc & UCase(Trim(rst.Fields("SETUP_DATE"))) & tdC & vbNewLine
        
        ''GET ALL INSTANCES''
'        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
'                    "(SELECT COUNT(*) AS FP_COUNT " & _
'                    "FROM ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
'                    "AND LOCKREFSOURCE = 'DWG_MASTER') LF, " & _
'                    "(SELECT COUNT(*) AS GFX_COUNT " & _
'                    "FROM ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
'                    "AND LOCKREFSOURCE = 'GFX_MASTER') LG"
        
        ''GET LAST 365 DAY ONLY''
        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
                    "(SELECT COUNT(*) AS FP_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE LOCKID > 0 " & _
                    "AND USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'DWG_MASTER' " & _
                    "AND ADDDTTM > SYSDATE - 365) LF, " & _
                    "(SELECT COUNT(*) AS GFX_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE LOCKID > 0 " & _
                    "AND USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'GFX_MASTER' " & _
                    "AND ADDDTTM > SYSDATE - 365) LG"
                    
        Set rstX = Conn.Execute(strSelect)
        If Not rstX.EOF Then
            sHTML = sHTML & tdOa & "10" & tdObc & rstX.Fields("FP_COUNT") & tdC & vbNewLine
            sHTML = sHTML & tdOa & "10" & tdObc & rstX.Fields("GFX_COUNT") & tdC & vbNewLine
        Else
            sHTML = sHTML & tdOa & "10" & tdObc & "" & tdC & vbNewLine
            sHTML = sHTML & tdOa & "10" & tdObc & "" & tdC & vbNewLine
        End If
        rstX.Close
        
        
        sHTML = sHTML & trC & vbNewLine
        
        rst.MoveNext
    Loop
    Set rstX = Nothing
    rst.Close: Set rst = Nothing
    
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath & "Users.htm"
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    GetUserList = tFile1
    
End Function

Public Function GetUserTypeList() As String
    Dim strSelect As String, sUType As String, sHTML As String
    Dim tFile1 As String
    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
    Dim lUser As Long
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String, f2bO As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, _
                tdOa As String, tdObl As String, tdObc As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    Dim dl As String, dlC As String, dt As String, dtC As String
    Dim divO As String, divC As String
    Dim iUserCnt As Integer
    
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY LINK=""black"" VLINK=""black"" ALINK=""blue"">": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    f2bO = "<FONT SIZE=2 COLOR=""000080"" FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""95%"" BORDER=0 ALIGN=""CENTER"" VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""000080"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=3><DIV align=left><FONT SIZE=2 FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdObl = "%"" ALIGN=left VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    tdOa = "<TD WIDTH=""": tdObc = "%"" ALIGN=center VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    dl = "<DL>": dlC = "</DL>": dt = "<DT>": dtC = "</DT>"
    divO = "<DIV ALIGN=""RIGHT"">": divC = "</DIV>"
    
    
    strSelect = "SELECT COUNT(*) AS USERCNT " & _
                "FROM IGLPROD.IGL_USER_APP_R R, IGLPROD.IGL_USER U " & _
                "WHERE R.APP_ID = 1002 " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then iUserCnt = rst.Fields("USERCNT")
    rst.Close
    
    sHTML = htmO & vbNewLine
    sHTML = sHTML & hdO & tiO & "GPJ Annotator Users (by User Type)" & tiC & hdC & vbNewLine
    sHTML = sHTML & bodO & vbNewLine
    sHTML = sHTML & f3O & bolO & "GPJ Annotator Users (by User Type)" & bolC & fC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & divO & vbNewLine
    sHTML = sHTML & tblO & vbNewLine
    sHTML = sHTML & tdOa & "40" & tdObl & "Total Number of Users:  " & bolO & iUserCnt & bolC & tdC & vbNewLine
    sHTML = sHTML & tdOa & "30" & tdObc & "Setup Date" & tdC & vbNewLine
    sHTML = sHTML & tdOa & "15" & tdObc & "Floorplans Accessed" & tdC & vbNewLine
    sHTML = sHTML & tdOa & "15" & tdObc & "Graphics Accessed" & tdC & vbNewLine
    sHTML = sHTML & trC & vbNewLine
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & divC & vbNewLine
    
    sHTML = sHTML & dl & vbNewLine
    
    sUType = "": lUser = 0
    strSelect = "SELECT UT.USERTYPE, TRIM(U.NAME_FIRST) || ' ' || TRIM(U.NAME_LAST) FULLNAME, " & _
                "TO_CHAR(UR.ADDDTTM, 'MON DD, YYYY') AS SETUP_DATE, U.USER_SEQ_ID, UT.USERTYPEID " & _
                "FROM IGLPROD.IGL_USER_APP_R UR, ANNOTATOR.ANO_USERTYPE UT, IGLPROD.IGL_USER U " & _
                "WHERE UR.APP_ID = 1002 " & _
                "AND UR.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "and UR.USER_PERMISSION_ID = UT.USERTYPEID " & _
                "ORDER BY UT.USERTYPE, U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If Trim(rst.Fields("USERTYPE")) <> sUType Then
            If sUType <> "" Then
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & divC & vbNewLine
            End If
            sUType = Trim(rst.Fields("USERTYPE"))
            sHTML = sHTML & dt & "<A HREF=""" & strHTMLPath & "Pass.htm?Desc=" & _
                    rst.Fields("USERTYPEID") & "-" & Trim(rst.Fields("USERTYPE")) & _
                    """ TITLE=""Click to View UserType Description"">" & f2O & bolO & _
                    UCase(sUType) & bolC & fC & "</A>" & dtC & vbNewLine
            sHTML = sHTML & divO & vbNewLine
            sHTML = sHTML & tblO & vbNewLine
            sHTML = sHTML & trO & vbNewLine
        End If
        sHTML = sHTML & tdOa & "40" & tdObl & "<A HREF=""" & strHTMLPath & "Pass.htm?Rights=" & _
                    rst.Fields("USER_SEQ_ID") & "-" & Trim(rst.Fields("FULLNAME")) & _
                    """ TITLE=""Click to View User Log & Client Access Rights"">" & _
                    UCase(Trim(rst.Fields("FULLNAME"))) & "</A>" & tdC & vbNewLine
        sHTML = sHTML & tdOa & "30" & tdObc & UCase(Trim(rst.Fields("SETUP_DATE"))) & tdC & vbNewLine
        
'''        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
'''                    "(SELECT COUNT(*) AS FP_COUNT " & _
'''                    "FROM ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
'''                    "AND LOCKREFSOURCE = 'DWG_MASTER') LF, " & _
'''                    "(SELECT COUNT(*) AS GFX_COUNT " & _
'''                    "FROM ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
'''                    "AND LOCKREFSOURCE = 'GFX_MASTER') LG"
                    
        ''GET LAST 365 DAY ONLY''
        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
                    "(SELECT COUNT(*) AS FP_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'DWG_MASTER' " & _
                    "AND ADDDTTM > SYSDATE - 365) LF, " & _
                    "(SELECT COUNT(*) AS GFX_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'GFX_MASTER' " & _
                    "AND ADDDTTM > SYSDATE - 365) LG"
                    
        Set rstX = Conn.Execute(strSelect)
        If Not rstX.EOF Then
            sHTML = sHTML & tdOa & "15" & tdObc & rstX.Fields("FP_COUNT") & tdC & vbNewLine
            sHTML = sHTML & tdOa & "15" & tdObc & rstX.Fields("GFX_COUNT") & tdC & vbNewLine
        Else
            sHTML = sHTML & tdOa & "15" & tdObc & "" & tdC & vbNewLine
            sHTML = sHTML & tdOa & "15" & tdObc & "" & tdC & vbNewLine
        End If
        rstX.Close
        
        
        sHTML = sHTML & trC & vbNewLine
        
        rst.MoveNext
    Loop
    Set rstX = Nothing
    rst.Close: Set rst = Nothing
    
    sHTML = sHTML & tblC & vbNewLine
    sHTML = sHTML & divC & vbNewLine
    sHTML = sHTML & dlC & vbNewLine
    sHTML = sHTML & hr & vbNewLine
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath & "Users.htm"
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    GetUserTypeList = tFile1
    
End Function

