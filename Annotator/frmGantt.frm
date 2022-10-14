VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmGantt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGantt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Display Date Range of Current Show"
      Height          =   255
      Left            =   7500
      TabIndex        =   5
      Top             =   60
      Width           =   2955
   End
   Begin VB.PictureBox picOuter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   1080
      ScaleHeight     =   3135
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   600
      Width           =   8955
      Begin VB.PictureBox picInner 
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   0
         ScaleHeight     =   1395
         ScaleWidth      =   5235
         TabIndex        =   2
         Top             =   1320
         Width           =   5235
         Begin VB.CommandButton cmd1 
            BackColor       =   &H000080FF&
            Height          =   195
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Line lin1 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   2280
            X2              =   2280
            Y1              =   -120
            Y2              =   1200
         End
         Begin VB.Shape shpRange 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000040C0&
            Height          =   1335
            Left            =   3900
            Top             =   0
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Shape shpShow 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            Height          =   1335
            Left            =   3120
            Top             =   -60
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Shape shpCurrent 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000010&
            FillColor       =   &H0080C0FF&
            FillStyle       =   7  'Diagonal Cross
            Height          =   1335
            Left            =   2940
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shp1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000040C0&
            Height          =   435
            Index           =   0
            Left            =   2100
            Top             =   480
            Visible         =   0   'False
            Width           =   435
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   15
      Cols            =   16
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   1
      BorderStyle     =   0
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:  Date Blocks represent Element Reservation Dates, not Show Dates."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   5430
   End
   Begin VB.Menu mnuShowData 
      Caption         =   "mnuShowData"
      Visible         =   0   'False
      Begin VB.Menu mnuData 
         Caption         =   "Client:  "
         Index           =   0
      End
      Begin VB.Menu mnuData 
         Caption         =   "Show:  "
         Index           =   1
      End
      Begin VB.Menu mnuData 
         Caption         =   "Facility:  "
         Index           =   2
      End
      Begin VB.Menu mnuData 
         Caption         =   "Location:  "
         Index           =   3
      End
      Begin VB.Menu mnuData 
         Caption         =   "Show Open:  "
         Index           =   4
      End
      Begin VB.Menu mnuData 
         Caption         =   "Show Close:  "
         Index           =   5
      End
      Begin VB.Menu mnuData 
         Caption         =   "Setup Begins:"
         Index           =   6
      End
      Begin VB.Menu mnuData 
         Caption         =   "Takedown Ends:"
         Index           =   7
      End
      Begin VB.Menu mnuData 
         Caption         =   "IGL Start:"
         Index           =   8
      End
      Begin VB.Menu mnuData 
         Caption         =   "IGL End:"
         Index           =   9
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmGantt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sData(0 To 20, 0 To 9) As String
Dim bGetRange As Boolean
Dim d0 As Date
Dim CRDate As Date
Dim CRWidth As Long
Dim iCmd As Integer

Dim tLink As String

Public Property Get PassLink() As String
    PassLink = tLink
End Property
Public Property Let PassLink(ByVal vNewValue As String)
    tLink = vNewValue
End Property





Private Sub Check1_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim dS As Date, dE As Date, dSS As Date, dTE As Date
    Dim iCol As Integer
    
    Select Case Check1.Value
        Case 0
            shpCurrent.Visible = False
        Case 1
            strSelect = "SELECT " & _
                        "IGL_JDEDATE_TOCHAR(SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, " & _
                        "IGL_JDEDATE_TOCHAR(SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, " & _
                        "IGL_JDEDATE_TOCHAR(SHY56SBEDT, 'MM/DD/YYYY') AS SBEG, " & _
                        "IGL_JDEDATE_TOCHAR(SHY56TEDDT, 'MM/DD/YYYY') AS TEND " & _
                        "From " & F5601 & " " & _
                        "Where SHY56SHCD = " & SHCD & " " & _
                        "AND SHY56SHYR = " & SHYR
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                dSS = DateValue(rst.Fields("SBEG"))
                dTE = DateValue(rst.Fields("TEND"))
                dS = DateValue(rst.Fields("BEGD"))
                dE = DateValue(rst.Fields("ENDD"))
                rst.Close
                
                shpCurrent.Left = (DateValue(dSS) - d0) * 50
                shpCurrent.Width = ((dTE - dSS) + 1) * 50
                shpCurrent.Visible = True
                
                shpShow.Left = (DateValue(dS) - d0) * 50
                shpShow.Width = ((dE - dS) + 1) * 50
                shpShow.Visible = True
                
                iCol = Int((dSS - d0) / 30.4) - 1
                If iCol < 1 Then iCol = 1
                flx1.LeftCol = iCol
            Else
                rst.Close
                shpCurrent.Visible = False
            End If
            Set rst = Nothing
            
    End Select
        
End Sub


Private Sub cmd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim sMess As String
    
'''    If Button = vbRightButton Then
        sMess = ""
        For i = mnuData.LBound To mnuData.UBound
            sMess = sMess & sData(Index, i) & vbNewLine
            mnuData(i).Caption = sData(Index, i)
        Next i
        iCmd = Index
        
        MsgBox sMess, vbInformation, flx1.TextMatrix(Index, 0)
'        frmGantt.PopupMenu mnuShowData
'''    End If
End Sub

Private Sub flx1_Click()
    Dim iCol As Integer
    If bGetRange Then
        shpRange.Left = cmd1(flx1.RowSel).Left
        shpRange.Width = cmd1(flx1.RowSel).Width
        iCol = Int((DateValue(Right(sData(flx1.RowSel, 4), 12)) - d0) / 30.4) - 1
        If iCol < 1 Then iCol = 1
        flx1.LeftCol = iCol
        shpRange.Visible = True
    End If
End Sub

Private Sub flx1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x < flx1.ColWidth(0) And y > flx1.RowHeight(0) Then
        bGetRange = True
    Else
        bGetRange = False
    End If
End Sub

Private Sub flx1_Scroll()
    Dim i As Integer, iCol As Integer
    Dim lWidth As Long
    
    iCol = flx1.LeftCol - 1
    lWidth = 0
    For i = 1 To iCol
        lWidth = lWidth + flx1.ColWidth(i)
    Next i
    picInner.Left = lWidth * -1
    
    Debug.Print picInner.Left / -50
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer, iCount As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
    Dim lBCC As Long, lYear As Long, lWidth As Long, lLeft As Long
    Dim sElt As String, sMess As String, sClient As String, sTitle As String

    
    flx1.ColAlignment(0) = 1
    flx1.ColWidth(0) = 3000
    
    lBCC = CLng(Left(tLink, 8))
    sElt = Mid(tLink, 10)
    sClient = "": sMess = "": lYear = 0
    
    strSelect = "SELECT COUNT(S.ABALPH) AS CNT " & _
                "FROM " & AQUAEltU & " EU, " & F0101 & " S " & _
                "WHERE EU.ELTID IN " & _
                "(SELECT E.ELTID " & _
                "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K " & _
                "Where K.AN8_CUNO = " & lBCC & " " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTFNAME = '" & sElt & "') " & _
                "AND EU.DTASNSTR > SYSDATE-365 " & _
                "AND EU.DTASNSTR < SYSDATE +365 " & _
                "AND EU.AN8_SHCD = S.ABAN8"
    Set rst = Conn.Execute(strSelect)
    iCount = rst.Fields("CNT")
    flx1.Rows = iCount + 1
    rst.Close
    If iCount = 1 Then
        flx1.TextMatrix(0, 0) = "( " & iCount & " ) Use in 2-Year Window"
    Else
        flx1.TextMatrix(0, 0) = "( " & iCount & " ) Uses in 2-Year Window"
    End If
    
    If iCount = 0 Then
        sMess = "Element " & sElt & " is not scheduled within the scope of the search." & _
                    vbNewLine & vbNewLine & _
                    "Search Dates:  " & Format(Now - 365, "mmmm d, yyyy") & " to " & _
                    Format(Now + 365, "mmmm d, yyyy")
        MsgBox sMess, vbExclamation, "Element is not Scheduled..."
        
        Unload Me
'''        GoTo GettingOut
        Exit Sub
    End If
    
    picOuter.Top = flx1.Top + flx1.RowHeight(0)
    picOuter.Left = flx1.Left + flx1.ColWidth(0)
    picOuter.Width = flx1.Width - flx1.ColWidth(0)
    picOuter.Height = flx1.RowHeight(0) * (flx1.Rows - 1) ' - 250
    picInner.Height = picOuter.Height
    picInner.Top = 0
    picInner.Left = 0
    
    shpRange.Top = 0: shpRange.Height = picInner.Height
    shpCurrent.Top = 0: shpCurrent.Height = picInner.Height
    shpShow.Top = 0: shpShow.Height = picInner.Height
    
    i = 0
    strSelect = "SELECT C.ABALPH AS CLIENT, SM.SHY56FCCDT AS FACILITY, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, " & _
                "IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, " & _
                "IGL_JDEDATE_TOCHAR(SHY56SBEDT, 'MM/DD/YYYY') AS SBEG, " & _
                "IGL_JDEDATE_TOCHAR(SHY56TEDDT, 'MM/DD/YYYY') AS TEND, " & _
                "EU.SHYR, EU.AN8_SHCD, S.ABALPH AS SHOW, " & _
                "TO_CHAR(EU.DTASNSTR, 'MON DD, YYYY') AS D1, " & _
                "TO_CHAR(EU.DTASNEND, 'MON DD, YYYY') AS D2 " & _
                "FROM " & AQUAEltU & " EU, " & F0101 & " S, " & _
                "" & F5601 & " SM, " & F0101 & " C " & _
                "WHERE EU.ELTID IN " & _
                "(SELECT E.ELTID " & _
                "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K " & _
                "Where K.AN8_CUNO = " & lBCC & " " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTFNAME = '" & sElt & "') " & _
                "AND EU.DTASNSTR > SYSDATE-365 " & _
                "AND EU.DTASNSTR < SYSDATE +365 " & _
                "AND EU.AN8_SHCD = S.ABAN8 " & _
                "AND EU.AN8_CUNO = C.ABAN8 " & _
                "AND EU.SHYR = SM.SHY56SHYR " & _
                "AND EU.AN8_SHCD = SM.SHY56SHCD " & _
                "ORDER BY EU.DTASNSTR"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sTitle = "2-Year Gantt View:  " & Trim(rst.Fields("CLIENT")) & _
                    "  (Element - " & sElt & ")"
        d0 = DateValue(rst.Fields("D1"))
        d0 = DateAdd("d", (CInt(Format(d0, "d")) * -1) + 1, d0)
        Call SetGanttBackground(d0)
        Do While Not rst.EOF
            i = i + 1
            Load cmd1(i)
            lWidth = (CLng(DateValue(rst.Fields("D2"))) - _
                        CLng(DateValue(rst.Fields("D1"))) + 1) * 50
            lLeft = (DateValue(rst.Fields("D1")) - d0) * 50
            cmd1(i).Width = lWidth
            cmd1(i).Left = lLeft
            flx1.Row = i: flx1.Col = 0
            flx1.Text = rst.Fields("SHYR") & " - " & Trim(rst.Fields("SHOW"))
            cmd1(i).Top = flx1.CellTop - flx1.RowHeight(0) + 30
            cmd1(i).Visible = True
            cmd1(i).ToolTipText = rst.Fields("SHYR") & " - " & Trim(rst.Fields("SHOW")) & _
                        "  (" & Trim(rst.Fields("D1")) & " - " & _
                        Trim(rst.Fields("D2")) & ")"
            
            sData(i, 0) = "Client:  " & Trim(rst.Fields("CLIENT"))
            sData(i, 1) = "Show:  " & Trim(rst.Fields("SHOW"))
            sData(i, 4) = "Show Open:  " & Format(DateValue(Trim(rst.Fields("BEGD"))), "dddd, mmm d, yyyy")
            sData(i, 5) = "Show Close:  " & Format(DateValue(Trim(rst.Fields("ENDD"))), "dddd, mmm d, yyyy")
            sData(i, 6) = "Setup Begins:  " & Format(DateValue(Trim(rst.Fields("SBEG"))), "dddd, mmm d, yyyy")
            sData(i, 7) = "Takedown Ends:  " & Format(DateValue(Trim(rst.Fields("TEND"))), "dddd, mmm d, yyyy")
            
            strSelect = "SELECT AB.ABALPH, (TRIM(AL.ALCTY1) || ', ' || TRIM(AL.ALADDS))LOC " & _
                        "FROM " & F0101 & " AB, " & F0116 & " AL " & _
                        "Where AB.ABAN8 = " & rst.Fields("FACILITY") & " " & _
                        "AND AB.ABAN8 = AL.ALAN8"
            Set rstX = Conn.Execute(strSelect)
            If Not rstX.EOF Then
                sData(i, 2) = "Facility:  " & Trim(rstX.Fields("ABALPH"))
                sData(i, 3) = "Location:  " & Trim(rstX.Fields("LOC"))
            End If
            
            rstX.Close: Set rstX = Nothing
            
            sData(i, 8) = "IGL Ship Date:  " & Format(DateValue(Trim(rst.Fields("D1"))), "dddd, mmm d, yyyy")
            sData(i, 9) = "IGL Return Date:  " & Format(DateValue(Trim(rst.Fields("D2"))), "dddd, mmm d, yyyy")
            
            If rst.Fields("AN8_SHCD") = SHCD Then
                CRDate = DateValue(rst.Fields("D1"))
                CRWidth = lWidth
                shpCurrent.Width = CRWidth
            End If

            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    

    
    flx1.Height = flx1.Rows * flx1.RowHeight(0) + 250
    Me.Height = (Me.Height - Me.ScaleHeight) + flx1.Height + flx1.Top + flx1.Left
    Me.Caption = sTitle
    
End Sub

Public Sub SetGanttBackground(d0 As Date)
    Dim d1 As Date, d2 As Date
    Dim i As Integer, iMonLen As Integer
    Dim lHgt As Long
    Dim lTotalWidth As Long
    Dim iMons As Integer
    
    lTotalWidth = 0
    
    d1 = DateValue(Format(DateAdd("m", 1, Now), "mm") & "/01/" & _
                CInt(Format(DateAdd("m", 1, Now), "yyyy")) + 1)
    iMons = 0
    Do While DateAdd("m", iMons, d0) <> d1
        iMons = iMons + 1
    Loop
    flx1.Cols = iMons + 1
    
    d1 = d0
    For i = 1 To iMons
        d2 = DateAdd("m", i, d0) - 1
        iMonLen = CInt(Format(d2, "d"))
        flx1.ColWidth(i) = iMonLen * 50
        lTotalWidth = lTotalWidth + flx1.ColWidth(i)
        flx1.Row = 0: flx1.Col = i: flx1.CellAlignment = 4
        flx1.TextMatrix(0, i) = Format(d2, "MMMM YYYY")
        Load shp1(i)
        shp1(i).Visible = True
        shp1(i).Width = flx1.ColWidth(i)
        shp1(i).Height = picOuter.Height
        shp1(i).Top = 0
        shp1(i).Left = (d1 - d0) * 50
        If i Mod 2 = 1 Then
            shp1(i).BackColor = vb3DLight
            shp1(i).BorderColor = vb3DLight
        Else
            shp1(i).BackColor = vbWindowBackground
            shp1(i).BorderColor = vbWindowBackground
        End If
        d1 = d2 + 1
    Next i
    
    lin1.X1 = ((Now - d0) * 50) + 25
    lin1.X2 = lin1.X1
    lin1.Y1 = 0
    lin1.Y2 = picInner.Height
    
    picInner.Width = lTotalWidth
    
    On Error Resume Next
    flx1.LeftCol = Int((Now - d0) / 30.4) - 1
End Sub

