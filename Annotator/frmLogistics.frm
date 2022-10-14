VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmLogistics 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntireInv 
      Caption         =   "View Entire Inventory..."
      Height          =   435
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4860
      Width           =   1995
   End
   Begin VB.CommandButton cmdEntireKit 
      Caption         =   "View Entire Kit..."
      Height          =   435
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4860
      Width           =   1995
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogistics.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogistics.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogistics.frx":0B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogistics.frx":1118
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogistics.frx":16B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw1 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   6482
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Click Element Node to view Element Usage details "
      Height          =   855
      Left            =   4140
      TabIndex        =   4
      Top             =   3900
      UseMnemonic     =   0   'False
      Width           =   1275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3900
      UseMnemonic     =   0   'False
      Width           =   5235
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "mnuCheck"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckUsage 
         Caption         =   "Check Usage..."
      End
   End
End
Attribute VB_Name = "frmLogistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCurrCNode As String, sCurrKNode As String, sCurrENode As String
Dim lBCC As Long
Dim sElt As String
Dim sTempLink As String

Dim pBCC As Long
Dim pFBCN As String
Dim pEID As Long
Dim pElem As String
Dim pFrom As String

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = pFBCN
End Property
Public Property Let PassFBCN(ByVal vNewValue As String)
    pFBCN = vNewValue
End Property

Public Property Get PassEID() As Long
    PassEID = pEID
End Property
Public Property Let PassEID(ByVal vNewValue As Long)
    pEID = vNewValue
End Property

Public Property Get PassElem() As String
    PassElem = pElem
End Property
Public Property Let PassElem(ByVal vNewValue As String)
    pElem = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property


Private Sub cmdEntireInv_Click()
    Dim strSelect As String
    strSelect = "SELECT AB.ABAN8 AS CUNO, AB.ABALPH AS CLIENT, " & _
                "K.KITID, K.KITFNAME, " & _
                "E.ELTID, E.ELTFNAME, E.ELTDESC, P.PARTID, " & _
                "P.PARTDESC, P.WEIGHT, P.WTUNIT, " & _
                "(P.FABLOC || TRIM(TO_CHAR(P.YRBUILT, 'YY')) || '-' || P.PNUMBER) AS PNUM " & _
                "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K, IGLPROD.IGL_PART P, " & F0101 & " AB " & _
                "Where K.KITID > 0 " & _
                "AND K.AN8_CUNO = " & lBCC & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ELTID > 0 " & _
                "AND E.ESTATUS > 2 " & _
                "AND P.KITID = K.KITID " & _
                "AND P.KITID = E.KITID " & _
                "AND P.ELTID = E.ELTID " & _
                "AND P.PARTID > 0 " & _
                "AND P.TSTATUS > 0 " & _
                "AND AB.ABAN8 = " & lBCC & " " & _
                "ORDER BY K.KITREF, E.ELTCODE, E.ELSUFFIX, PNUM"
    Me.Caption = GetInven(strSelect, "ALL")
    
End Sub

Private Sub cmdEntireKit_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sMess As String, wgtSuf As String, sDisclaim As String
    Dim dWgt As Double
    Dim bWgtErr As Boolean
    Dim nodX As Node
    Dim sCNode As String, sKNode As String, sENode As String, sPNode As String
    Dim sDesc As String
    Dim sHDR As String
    
'''    tvw1.Visible = False: tvw1.Nodes.Clear: tvw1.Visible = True
'''    tvw1.ImageList = ImageList1
'''    dWgt = 0: bWgtErr = False
    strSelect = "SELECT AB.ABAN8 AS CUNO, AB.ABALPH AS CLIENT, " & _
                "K.KITID, K.KITFNAME, " & _
                "E.ELTID, E.ELTFNAME, E.ELTDESC, P.PARTID, " & _
                "P.PARTDESC, P.WEIGHT, P.WTUNIT, " & _
                "(P.FABLOC || TRIM(TO_CHAR(P.YRBUILT, 'YY')) || '-' || P.PNUMBER) AS PNUM " & _
                "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K, IGLPROD.IGL_PART P, " & F0101 & " AB " & _
                "Where K.KITID = " & Mid(sCurrKNode, 2) & " " & _
                "AND K.KSTATUS > 0 " & _
                "AND K.KITID = E.KITID " & _
                "AND E.ESTATUS > 2 " & _
                "AND P.KITID = K.KITID " & _
                "AND P.KITID = E.KITID " & _
                "AND P.ELTID = E.ELTID " & _
                "AND P.TSTATUS > 0 " & _
                "AND AB.ABAN8 = " & lBCC & " " & _
                "ORDER BY K.KITREF, E.ELTCODE, E.ELSUFFIX, PNUM"
    Me.Caption = GetInven(strSelect, "KIT")
    
    tvw1.Nodes(sCurrCNode).Expanded = True
    tvw1.Nodes(sCurrKNode).Expanded = True
    tvw1.Nodes(sCurrENode).Expanded = True
    
End Sub

Private Sub Form_Load()
        Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sMess As String, wgtSuf As String, sDisclaim As String
    Dim dWgt As Double
    Dim bWgtErr As Boolean
    Dim nodX As Node
    Dim sCNode As String, sKNode As String, sENode As String, sPNode As String
    Dim sDesc As String
    Dim sHDR As String
    
    Select Case pFrom
        Case "GFX"
            strSelect = "SELECT PARTID, PARTDESC, PKGTYPE, " & _
                        "NVL(WIDTH, 0) AS WIDTH, NVL(HEIGHT, 0) AS HEIGHT, " & _
                        "NVL(LENGTH, 0) AS LENGTH, SIZEUNIT, " & _
                        "NVL(WEIGHT, 0) AS WEIGHT, WTUNIT, " & _
                        "(FABLOC||TO_CHAR(YRBUILT, 'YY')||'-'||PNUMBER)PNUM " & _
                        "FROM IGLPROD.IGL_PART " & _
                        "WHERE ELTID = " & pEID & " " & _
                        "AND TSTATUS > 0 " & _
                        "ORDER BY PARTDESC, PNUM"
            Me.Caption = GetElement(strSelect)
            Me.Height = 5280
            
        Case "FP"
            lBCC = CLng(Left(sLinkID, 8))
            sElt = Mid(sLinkID, 10)
            
            strSelect = "SELECT AB.ABAN8 AS CUNO, AB.ABALPH AS CLIENT, " & _
                        "K.KITID, K.KITFNAME, " & _
                        "E.ELTID, E.ELTFNAME, E.ELTDESC, " & _
                        "P.PARTID, P.PARTDESC, P.WEIGHT, P.WTUNIT, " & _
                        "(P.FABLOC || TRIM(TO_CHAR(P.YRBUILT, 'YY')) || '-' || P.PNUMBER) AS PNUM " & _
                        "FROM IGLPROD.IGL_ELEMENT E, IGLPROD.IGL_KIT K, IGLPROD.IGL_PART P, " & F0101 & " AB " & _
                        "Where K.KITID > 0 " & _
                        "AND K.AN8_CUNO = " & lBCC & " " & _
                        "AND K.KSTATUS > 0 " & _
                        "AND K.KITID = E.KITID " & _
                        "AND E.ELTID > 0 " & _
                        "AND E.ELTFNAME = '" & sElt & "' " & _
                        "AND E.ESTATUS > 2 " & _
                        "AND P.KITID = K.KITID " & _
                        "AND P.KITID = E.KITID " & _
                        "AND P.ELTID = E.ELTID " & _
                        "AND P.TSTATUS > 0 " & _
                        "AND AB.ABAN8 = " & lBCC & " " & _
                        "ORDER BY K.KITREF, E.ELTCODE, E.ELSUFFIX, PNUM"
            Me.Caption = GetInven(strSelect, "INIT")
            
            If tvw1.Nodes.Count > 0 Then
                tvw1.Nodes(sCurrCNode).Expanded = True
                tvw1.Nodes(sCurrKNode).Expanded = True
                tvw1.Nodes(sCurrENode).Expanded = True
                Me.Height = 5910
            Else
                MsgBox "No Inventory data is available for this Element", _
                            vbExclamation, "Sorry..."
                Unload Me
                Exit Sub
            End If
            
    End Select
    
End Sub

Private Sub mnuCheckUsage_Click()
    On Error Resume Next
    frmGantt.PassLink = sTempLink
    frmGantt.Show ''' 1

End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If UCase(Left(Node.Key, 1)) <> "E" Then Exit Sub
    
    strSelect = "SELECT ELTFNAME FROM IGLPROD.IGL_ELEMENT " & _
                "WHERE ELTID = " & Mid(Node.Key, 2)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sTempLink = Right("00000000" & CStr(lBCC), 8) & "-" & Trim(rst.Fields("ELTFNAME"))
    End If
    rst.Close: Set rst = Nothing
    
    On Error Resume Next
    frmGantt.PassLink = sTempLink
    frmGantt.Show 1, Me
'''    Me.PopupMenu mnuCheck
        
End Sub

Public Function GetInven(strSelect As String, sType As String) As String
    Dim rst As ADODB.Recordset
    Dim sMess As String, wgtSuf As String, sDisclaim As String
    Dim dWgt As Double
    Dim bWgtErr As Boolean
    Dim nodX As Node
    Dim sCNode As String, sKNode As String, sENode As String, sPNode As String
    Dim sDesc As String
    Dim sHDR As String
    
    tvw1.Visible = False: tvw1.Nodes.Clear: tvw1.Visible = True
    tvw1.ImageList = ImageList1
    dWgt = 0: bWgtErr = False
    
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If rst.Fields("WTUNIT") Then
            wgtSuf = " lbs"
        Else: wgtSuf = " kg"
        End If
        Do While Not rst.EOF
            If rst.Fields("WEIGHT") = 0 Then
                bWgtErr = True
            Else
                dWgt = dWgt + rst.Fields("WEIGHT")
            End If
            If sCNode = "" Then
                sCNode = "c" & rst.Fields("CUNO")
                Set nodX = tvw1.Nodes.Add(, , sCNode, Trim(rst.Fields("CLIENT")), 1)
                If sType = "INIT" Then sCurrCNode = sCNode
                Select Case sType
                    Case "ALL": sHDR = UCase(Trim(rst.Fields("CLIENT"))) & " Inventory"
                    Case Else: sHDR = UCase(Trim(rst.Fields("CLIENT")))
                End Select
            End If
            If sKNode = "" Or sKNode <> "k" & rst.Fields("KITID") Then
                sKNode = "k" & rst.Fields("KITID")
                If sType = "INIT" Then sCurrKNode = sKNode
                Set nodX = tvw1.Nodes.Add(sCNode, tvwChild, sKNode, _
                            "Kit: " & Trim(rst.Fields("KITFNAME")), 2)
                Select Case sType
                    Case "KIT", "INIT": sHDR = sHDR & " - " & Trim(rst.Fields("KITFNAME"))
                End Select
            End If
            If sENode = "" Or sENode <> "e" & rst.Fields("ELTID") Then
                sENode = "e" & rst.Fields("ELTID")
                If sType = "INIT" Then sCurrENode = sENode
                Set nodX = tvw1.Nodes.Add(sKNode, tvwChild, sENode, _
                            "Element: " & Trim(rst.Fields("ELTFNAME")), 3)
                If sType = "INIT" Then sHDR = sHDR & "/" & Trim(rst.Fields("ELTFNAME"))
            End If
            
            sPNode = "p" & rst.Fields("PARTID")
            sDesc = "Part: " & Trim(rst.Fields("PNUM")) & " - " & _
                        Trim(rst.Fields("PARTDESC"))
            Set nodX = tvw1.Nodes.Add(sENode, tvwChild, sPNode, sDesc, 4)
            
            rst.MoveNext
        Loop
    End If
    rst.Close: Set rst = Nothing
    
    If bWgtErr Then
        sDisclaim = "*" & vbNewLine & "* Unweighed Parts were found." & _
                    vbNewLine & "    Actual weight may be greater."
    End If
    Select Case sType
        Case "INIT"
            sMess = "Total Shipping Weight:    " & _
                        Format(dWgt, "#,##0") & wgtSuf & sDisclaim
        Case "KIT"
            sMess = "Total Shipping Weight of Kit:    " & _
                        Format(dWgt, "#,##0") & wgtSuf & sDisclaim
        Case "ALL"
            sMess = "Total Shipping Weight of Inventory:    " & _
                        Format(dWgt, "#,##0") & wgtSuf & sDisclaim
    End Select
    lblMess.Caption = sMess
    
    GetInven = sHDR
End Function

Public Function GetElement(strSelect As String)
    Dim rst As ADODB.Recordset
    Dim sENode As String, sPNode As String, sTNode As String
    Dim sDesc As String, sMess As String
    Dim nodX As Node
    Dim dVolT As Double, dWgtT As Double
    Dim bBadWgt As Boolean, bBadVol As Boolean
    Dim iWgtU As Integer, iVolU As Integer
    Dim sWgtU(1 To 2) As String
    
    sWgtU(1) = " lbs": sWgtU(2) = " kg"
    tvw1.ImageList = ImageList1
    
    sENode = "e" & pEID
    sDesc = pElem
    Set nodX = tvw1.Nodes.Add(, , sENode, sDesc, 3)
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        iWgtU = rst.Fields("WTUNIT")
        iVolU = rst.Fields("SIZEUNIT")
    End If
    Do While Not rst.EOF
        sPNode = "p" & rst.Fields("PARTID")
        sDesc = "Part:  " & UCase(Trim(rst.Fields("PARTDESC")))
        Set nodX = tvw1.Nodes.Add(sENode, tvwChild, sPNode, sDesc, 4)
            sTNode = "t" & rst.Fields("PARTID") & "-1"
            sDesc = "Part No:  " & Trim(rst.Fields("PNUM"))
            Set nodX = tvw1.Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 5)
            
            sTNode = "t" & rst.Fields("PARTID") & "-2"
            sDesc = "Pkg Type:  " & Trim(rst.Fields("PKGTYPE"))
            Set nodX = tvw1.Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 5)
            
            sTNode = "t" & rst.Fields("PARTID") & "-3"
            sDesc = "Part Size:  " & rst.Fields("LENGTH") & "L x " & _
                        rst.Fields("WIDTH") & "W x " & rst.Fields("HEIGHT") & "H"
            If rst.Fields("LENGTH") = 0 Or rst.Fields("WIDTH") = 0 _
                        Or rst.Fields("HEIGHT") = 0 Then
                bBadVol = True
            Else
                dVolT = dVolT + (rst.Fields("LENGTH") * rst.Fields("WIDTH") * rst.Fields("HEIGHT"))
            End If
            Set nodX = tvw1.Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 5)
            
            sTNode = "t" & rst.Fields("PARTID") & "-4"
            sDesc = "Part Wgt:  " & rst.Fields("WEIGHT") & sWgtU(rst.Fields("WTUNIT"))
            If rst.Fields("WEIGHT") = 0 Then
                bBadWgt = True
            Else
                dWgtT = dWgtT + rst.Fields("WEIGHT")
            End If
            Set nodX = tvw1.Nodes.Add(sPNode, tvwChild, sTNode, sDesc, 5)
            
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    sMess = ""
    Select Case iWgtU
        Case 1
            sMess = sMess & "Total Element Weight:  " & Format(dWgtT, "#,##0") & " lbs"
        Case 2
            sMess = sMess & "Total Element Weight:  " & Format(dWgtT, "#,##0") & " kg"
    End Select
    If bBadWgt Then
        sMess = sMess & "*" & vbNewLine
        sMess = sMess & "* Unweighed Parts found.  Actual weight may be greater."
    End If
    sMess = sMess & vbNewLine
    
    Select Case iVolU
        Case 1
            dVolT = dVolT / 1728
            sMess = sMess & "Total Element Volume:  " & Format(dVolT, "#,##0.00") & " cu ft"
        Case 5
            dVolT = dVolT / 1000000
            sMess = sMess & "Total Element Volume:  " & Format(dVolT, "#,##0.00") & " cu M"
    End Select
    If bBadVol Then
        sMess = sMess & "**" & vbNewLine
        sMess = sMess & "** Undimensioned Parts found.  Actual volume may be greater."
    End If
    
    lblMess.Caption = sMess
End Function
