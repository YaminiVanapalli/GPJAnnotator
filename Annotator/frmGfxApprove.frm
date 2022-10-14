VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGfxApprove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graphic Approval..."
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "frmGfxApprove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstApprover 
      Height          =   1635
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   30
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Frame fraNotification 
      Caption         =   "Notification"
      Height          =   5475
      Left            =   5340
      TabIndex        =   16
      Top             =   120
      Width           =   2835
      Begin VB.ListBox lstGPJEmail 
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstGPJShort 
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstTeamShort 
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstTeamEmail 
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstClientEmail 
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   3300
         Visible         =   0   'False
         Width           =   135
      End
      Begin TabDlg.SSTab sstEmail 
         Height          =   2955
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   5212
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Team"
         TabPicture(0)   =   "frmGfxApprove.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstTeam"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "GPJ"
         TabPicture(1)   =   "frmGfxApprove.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstGPJ"
         Tab(1).ControlCount=   1
         Begin VB.ListBox lstGPJ 
            Height          =   2535
            Left            =   -74940
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   360
            Width           =   2475
         End
         Begin VB.ListBox lstTeam 
            Height          =   2535
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   360
            Width           =   2475
         End
      End
      Begin VB.ListBox lstClientShort 
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   3300
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox lstTeamClients 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   3480
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Notification Team:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label lblClients 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clients Users:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   3240
         Width           =   990
      End
   End
   Begin VB.TextBox txtApprove 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3480
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5820
      Width           =   1035
   End
   Begin VB.OptionButton optGfxApprove 
      Caption         =   "Set Status as 'CLIENT DRAFT', allowing Client Viewing"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   4395
   End
   Begin VB.OptionButton optGfxApprove 
      Caption         =   "Reset Status to 'INTERNAL DRAFT' and restrict Client Viewing"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   4875
   End
   Begin VB.OptionButton optGfxApprove 
      Caption         =   "APPROVE Graphic"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   2100
      Width           =   1695
   End
   Begin VB.OptionButton optGfxApprove 
      Caption         =   "'DELETE' Graphic"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton cmdGfxApprove 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5820
      Width           =   1695
   End
   Begin VB.Frame fraApprove 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   5055
      Begin VB.OptionButton optApprove 
         Caption         =   "File APPROVED with associated Redlines"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1140
         Width           =   3315
      End
      Begin VB.OptionButton optApprove 
         Caption         =   "File APPROVED with attached comments"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   3315
      End
      Begin VB.OptionButton optApprove 
         Caption         =   "File APPROVED as shown"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.OptionButton optGfxApprove 
      Caption         =   "RETURN FOR CHANGES"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   1200
      Width           =   2115
   End
   Begin VB.Frame fraNotApprove 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   5055
      Begin VB.OptionButton optNotApprove 
         Caption         =   "File NOT APPROVED with attached comments"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   3795
      End
      Begin VB.OptionButton optNotApprove 
         Caption         =   "File NOT APPROVED with associated Redlines"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   3795
      End
   End
   Begin VB.Label lblApprover 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Graphic Approver:"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   1785
   End
   Begin VB.Image img1 
      Height          =   480
      Index           =   4
      Left            =   9180
      Picture         =   "frmGfxApprove.frx":0902
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label lblComm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment (to be included with Automated Email):"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Image img1 
      Height          =   480
      Index           =   2
      Left            =   9180
      Picture         =   "frmGfxApprove.frx":11CC
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image img1 
      Height          =   480
      Index           =   1
      Left            =   9180
      Picture         =   "frmGfxApprove.frx":14D6
      Top             =   720
      Width           =   480
   End
   Begin VB.Image img1 
      Height          =   480
      Index           =   0
      Left            =   9180
      Picture         =   "frmGfxApprove.frx":17E0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmGfxApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sApprover As String
Dim iNewStatus As Integer, iChecked As Integer
Dim X1 As Long, Y1 As Long, pBCC As Long
Dim iVal As Integer
Dim sHDR As String, tFBCN As String, pType As String
Dim lAID_Current As Long, lAID_New As Long
Dim bClearingList As Boolean


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

Public Property Get PassVal() As Integer
    PassVal = iVal
End Property
Public Property Let PassVal(ByVal vNewValue As Integer)
    iVal = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = sHDR
End Property

Public Property Let PassHDR(ByVal vNewValue As String)
    sHDR = vNewValue
End Property

Public Property Get PassType() As String
    PassType = pType
End Property

Public Property Let PassType(ByVal vNewValue As String)
    pType = vNewValue
End Property

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property

Public Property Get PassFBCN() As String
    PassFBCN = tFBCN
End Property

Public Property Let PassFBCN(ByVal vNewValue As String)
    tFBCN = vNewValue
End Property



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGfxApprove_Click()
    Dim i As Integer
    Dim iStatus(0 To 4) As Integer
    Dim strUpdate As String, strSelect As String, strInsert As String, sComm As String
    Dim sNodeKey As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim GIDList As String
    Dim lCOMMID As Long
    Dim iErr As Integer
    Dim Resp As VbMsgBoxResult
    Dim sNote As String
    Dim bBeginTrans As Boolean
    
    Select Case iNewStatus
        Case 2 ''APPROVE''
            If optApprove(1).Value = True Then
                If txtApprove.Text = "" Then
                    MsgBox "There is no entered Comment", vbExclamation, "Approved with Attached Comments?"
                    Exit Sub
                End If
            End If
        Case 4 ''RETURN FOR CHANGED''
            If txtApprove.Text = "" Then
                MsgBox "There is no entered Comment", vbExclamation, "NOT Approved with Attached Comments?"
                Exit Sub
            End If
    End Select
    
'''    iNewStatus(0) = 5
'''    iNewStatus(1) = 15
'''    iNewStatus(2) = 25
    iStatus(0) = 10
    iStatus(1) = 20
    iStatus(2) = 30
    iStatus(4) = 27
    
    Select Case iVal
        Case 0: iStatus(3) = 2
        Case 1: iStatus(3) = 3
        Case 2: iStatus(3) = 4
    End Select
    
    If iNewStatus = 3 Then ''canceling''
        Select Case Me.Tag
            Case "ONE", "VWR"
                Resp = MsgBox("Are you certain you want to delete this file?", _
                            vbYesNo, "Confirming...")
            Case "VIEW"
                Resp = MsgBox("Are you certain you want to Delete these files?", _
                            vbYesNo, "Confirming...")
        End Select
        If Resp = vbNo Then Exit Sub
    End If
    
    Conn.BeginTrans: bBeginTrans = True
    On Error GoTo ErrorTrap
    Select Case Me.Tag
        Case "ONE", "VWR"
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = " & iStatus(iNewStatus) & ", " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                        "WHERE GID = " & StatusGID
            Conn.Execute (strUpdate)
            
            sComm = txtApprove.Text
            iErr = InsertComment(StatusGID, iNewStatus, sComm)
            If iErr > 0 Then GoTo ErrorTrap
            
            If Me.Tag = "VWR" Then
                ''RESET ON-SCREEN LABELS''
                Select Case iNewStatus
                    Case 0
                        frmGraphics.lblStatus.Caption = "STATUS:  INTERNAL (Last Status Update " & _
                                    UCase(Format(Now, "mmmm d, yyyy"))
                        frmGraphics.lblGraphic.Caption = "INTERNAL DRAFT"
                    Case 1
                        frmGraphics.lblStatus.Caption = "STATUS:  CLIENT DRAFT (Last Status Update " & _
                                    UCase(Format(Now, "mmmm d, yyyy"))
                        frmGraphics.lblGraphic.Caption = "CLIENT DRAFT"
                    Case 2
                        frmGraphics.lblStatus.Caption = "STATUS:  APPROVED (Last Status Update " & _
                                    UCase(Format(Now, "mmmm d, yyyy"))
                        frmGraphics.lblGraphic.Caption = "APPROVED"
                    Case 3
                        frmGraphics.lblStatus.Caption = "STATUS:  CANCELED (Last Status Update " & _
                                    UCase(Format(Now, "mmmm d, yyyy"))
                        frmGraphics.lblGraphic.Caption = "CANCELED"
                    Case 4
                        frmGraphics.lblStatus.Caption = "STATUS:  RETURNED FOR CHANGES (Last Status Update " & _
                                    UCase(Format(Now, "mmmm d, yyyy"))
                        frmGraphics.lblGraphic.Caption = "RETURNED FOR CHANGES"
                End Select
                
            End If
            
            If lAID_New <> lAID_Current And iNewStatus <> 2 Then ''APPROVER HAS BEEN CHANGED''
'''                Dim iErr As Integer
                
                
                strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                            "GAPPROVER_ID = " & lAID_New & ", " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                            "WHERE GID = " & StatusGID
                Conn.Execute (strUpdate)
                
                frmGraphics.flxApprove.Row = frmGraphics.iApprovalRow
                frmGraphics.flxApprove.Col = 6
                Set frmGraphics.flxApprove.CellPicture = LoadPicture("")
                frmGraphics.flxApprove.TextMatrix(frmGraphics.iApprovalRow, 6) _
                            = StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase)
                
                sNote = "Graphic Approver set to " & StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase) & _
                            " by " & StrConv(LogName, vbProperCase)
                iErr = InsertGfxComment(StatusGID, sNote)
                
            End If
            
            GIDList = CStr(StatusGID)
            StatusGID = 0
            
        Case "VIEW"
            GIDList = ""
            For i = 1 To frmGraphics.flxApprove.Rows - 1
                If GIDList = "" Then
                    GIDList = frmGraphics.flxApprove.TextMatrix(i, 0)
                Else
                    GIDList = GIDList & ", " & frmGraphics.flxApprove.TextMatrix(i, 0)
                End If
                
                sComm = txtApprove.Text
                iErr = InsertComment(frmGraphics.flxApprove.TextMatrix(i, 0), iNewStatus, sComm)
                If iErr > 0 Then GoTo ErrorTrap
                
                strUpdate = "UPDATE ANNOTATOR.GFX_MASTER SET " & _
                            "GAPPROVER_ID = " & lAID_New & ", " & _
                            "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                            "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                            "WHERE GID = " & frmGraphics.flxApprove.TextMatrix(i, 0)
                Conn.Execute (strUpdate)
                
                frmGraphics.flxApprove.Row = i
                frmGraphics.flxApprove.Col = 6
                Set frmGraphics.flxApprove.CellPicture = LoadPicture("")
                frmGraphics.flxApprove.TextMatrix(i, 6) _
                            = StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase)
                
                sNote = "Graphic Approver set to " & StrConv(lstApprover.List(lstApprover.ListIndex), vbProperCase) & _
                            " by " & StrConv(LogName, vbProperCase)
                iErr = InsertGfxComment(frmGraphics.flxApprove.TextMatrix(i, 0), sNote)
                
            Next i
            
            strUpdate = "UPDATE " & GFXMas & " " & _
                        "SET GSTATUS = " & iStatus(iNewStatus) & ", " & _
                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT +1 " & _
                        "WHERE GID IN (" & GIDList & ")"
            Conn.Execute (strUpdate)
            
            
    End Select
    Conn.CommitTrans
    
    ''COMMENTED OUT FOR NOW.  IF USED, SendGPFEmail NEEDS TO BE REENABLED''
    ''Call SendGFXEmail(pBCC, tFBCN, GIDList, sComm, iStatus(iNewStatus))
    
    GIDList = ""
    
'''    frmGraphics.cboCUNO(4).Text = frmGraphics.cboCUNO(4).Text
    Call frmGraphics.cmdRefresh_Click
    
    Unload Me
    

Exit Sub
ErrorTrap:
    On Error Resume Next
    If bBeginTrans Then Conn.RollbackTrans
    MsgBox "Error Encountered during Status Change." & vbNewLine & vbNewLine & _
                "Error:  " & Err.Description, vbCritical, "Status Change Aborted..."
    Err.Clear

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
'    If X1 <> 0 And Y1 <> 0 Then
'        Me.Left = X1
'        Me.Top = Y1
'    Else
'        Me.Left = (Screen.Width - Me.Width) / 2
'        Me.Top = (Screen.Height - Me.Height) / 2
'    End If
    Me.Tag = pType
    
    If sHDR <> "" Then Me.Caption = sHDR ''' lblGfxApprove.Caption = sHDR
    
    If Not bGPJ Then
        optGfxApprove(0).Visible = False
        optGfxApprove(3).Visible = False
    End If
    
'''    optGfxApprove(iVal).value = True
'''    Set Me.Icon = img1(iVal).Picture
    Select Case iVal
        Case 0
            optGfxApprove(iVal).Value = True
            Set Me.Icon = img1(iVal).Picture
            optGfxApprove(iVal).Caption = "Current Status is 'INTERNAL DRAFT', restricting Client Viewing"
            optGfxApprove(iVal).BackColor = vbRed
        Case 1
            optGfxApprove(iVal).Value = True
            Set Me.Icon = img1(iVal).Picture
            optGfxApprove(iVal).Caption = "Current Status is 'CLIENT DRAFT', allowing Client Viewing"
            optGfxApprove(iVal).BackColor = vbYellow
        Case 2
            optGfxApprove(iVal).Value = True
            Set Me.Icon = img1(iVal).Picture
            optGfxApprove(iVal).BackColor = vbGreen
            fraApprove.Enabled = False
        Case 4
            optGfxApprove(iVal).Value = True
            Set Me.Icon = img1(iVal).Picture
            optGfxApprove(iVal).BackColor = RGB(64, 128, 255) ''(255, 127, 0)
            fraNotApprove.Enabled = False
    End Select
    iNewStatus = iVal
    
    Call GetTeam(pBCC)
    
    Select Case pType
        Case "VIEW"
            Call GetApprover("ALL")
            lAID_Current = 0
        Case Else
            Call GetApprover("ONE")
            SetApproverID (StatusGID)
    End Select
    
    Call UnselectSelf(LogName, Me.lstTeam)
    
    '///// GET GPJ PERSONNEL \\\\\
    If bGPJ Then
        sstEmail.TabVisible(1) = True
        Call GetGPJEmail
    Else
        sstEmail.TabVisible(1) = False
    End If
    
    
    
    
        
End Sub

Private Sub lstApprover_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If bClearingList Then
        iChecked = Item
        lAID_New = lstApprover.ItemData(iChecked)
        Debug.Print "lAID_Current = " & lAID_Current & ": lAID_New = " & lAID_New
        bClearingList = False
    End If
    For i = 0 To lstApprover.ListCount - 1
        If i <> iChecked Then lstApprover.Selected(i) = False
    Next i
    
    Call SyncLists
    
'    lstApprover.Visible = False
'
'    lstApprover.Selected(Item) = True
'    lstApprover.Visible = True
    
End Sub

Private Sub lstApprover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    bClearingList = True
''''    For i = 0 To lstApprover.ListCount - 1
''''        lstApprover.Selected(i) = False
''''    Next i
End Sub

Private Sub lstTeam_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If lstTeam.ItemData(Item) = 1 And lstTeam.Selected(Item) = False Then
        lstTeam.Selected(Item) = True
        MsgBox "The person selected is a Mandatory Recipient.", _
                    vbExclamation, "Member cannot be Unchecked..."
    End If
End Sub

'''Private Sub lstTeamClients_ItemCheck(Item As Integer)
'''    Dim i As Integer
''''    Debug.Print lstTeamClients.Selected(Item)
'''    If lstTeamClients.ItemData(Item) = 1 And lstTeamClients.Selected(Item) = False Then
'''        lstTeamClients.Selected(Item) = True
'''        MsgBox "The Client selected is a Graphic Approver.", _
'''                    vbExclamation, "Member cannot be Unchecked..."
'''    End If
'''End Sub

Private Sub optApprove_Click(Index As Integer)
    If optApprove(0).Value = True Then
        cmdGfxApprove.Enabled = True
    ElseIf optApprove(1).Value = True Then
'        If txtApprove.Text <> "" Then
            cmdGfxApprove.Enabled = True
'        Else
'            cmdGfxApprove.Enabled = False
'        End If
    ElseIf optApprove(2).Value = True Then
        cmdGfxApprove.Enabled = True
    Else
        cmdGfxApprove.Enabled = False
    End If
End Sub

'''Private Sub imgGfxApprove_Click()
'''    Unload Me
'''End Sub

Private Sub optGfxApprove_Click(Index As Integer)
    iNewStatus = Index
    
    If Index = 0 Then
        lblClients.Visible = False
        lstTeamClients.Visible = False
    Else
        lblClients.Visible = True
        lstTeamClients.Visible = True
    End If
    
    If Index <> iVal Then
        cmdGfxApprove.Enabled = True
    Else
        cmdGfxApprove.Enabled = False
    End If
    
    If optGfxApprove(2).Value = True Then
        fraApprove.Enabled = True
        If optApprove(0).Value = True Then
            cmdGfxApprove.Enabled = True
        ElseIf optApprove(1).Value = True Then
'            If txtApprove.Text <> "" Then
                cmdGfxApprove.Enabled = True
'            Else
'                cmdGfxApprove.Enabled = False
'            End If
        ElseIf optApprove(2).Value = True Then
            cmdGfxApprove.Enabled = True
        Else
            cmdGfxApprove.Enabled = False
        End If
    Else
        fraApprove.Enabled = False
    End If
    
    If optGfxApprove(4).Value = True Then
        fraNotApprove.Enabled = True
        If optNotApprove(1).Value = True Then
'''            If txtApprove.Text <> "" Then
                cmdGfxApprove.Enabled = True
'''            Else
'''                cmdGfxApprove.Enabled = False
'''            End If
        ElseIf optApprove(2).Value = True Then
            cmdGfxApprove.Enabled = True
        Else
            cmdGfxApprove.Enabled = False
        End If
    Else
        fraNotApprove.Enabled = False
    End If
    
    ''RESET APPROVER LIST''
    Call ResetApproverList(pBCC, iNewStatus)
    
End Sub

Private Sub optNotApprove_Click(Index As Integer)
    If optNotApprove(1).Value = True Then
'        If txtApprove.Text <> "" Then
            cmdGfxApprove.Enabled = True
'        Else
'            cmdGfxApprove.Enabled = False
'        End If
    ElseIf optNotApprove(2).Value = True Then
        cmdGfxApprove.Enabled = True
    Else
        cmdGfxApprove.Enabled = False
    End If
End Sub


Public Sub GetTeam(tBCC As Long)
    Dim strSelect As String, sEmployer As String
    Dim rst As ADODB.Recordset
    
        
    '/// 0=Floorplan,1=Graphics,2=Const Dwg \\\
    
    '///// FIRST, GET TEAM \\\\\
    '///// SEE IF CLIENT-SHOW TEAM EXISTS \\\\\
    strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS, " & _
                "U.NAME_LOGON, U.EMPLOYER, R.RECIPIENT_FLAG1 " & _
                "FROM " & ANOETeamUR & " R, " & IGLUser & " U, " & IGLUserAR & " AR " & _
                "WHERE R.TEAM_ID IN (" & _
                "SELECT TEAM_ID FROM " & ANOETeam & " " & _
                "WHERE AN8_CUNO = " & tBCC & ") " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "AND U.USER_SEQ_ID = AR.USER_SEQ_ID " & _
                "AND AR.APP_ID = 1002 " & _
                "AND AR.PERMISSION_STATUS > 0 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    lstTeam.Clear: lstTeamEmail.Clear
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) = "GPJ" Then
        
'''        If Left(rst.Fields("EMPLOYER"), 3) <> "GPJ" Then
'''            sEmployer = " (" & Trim(rst.Fields("EMPLOYER")) & ")"
'''        Else
            sEmployer = ""
'''        End If
''''''        If StrConv(Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")), vbProperCase) = LogName Then
''''''            bTeamMember = True
''''''            If rst.Fields("RECIPIENT_FLAG1") = 1 Then bMandRecip = True
''''''        End If
            lstTeam.AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & _
                        UCase(Trim(rst.Fields("NAME_LAST"))) & sEmployer
            lstTeam.ItemData(lstTeam.NewIndex) = rst.Fields("RECIPIENT_FLAG1")
            lstTeam.Selected(lstTeam.NewIndex) = CBool(rst.Fields("RECIPIENT_FLAG1") * -1)
            lstTeamEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
            lstTeamShort.AddItem Trim(rst.Fields("NAME_LOGON"))
            
'            lstApprover.AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & _
'                        UCase(Trim(rst.Fields("NAME_LAST")))
'            lstApprover.ItemData(lstApprover.NewIndex) = rst.Fields("USER_SEQ_ID")
            
        End If
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    
'''    strSelect = "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, " & _
'''                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME " & _
'''                "FROM IGL_USER_CUNO_R CR, IGL_CUNO_GROUP_R GR, " & _
'''                "IGL_USER US, IGL_USER_APP_R AR " & _
'''                "Where GR.AN8_CUNO = " & tBCC & " " & _
'''                "AND GR.CUNO_GROUP_ID = CR.CUNO_GROUP_ID " & _
'''                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
'''                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
'''                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
'''                "AND AR.APP_ID = 1002 " & _
'''                "ORDER BY US.NAME_LAST"
                
    strSelect = "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, US.NAME_LAST, US.NAME_LOGON, " & _
                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME, US.EMPLOYER " & _
                "FROM IGLPROD.IGL_USER_CUNO_R CR, IGLPROD.IGL_CUNO_GROUP_R GR, IGLPROD.IGL_USER US, IGLPROD.IGL_USER_APP_R AR " & _
                "Where GR.AN8_CUNO = " & tBCC & " " & _
                "AND GR.CUNO_GROUP_ID = CR.CUNO_GROUP_ID " & _
                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
                "AND US.USER_STATUS > 0 " & _
                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
                "AND AR.APP_ID = 1002 " & _
                "AND AR.PERMISSION_STATUS > 0 " & _
                "Union " & _
                "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, US.NAME_LAST, US.NAME_LOGON, " & _
                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME, US.EMPLOYER " & _
                "FROM IGLPROD.IGL_USER_CUNO_R CR, IGLPROD.IGL_USER US, IGLPROD.IGL_USER_APP_R AR " & _
                "Where CR.AN8_CUNO = " & tBCC & " " & _
                "AND CR.CUNO_GROUP_ID = 0 " & _
                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
                "AND US.USER_STATUS > 0 " & _
                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
                "AND AR.APP_ID = 1002 " & _
                "AND AR.PERMISSION_STATUS > 0 " & _
                "ORDER BY NAME_LAST"
    Set rst = Conn.Execute(strSelect)
    lstTeamClients.Clear: lstClientEmail.Clear
    Do While Not rst.EOF
        lstTeamClients.AddItem UCase(Trim(rst.Fields("FULLNAME")))
        lstClientEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstClientShort.AddItem Trim(rst.Fields("NAME_LOGON"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    
'''    If bMandRecip Then
'''        chkReceive.value = 1
'''        chkReceive.value = 2
'''        chkReceive.Enabled = False
'''    ElseIf bTeamMember Then
'''        chkReceive.value = 1
'''        chkReceive.Enabled = True
'''    Else
'''        chkReceive.value = 0
'''        chkReceive.Enabled = True
'''    End If
'''    chkReceive.Tag = LogAddress

End Sub

'''Public Sub GetRecips(tStatus As Integer)
'''    Dim iAdd As Integer, i As Integer
'''
'''    sCC = ""
'''    iAdd = -1
'''    For i = 0 To lstTeam.ListCount - 1
'''        If lstTeam.Selected(i) = True Then
'''            iAdd = iAdd + 1
'''            ReDim Preserve GFXAddress(iAdd)
'''            ReDim Preserve GFXMandRecip(iAdd)
'''            GFXAddress(iAdd) = lstTeamEmail.List(i)
'''            GFXMandRecip(iAdd) = lstTeam.List(i)
'''            sCC = sCC & vbTab & lstTeam.List(i) & vbNewLine
'''        End If
'''    Next i
'''
'''    If tStatus > 10 Then
'''        For i = 0 To lstTeam.ListCount - 1
'''            If lstTeamClients.Selected(i) = True Then
'''                iAdd = iAdd + 1
'''                ReDim Preserve GFXAddress(iAdd)
'''                ReDim Preserve GFXMandRecip(iAdd)
'''                GFXAddress(iAdd) = lstClientEmail.List(i)
'''                GFXMandRecip(iAdd) = lstTeamClients.List(i)
'''                sCC = sCC & vbTab & lstTeamClients.List(i) & vbNewLine
'''            End If
'''        Next i
'''    End If
'''
'''End Sub



Public Sub GetApprover(tType As String)
    Dim i As Integer, iList As Integer
    Dim sCheck As String
    
    Select Case tType
        Case "ALL"
            For i = 1 To frmGraphics.flxApprove.Rows - 1
                If frmGraphics.flxApprove.TextMatrix(i, 6) <> "" Then
                    sCheck = UCase(frmGraphics.flxApprove.TextMatrix(i, 6))
                    For iList = 0 To lstTeam.ListCount - 1
                        If sCheck = UCase(lstTeam.List(iList)) _
                                    And lstTeam.Selected(iList) = False Then
                            lstTeam.Selected(iList) = True
                            lstTeam.ItemData(iList) = 1
                            GoTo FoundIt1
                        End If
                    Next iList
                    For iList = 0 To lstTeamClients.ListCount - 1
                        If sCheck = UCase(lstTeamClients.List(iList)) _
                                    And lstTeamClients.Selected(iList) = False Then
                            lstTeamClients.Selected(iList) = True
                            lstTeamClients.ItemData(iList) = 1
                            GoTo FoundIt1
                        End If
                    Next iList
FoundIt1:
                End If
            Next i
        Case Else
            sCheck = UCase(frmGraphics.flxApprove.TextMatrix(frmGraphics.CurrIndex + 1, 6))
            If sCheck <> "" Then
                For iList = 0 To lstTeam.ListCount - 1
                    If sCheck = UCase(lstTeam.List(iList)) _
                                And lstTeam.Selected(iList) = False Then
                        lstTeam.Selected(iList) = True
                        lstTeam.ItemData(iList) = 1
                        GoTo FoundIt2
                    End If
                Next iList
                For iList = 0 To lstTeamClients.ListCount - 1
                    If sCheck = UCase(lstTeamClients.List(iList)) _
                                And lstTeamClients.Selected(iList) = False Then
                        lstTeamClients.Selected(iList) = True
                        lstTeamClients.ItemData(iList) = 1
                        GoTo FoundIt2
                    End If
                Next iList
            End If
FoundIt2:
    End Select
                            
End Sub

Public Sub GetGPJEmail()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    lstGPJ.Clear: lstGPJEmail.Clear: lstGPJShort.Clear
    '///// POP GPJ EMPLOYEE LIST \\\\\
    strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LOGON, " & _
                "U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
                "FROM " & IGLUser & " U, " & IGLUserAR & " R " & _
                "WHERE U.USER_STATUS = 1 " & _
                "AND SUBSTR(U.EMPLOYER, 1, 3) = 'GPJ' " & _
                "AND U.EMAIL_ADDRESS IS NOT NULL " & _
                "AND U.USER_SEQ_ID = R.USER_SEQ_ID " & _
                "AND R.APP_ID = 1002 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstGPJ.AddItem Trim(rst.Fields("NAME_FIRST")) & " " & _
                    Trim(rst.Fields("NAME_LAST"))
        lstGPJEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        lstGPJShort.AddItem LCase(Trim(rst.Fields("NAME_LOGON")))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub



Public Sub ResetApproverList(tBCC As Long, tStatus As Integer)
    Dim strSelect As String, sEmployer As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    lstApprover.Clear
    Select Case tStatus
        Case 2, 3 ''APPROVE OR DELETE'
            Exit Sub
        Case 0 ''INTERNAL DRAFT - GPJ ONLY''
            strSelect = "SELECT U.USER_SEQ_ID, (TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS FULLNAME, " & _
                        "U.NAME_LAST , U.EMPLOYER " & _
                        "FROM ANNOTATOR.ANO_EMAIL_TEAM T, ANNOTATOR.ANO_EMAIL_TEAM_USER_R R, IGLPROD.IGL_USER U " & _
                        "Where T.AN8_CUNO = " & tBCC & " " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "AND U.EMPLOYER LIKE 'GPJ%' " & _
                        "ORDER BY NAME_LAST"
        Case Else ''GPJ TEAM & CLIENT APPROVERS''
            strSelect = "SELECT U.USER_SEQ_ID, (TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS FULLNAME, " & _
                        "U.NAME_LAST , U.EMPLOYER " & _
                        "FROM ANNOTATOR.ANO_EMAIL_TEAM T, ANNOTATOR.ANO_EMAIL_TEAM_USER_R R, IGLPROD.IGL_USER U " & _
                        "Where T.AN8_CUNO = " & tBCC & " " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "AND U.EMPLOYER LIKE 'GPJ%' "
            strSelect = strSelect & "Union " & _
                        "SELECT U.USER_SEQ_ID, (TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS FULLNAME, " & _
                        "U.NAME_LAST , U.EMPLOYER " & _
                        "FROM ANNOTATOR.ANO_EMAIL_TEAM T, ANNOTATOR.ANO_EMAIL_TEAM_USER_R R, IGLPROD.IGL_USER U " & _
                        "Where T.AN8_CUNO = " & tBCC & " " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.EXTCLIENTAPPROVER_FLAG = 1 " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "ORDER BY NAME_LAST"
    End Select
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If Left(rst.Fields("EMPLOYER"), 3) = "GPJ" Then
            sEmployer = ""
        Else
            sEmployer = "  (" & Trim(rst.Fields("EMPLOYER")) & ")"
        End If
        lstApprover.AddItem Trim(rst.Fields("FULLNAME")) ''& sEmployer
        lstApprover.ItemData(lstApprover.NewIndex) = rst.Fields("USER_SEQ_ID")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
    For i = 0 To lstApprover.ListCount - 1
        If lstApprover.ItemData(i) = lAID_Current Then
            bClearingList = True
            lstApprover.Selected(i) = True
            Exit For
        End If
    Next i
    
End Sub

Public Sub SetApproverID(tGID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    strSelect = "SELECT GAPPROVER_ID " & _
                "FROM ANNOTATOR.GFX_MASTER " & _
                "WHERE GID = " & tGID
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        lAID_Current = rst.Fields("GAPPROVER_ID")
    Else
        lAID_Current = 0
    End If
    rst.Close: Set rst = Nothing
    
    For i = 0 To lstApprover.ListCount - 1
        If lstApprover.ItemData(i) = lAID_Current Then
            bClearingList = True
            lstApprover.Selected(i) = True
            lblApprover.Caption = "Reset Graphic Approver (" & lstApprover.List(i) & "):"
            Exit For
        End If
    Next i
    
End Sub

Public Sub SyncLists()
    Dim i As Integer

    If sApprover <> "" Then
        ''ATTEMPT TO CLEAR FROM LISTS''
        For i = 0 To lstTeam.ListCount - 1
            If sApprover = lstTeam.List(i) Then
                If lstTeam.ItemData(i) = 0 Then
                    lstTeam.Selected(i) = False
                End If
                GoTo FoundIt
            End If
        Next i
        For i = 0 To lstTeamClients.ListCount - 1
            If sApprover = lstTeamClients.List(i) Then
                If lstTeamClients.ItemData(i) = 0 Then
                    lstTeamClients.Selected(i) = False
                End If
                GoTo FoundIt
            End If
        Next i
FoundIt:
    End If
        
    ''NOW CHECK FOR NEW APPROVER''
    sApprover = ""
    For i = 0 To lstApprover.ListCount - 1
        If lstApprover.Selected(i) Then
            sApprover = lstApprover.List(i)
            Exit For
        End If
    Next i
    
    If sApprover <> "" Then
        ''PLACE CHECK IN NEW APPROVER'S NAME''
        For i = 0 To lstTeam.ListCount - 1
            If lstTeam.List(i) = sApprover Then
                lstTeam.Selected(i) = True
                GoTo CheckedIt
            End If
        Next i
        For i = 0 To lstTeamClients.ListCount - 1
            If lstTeamClients.List(i) = sApprover Then
                lstTeamClients.Selected(i) = True
                GoTo CheckedIt
            End If
        Next i
CheckedIt:
    End If
    
End Sub
