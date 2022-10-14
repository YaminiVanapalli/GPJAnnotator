VERSION 5.00
Begin VB.Form frmAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphics Add Mode is Active..."
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
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
   Icon            =   "frmAssign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4170
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4380
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1395
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   4335
      Begin VB.ComboBox cboSHYR 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   3015
      End
      Begin VB.ComboBox cboSHCD 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   3915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Year:"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Left            =   900
         TabIndex        =   10
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show:"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Graphics to Attach to Show:"
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   1200
         Width           =   2505
      End
   End
   Begin VB.CommandButton cmdSaveList 
      Caption         =   " Save 'ADD' List To Show"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2115
   End
   Begin VB.CommandButton cmdRemoveSels 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to Remove Selected 'ADD's"
      Top             =   4380
      Width           =   1035
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Current Graphic To List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1755
   End
   Begin VB.ListBox lstGraphics 
      Height          =   2400
      ItemData        =   "frmAssign.frx":000C
      Left            =   120
      List            =   "frmAssign.frx":000E
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   1500
      Width           =   3915
   End
End
Attribute VB_Name = "frmAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tBCC As String, tFBCN As String, tSHNM As String
Private tSHYR As Integer
Private tSHCD As Long

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

Public Property Get PassSHNM() As String
    PassSHNM = tSHNM
End Property
Public Property Let PassSHNM(ByVal vNewValue As String)
    tSHNM = vNewValue
End Property

Public Property Get PassSHYR() As Integer
    PassSHYR = tSHYR
End Property
Public Property Let PassSHYR(ByVal vNewValue As Integer)
    tSHYR = vNewValue
End Property

Public Property Get PassSHCD() As Long
    PassSHCD = tSHCD
End Property
Public Property Let PassSHCD(ByVal vNewValue As Long)
    tSHCD = vNewValue
End Property

Private Sub cboCUNO_Click()
    If cboCUNO.Text <> "" Then
        tBCC = Right("00000000" & CStr(cboCUNO.ItemData(cboCUNO.ListIndex)), 8)
        tFBCN = GetBCN(tBCC)
        Call GetShows(cboSHCD, tSHYR, tBCC)
    End If
End Sub

Private Sub cboSHCD_Change()
    If cboSHCD.Text <> "" Then
        tSHCD = cboSHCD.ItemData(cboSHCD.ListIndex)
        tSHNM = GetSHNM(tSHCD, tSHYR)
        Call PopShowGraphics
    End If
End Sub

Private Sub cboSHCD_Click()
    If cboSHCD.Text <> "" Then
        tSHCD = cboSHCD.ItemData(cboSHCD.ListIndex)
        tSHNM = GetSHNM(tSHCD, tSHYR)
        Call PopShowGraphics
    End If
End Sub

Private Sub cboSHYR_Click()
    If cboSHYR.Text <> "" Then
        tSHYR = CInt(cboSHYR.Text)
        Call GetShowClients(cboCUNO, tSHYR)
    End If
End Sub

Private Sub cmdAdd_Click()
    lstGraphics.AddItem sCGDesc & " [ADD]"
    lstGraphics.ItemData(lstGraphics.NewIndex) = CLng(iCurrGType & lGID)
    cmdAdd.FontBold = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemoveSels_Click()
    Dim i As Integer
    For i = lstGraphics.ListCount - 1 To 0 Step -1
        If lstGraphics.Selected(i) And InStr(1, lstGraphics.List(i), "[ADD]") > 0 Then
            lstGraphics.RemoveItem (i)
        Else
            lstGraphics.Selected(i) = False
        End If
    Next i
End Sub

Private Sub cmdSaveList_Click()
    Dim rstL As ADODB.Recordset
    Dim strSelect As String, strInsert As String
    Dim i As Integer
    Dim tGID As Long, tShow_ID As Long
    Dim nodX As Node
    
    Err = 0
    On Error Resume Next
    Conn.BeginTrans
    For i = lstGraphics.ListCount - 1 To 0 Step -1
        If InStr(1, lstGraphics.List(i), "[ADD]") > 0 Then
            '///// FIRST, GET NEW FILE NAME \\\\\
            Set rstL = Conn.Execute("SELECT " & GFXSeq & ".NEXTVAL FROM DUAL")
            tShow_ID = rstL.Fields("nextval")
            rstL.Close
            Set rstL = Nothing
            
            tGID = CLng(Mid(lstGraphics.ItemData(i), 2))
            
            '///// NOW WRITE TO DATABASE \\\\\ ***NEED TO ADD STATUS TO ASSIGNMENT***
            strInsert = "INSERT INTO " & GFXShow & " " & _
                        "(GID, SHOW_ID, SHYR, AN8_SHCD, " & _
                        "AN8_CUNO, ADDUSER, ADDDTTM, UPDUSER, " & _
                        "UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & tGID & ", " & tShow_ID & ", " & tSHYR & ", " & tSHCD & ", " & _
                        "" & CLng(tBCC) & ", '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', " & _
                        "SYSDATE, 1)"
            Conn.Execute (strInsert)
            If Err Then
                Conn.RollbackTrans
                GoTo Problem
            End If
'''            If frmGraphics.cboCUNO(1).Text = cboCUNO.Text Then
'''                Set nodX = frmGraphics.tvwGraphics(1).Nodes.Add("s" & tSHCD, tvwChild, _
'''                            "ga" & tShow_ID, lstGraphics.List(i), CInt(Left(lstGraphics.ItemData(i), 1)))
'''                frmGraphics.tvwGraphics(1).Nodes("s" & tSHCD).Image = 5
'''                frmGraphics.tvwGraphics(1).Nodes("s" & tSHCD).Parent.Image = 5
'''            End If
        End If
    Next i
    Conn.CommitTrans
    
    '///// NOW, RUN POPSHOWGRAPHICS TO VERIFY ATTACHMENT \\\\\
    Call PopShowGraphics
    
Exit Sub
Problem:
    MsgBox "Error:  " & Err.Description, vbCritical, "File Assignment failed..."
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim iSHYR As Integer
    
    '///// POSITION MENU \\\\\
    Me.Left = frmGraphics.Left + 120 '' + frmGraphics.ScaleWidth - 30 - Me.Width
    Me.Top = frmGraphics.Top + frmGraphics.Height - 120 - Me.Height '' (frmGraphics.Height - frmGraphics.ScaleHeight) + 30
    
    '///// SETUP SHYR \\\\\
    iSHYR = CInt(Format(Now, "YYYY"))
    For i = -2 To 2
        cboSHYR.AddItem iSHYR + i
    Next i
    On Error Resume Next
    cboSHYR.Text = tSHYR
    If Err Then cboSHYR.Text = iSHYR
    Err.Clear
    
    cboCUNO.Text = tFBCN
    If Err = 0 Then
        cboSHCD.Text = tSHNM
    End If
End Sub

'''Private Sub Form_Unload(Cancel As Integer)
'''    frmGraphics.cmdAssign.Visible = True
'''End Sub

Public Sub PopShowGraphics()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim nodX As Node
    Dim sPNode As String, sGNode As String, sDesc As String, sDescPar As String
    Dim sKNode As String, sENode As String
    Dim lParent As Long, lElem As Long
    Dim iType As Integer
    
    lstGraphics.Clear
    strSelect = "SELECT GM.GDESC, GM.GTYPE, GS.SHOW_ID " & _
                "FROM " & GFXShow & " GS, " & GFXMas & " GM " & _
                "WHERE GS.SHYR = " & tSHYR & " " & _
                "AND GS.AN8_SHCD = " & tSHCD & " " & _
                "AND GS.AN8_CUNO = " & CLng(tBCC) & " " & _
                "AND GS.GID = GM.GID " & _
                "AND GM.GSTATUS > 0"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        '///// ADD TO ASSIGN LIST W/NO ITEMDATA \\\\\
        lstGraphics.AddItem Trim(rst.Fields("GDESC"))
        lstGraphics.ItemData(lstGraphics.NewIndex) = rst.Fields("SHOW_ID")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGraphics.bAddMode = False
End Sub
