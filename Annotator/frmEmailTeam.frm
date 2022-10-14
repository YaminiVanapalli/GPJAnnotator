VERSION 5.00
Begin VB.Form frmEmailTeam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Notification Team Administration"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmailTeam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Client Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   8115
      Left            =   6900
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4275
      Begin VB.ComboBox cboGroups 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   600
         Width           =   3915
      End
      Begin VB.ListBox lstClients 
         Height          =   5460
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   1680
         Width           =   3915
      End
      Begin VB.CommandButton cmdCloseClients 
         Caption         =   "Close Team Replicator"
         Height          =   615
         Left            =   2220
         TabIndex        =   2
         Top             =   7380
         Width           =   1875
      End
      Begin VB.CommandButton cmdSaveClients 
         Caption         =   "Replicate Teams"
         Height          =   615
         Left            =   180
         TabIndex        =   1
         Top             =   7380
         Width           =   1875
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--OR --"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   1875
         TabIndex        =   39
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Clients to replicate to by using IGL Groups"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   420
         TabIndex        =   38
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Highlight Clients from list below, that you would like the Current Team associated with."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   420
         Left            =   165
         TabIndex        =   5
         Top             =   1200
         Width           =   3945
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picRepTeamHelp 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      ScaleHeight     =   375
      ScaleMode       =   0  'User
      ScaleWidth      =   375
      TabIndex        =   36
      Top             =   7620
      Width           =   375
      Begin VB.Image imgRepTeamHelp 
         Height          =   375
         Left            =   0
         MousePointer    =   99  'Custom
         Picture         =   "frmEmailTeam.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "View Replicate Team QuickHelp"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5130
      TabIndex        =   35
      Top             =   6900
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "Team Type"
      Height          =   1515
      Left            =   180
      TabIndex        =   23
      Top             =   120
      Width           =   5775
      Begin VB.Frame fraSHCD 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         TabIndex        =   28
         Top             =   1020
         Visible         =   0   'False
         Width           =   5655
         Begin VB.ComboBox cboSHCD 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   60
            Width           =   4395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Show:"
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.OptionButton optTeamType 
         Caption         =   "Project-based"
         Height          =   255
         Index           =   2
         Left            =   4260
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTeamType 
         Caption         =   "Client/Show-based"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   1995
      End
      Begin VB.OptionButton optTeamType 
         Caption         =   "Client-based (Typ)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.ComboBox cboCUNO 
         Height          =   315
         ItemData        =   "frmEmailTeam.frx":0614
         Left            =   1200
         List            =   "frmEmailTeam.frx":0616
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   660
         Width           =   4395
      End
      Begin VB.Frame fraCPRJ 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         TabIndex        =   31
         Top             =   1020
         Visible         =   0   'False
         Width           =   5595
         Begin VB.ComboBox cboCPRJ 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   60
            Width           =   4395
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Project:"
            Height          =   195
            Left            =   60
            TabIndex        =   33
            Top             =   120
            Width           =   1050
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Client:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team Members"
      Enabled         =   0   'False
      Height          =   6495
      Left            =   180
      TabIndex        =   11
      Top             =   1740
      Width           =   10995
      Begin VB.Frame fraClientApprovers 
         Caption         =   "External Clients w/Graphic Approval Rights"
         Height          =   1455
         Left            =   6720
         TabIndex        =   50
         Top             =   4860
         Width           =   4095
         Begin VB.ListBox lstClientApprovers 
            ForeColor       =   &H80000007&
            Height          =   1035
            ItemData        =   "frmEmailTeam.frx":0618
            Left            =   180
            List            =   "frmEmailTeam.frx":061A
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mandatory Recipients"
         Height          =   4635
         Left            =   6720
         TabIndex        =   40
         Top             =   180
         Width           =   4095
         Begin VB.ListBox lstMandRips 
            ForeColor       =   &H80000007&
            Height          =   1035
            Index           =   2
            ItemData        =   "frmEmailTeam.frx":061C
            Left            =   180
            List            =   "frmEmailTeam.frx":061E
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   3420
            Width           =   3735
         End
         Begin VB.ListBox lstMandRips 
            ForeColor       =   &H80000007&
            Height          =   1230
            Index           =   1
            ItemData        =   "frmEmailTeam.frx":0620
            Left            =   180
            List            =   "frmEmailTeam.frx":0622
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3735
         End
         Begin VB.ListBox lstMandRips 
            ForeColor       =   &H80000007&
            Height          =   1230
            Index           =   0
            ItemData        =   "frmEmailTeam.frx":0624
            Left            =   180
            List            =   "frmEmailTeam.frx":0626
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   420
            Width           =   3735
         End
         Begin VB.Label lblMandRips 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Construction Drawing Recipients"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   46
            Top             =   3180
            Width           =   2325
         End
         Begin VB.Label lblMandRips 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Graphic File Recipients"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   45
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label lblMandRips 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Floorplan Recipients"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   44
            Top             =   210
            Width           =   1440
         End
      End
      Begin VB.ListBox lstNonGPJ 
         ForeColor       =   &H80000007&
         Height          =   1815
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   4560
         Width           =   3075
      End
      Begin VB.TextBox txtGetName 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3075
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Team"
         Height          =   495
         Left            =   3540
         TabIndex        =   16
         Top             =   5160
         Width           =   1275
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   795
         Left            =   5040
         Picture         =   "frmEmailTeam.frx":0628
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to Team"
         Height          =   795
         Left            =   3540
         Picture         =   "frmEmailTeam.frx":0932
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1215
      End
      Begin VB.ListBox lstTeam 
         ForeColor       =   &H80000007&
         Height          =   2985
         Left            =   3360
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3075
      End
      Begin VB.ListBox lstGPJEmps 
         ForeColor       =   &H80000007&
         Height          =   2985
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   1260
         Width           =   3075
      End
      Begin VB.CommandButton cmdReplicateTeam 
         Caption         =   "          Replicate Team"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5760
         Width           =   2715
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Team List"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   4545
         TabIndex        =   49
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label lblInstruct2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   855
         Left            =   3360
         TabIndex        =   48
         Top             =   4260
         Width           =   3075
      End
      Begin VB.Label lblDrag 
         Height          =   255
         Left            =   2700
         TabIndex        =   47
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Or Select from Non-GPJ List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   630
         TabIndex        =   22
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Or Select from Employee List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   210
         Left            =   630
         TabIndex        =   20
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Type in GPJ Employee Name..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3075
      End
   End
   Begin VB.ListBox lstGPJEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstNonGPJEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3420
      TabIndex        =   9
      Top             =   8700
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstTeamEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1980
      TabIndex        =   8
      Top             =   8580
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Top             =   8700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstTeamInternet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1920
      TabIndex        =   6
      Top             =   8880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picTeam 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   7140
      ScaleHeight     =   1440
      ScaleWidth      =   1875
      TabIndex        =   3
      ToolTipText     =   "Dbl-click to Re-Dock"
      Top             =   2100
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgEmailHelp 
      Height          =   435
      Left            =   6120
      MousePointer    =   99  'Custom
      Picture         =   "frmEmailTeam.frx":0C3C
      Stretch         =   -1  'True
      ToolTipText     =   "View Email-Notification QuickHelp"
      Top             =   240
      Width           =   435
   End
   Begin VB.Menu mnuRemove 
      Caption         =   "mnuRemove"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveRecip 
         Caption         =   "Remove Recipient"
      End
   End
   Begin VB.Menu mnuClient 
      Caption         =   "mnuClient"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveClient 
         Caption         =   "Remove Client Approver"
      End
   End
End
Attribute VB_Name = "frmEmailTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lMeWidth As Long
Dim bRemoving As Boolean, bResetting As Boolean, bGPJEmp As Boolean, bRepOpen As Boolean, bTeam As Boolean
Dim TPos As Integer, iTeamType As Integer, iDrag As Integer, iCurrMenu As Integer, iButton As Integer

Dim tBCC As String, tFBCN As String, tCPRJ As String
Dim tSHCD As Long

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

Private Sub cboCPRJ_Change()
    If cboCPRJ.Text <> "" Then
        tCPRJ = Left(cboCPRJ.List(cboCPRJ.ListIndex), InStr(1, cboCPRJ.List(cboCPRJ.ListIndex), "-") - 2)
        fraTeam.Enabled = True
    Else
        fraTeam.Enabled = False
    End If
End Sub

Private Sub cboCPRJ_Click()
    If cboCPRJ.Text <> "" Then
        tCPRJ = Left(cboCPRJ.List(cboCPRJ.ListIndex), InStr(1, cboCPRJ.List(cboCPRJ.ListIndex), "-") - 2)
        Debug.Print tCPRJ
        fraTeam.Enabled = True
    Else
        fraTeam.Enabled = False
    End If
End Sub

Private Sub cboCUNO_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If cboCUNO.Text <> "" Then
        tBCC = cboCUNO.ItemData(cboCUNO.ListIndex)
        tFBCN = cboCUNO.Text
        Call CheckForClients(CLng(tBCC))
        If optTeamType(0).Value = True Then
            fraTeam.Enabled = True
            CheckForTeam
        Else
            fraTeam.Enabled = False
            '///// GET SHCDS OR CPRJS \\\\\
            Select Case iTeamType
                Case 1
                    cboSHCD.Clear
                    strSelect = "SELECT CS.CSY56SHCD, SM.SHY56NAMA " & _
                                "FROM " & F5611 & " CS, " & F5601 & " SM " & _
                                "WHERE CS.CSY56CUNO = " & CLng(tBCC) & " " & _
                                "AND CS.CSY56SHYR = " & iCurrSHYR & " " & _
                                "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                                "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                                "ORDER BY UPPER(SM.SHY56NAMA)"
                    Set rst = Conn.Execute(strSelect)
                    Do While Not rst.EOF
                        cboSHCD.AddItem UCase(Trim(rst.Fields("SHY56NAMA")))
                        cboSHCD.ItemData(cboSHCD.NewIndex) = rst.Fields("CSY56SHCD")
                        rst.MoveNext
                    Loop
                    rst.Close
                    Set rst = Nothing
                Case 2
                    cboCPRJ.Clear
                    strSelect = "SELECT MCMCU, MCDL01 " & _
                                "FROM " & F0006 & " " & _
                                "WHERE MCAN8O = " & CLng(tBCC) & " " & _
                                "AND MCSTYL IN ('JS', 'JC', 'JR') " & _
                                "AND MCRP21 = " & iCurrSHYR & " " & _
                                "AND TRIM(MCRP14) = '002' " & _
                                "ORDER BY MCSTYL, UPPER(MCDL01)"
                    Set rst = Conn.Execute(strSelect)
                    Do While Not rst.EOF
                        cboCPRJ.AddItem UCase(Trim(rst.Fields("MCMCU"))) & " - " & _
                                    UCase(Trim(rst.Fields("MCDL01")))
                        rst.MoveNext
                    Loop
                    rst.Close
                    Set rst = Nothing
            End Select
        End If
    End If
End Sub


Private Sub cboCUNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboCUNO.Text <> "" Then
            tBCC = Right("00000000" & CStr(cboCUNO.ItemData(cboCUNO.ListIndex)), 8)
            CheckForTeam
        End If
    End If
End Sub

Private Sub cboGroups_Click()
    Dim lGroupID As Long
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    
    lstClients.Visible = False
    
    '///// FIRST, CLEAR EXISTING CHECKS \\\\\
    For i = 0 To lstClients.ListCount - 1
        lstClients.Selected(i) = False
    Next i
        
    If Trim(cboGroups.Text) <> "" Then
        '///// NOW, CHECK 'EM OFF \\\\\
        lGroupID = cboGroups.ItemData(cboGroups.ListIndex)
        strSelect = "SELECT G.AN8_CUNO, C.ABALPH " & _
                    "FROM " & IGLCGR & " G, " & F0101 & " C " & _
                    "WHERE G.CUNO_GROUP_ID = " & lGroupID & " " & _
                    "AND G.AN8_CUNO = C.ABAN8 " & _
                    "ORDER BY UPPER(C.ABALPH)" 'CHANGE - CHECK TABLE ACCESS
        Set rst = Conn.Execute(strSelect)
        Do While Not rst.EOF
            For i = 0 To lstClients.ListCount - 1
                If lstClients.ItemData(i) = rst.Fields("AN8_CUNO") Then
                    lstClients.Selected(i) = True
                    GoTo FoundIt
                End If
            Next i
FoundIt:
            rst.MoveNext
        Loop
        rst.Close: Set rst = Nothing
    
    End If
    lstClients.Visible = True
End Sub

Private Sub cboSHCD_Click()
    If cboSHCD.Text <> "" Then
        Screen.MousePointer = 11
        tSHCD = CLng(cboSHCD.ItemData(cboSHCD.ListIndex))
        Debug.Print tSHCD
        CheckForTeam
        fraTeam.Enabled = True
        Screen.MousePointer = 0
    End If
End Sub


Private Sub cmdAdd_Click()
    Dim bChanged As Boolean
    Dim i As Integer, iC As Integer
    Dim lst1 As ListBox
    
    bChanged = False
    bResetting = True
    If bGPJEmp Then bChanged = AddToList(lstGPJEmps) Else bChanged = AddToList(lstNonGPJ)
    
'''    If bChanged = True Then
'''        cmdSave.Enabled = False
'''        For iC = 0 To lstTeam.ListCount - 1
'''            If lstTeam.Selected(iC) = True Then
'''                cmdSave.Enabled = True
'''                GoTo GetOut
'''            End If
'''        Next iC
'''    End If
GetOut:
    txtGetName.Text = ""
    Text1.Text = ""
    bResetting = False
    SeeIfSaveable
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCloseClients_Click()
    cboGroups.Text = " "
    Frame2.Visible = False
    fraTeam.Width = 10995
'''    Me.Move Me.Left + ((Me.Width - lMeWidth) / 2)
'''    Me.Width = 6975 '''6720
    bRepOpen = False
End Sub

Private Sub cmdRemove_Click()
    If cmdRemove.Caption = "Remove" Then
        bRemoving = True
        lblInstruct2.Caption = "Select Team member to remove."
        cmdRemove.Caption = "Cancel"
    Else
        bRemoving = False
        lblInstruct2.Caption = "At minimum, one Team Member must be assigned to each of the " & _
                    "catagories on the right.  To assign them, left-mouse-drag them from the Team List above."
        cmdRemove.Caption = "Remove"
    End If
End Sub

Private Sub cmdReplicateTeam_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    Select Case iTeamType
        Case 0
            '///// POP CLIENT LIST \\\\\
            lstClients.Clear
            strSelect = "SELECT ABAN8, ABALPH " & _
                        "FROM " & F0101 & " " & _
                        "WHERE ABAT1 = 'C' " & _
                        "ORDER BY UPPER(LTRIM(ABALPH))" 'TEST
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If Left(rst.Fields("ABALPH"), 1) <> "*" Then
                    lstClients.AddItem UCase(Trim(rst.Fields("ABALPH")))
                    lstClients.ItemData(lstClients.NewIndex) = rst.Fields("ABAN8")
                End If
                rst.MoveNext
            Loop
            rst.Close
            Set rst = Nothing
        Case 1
            '///// POP CURRENT CLIENT SHOW LIST \\\\\
            lstClients.Clear
            For i = 0 To cboSHCD.ListCount - 1
                lstClients.AddItem cboSHCD.List(i)
                lstClients.ItemData(i) = cboSHCD.ItemData(i)
            Next i
        Case 2
            '///// POP CURRENT PROJECT LIST \\\\\
            lstClients.Clear
            For i = 0 To cboCPRJ.ListCount - 1
                lstClients.AddItem cboCPRJ.List(i)
            Next i
    End Select
    
    If bRepOpen = False Then
        bRepOpen = True
'''        Me.Width = 11445 ''' 11115
        fraTeam.Width = 6555
        Frame2.Visible = True
'''        Me.Move Me.Left - ((Me.Width - lMeWidth) / 2)
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strSelect As String, strDelete As String, strInsert As String, strUpdate As String
    Dim rst As ADODB.Recordset, rstL As ADODB.Recordset
    Dim lTEAM_ID As Long
    Dim i As Integer, iCheck As Integer, iR As Integer
    
    Screen.MousePointer = 11
    Conn.BeginTrans
    '///// FIRST DELETE EXISTING TEAM \\\\\
    Select Case iTeamType
        Case 0
            strSelect = "SELECT TEAM_ID FROM " & ANOETeam & " " & _
                        "WHERE AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND AN8_SHCD IS NULL AND MCU IS NULL"
        Case 1
            strSelect = "SELECT TEAM_ID FROM " & ANOETeam & " " & _
                        "WHERE AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND AN8_SHCD = " & tSHCD & " " & _
                        "AND MCU IS NULL"
        Case 2
            strSelect = "SELECT TEAM_ID FROM " & ANOETeam & " " & _
                        "WHERE AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND MCU = '" & tCPRJ & "' " & _
                        "AND AN8_SHCD IS NULL"
    End Select
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        '///// A TEAM ALREADY EXISTS \\\\\
        lTEAM_ID = rst.Fields("TEAM_ID")
        rst.Close
        Set rst = Nothing
        strDelete = "DELETE FROM " & ANOETeamUR & " " & _
                    "WHERE TEAM_ID IN (" & strSelect & ")"
        Conn.Execute (strDelete)
    Else
        rst.Close
        Set rst = Nothing
        If lstTeam.ListCount > 0 Then
            Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
            lTEAM_ID = rstL.Fields("nextval")
            rstL.Close: Set rstL = Nothing
            Select Case iTeamType
                Case 0
                    strInsert = "INSERT INTO " & ANOETeam & " " & _
                                "(TEAM_ID, AN8_CUNO, " & _
                                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                                "VALUES " & _
                                "(" & lTEAM_ID & ", " & CLng(tBCC) & ", " & _
                                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
                Case 1
                    strInsert = "INSERT INTO " & ANOETeam & " " & _
                                "(TEAM_ID, AN8_CUNO, AN8_SHCD, " & _
                                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                                "VALUES " & _
                                "(" & lTEAM_ID & ", " & CLng(tBCC) & ", " & tSHCD & ", " & _
                                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
                Case 2
                    strInsert = "INSERT INTO " & ANOETeam & " " & _
                                "(TEAM_ID, AN8_CUNO, MCU, " & _
                                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                                "VALUES " & _
                                "(" & lTEAM_ID & ", " & CLng(tBCC) & ", '" & tCPRJ & "', " & _
                                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            End Select
            Conn.Execute (strInsert)
        End If
    End If
    
    '///// READY TO ADD TEAM MEMBERS \\\\\
    bTeam = False
    If lstTeam.ListCount > 0 Then
        For i = 0 To lstTeam.ListCount - 1
'''            If lstTeam.Selected(i) = True Then iCheck = 1 Else iCheck = 0
            strInsert = "INSERT INTO " & ANOETeamUR & " " & _
                        "(TEAM_ID, USER_SEQ_ID, RECIPIENT_FLAG0, RECIPIENT_FLAG1, RECIPIENT_FLAG2, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & lTEAM_ID & ", " & lstTeam.ItemData(i) & ", 0, 0, 0, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            bTeam = True
        Next i
    End If
    '///// NOW EDIT MANDATORY RECIPIENTS \\\\\
    For i = 0 To lstMandRips(0).ListCount - 1
        strUpdate = "UPDATE " & ANOETeamUR & " " & _
                    "SET RECIPIENT_FLAG0 = 1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                    "AND USER_SEQ_ID = " & lstMandRips(0).ItemData(i)
        Conn.Execute (strUpdate)
    Next i
    For i = 0 To lstMandRips(1).ListCount - 1
        strUpdate = "UPDATE " & ANOETeamUR & " " & _
                    "SET RECIPIENT_FLAG1 = 1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                    "AND USER_SEQ_ID = " & lstMandRips(1).ItemData(i)
        Conn.Execute (strUpdate)
    Next i
    For i = 0 To lstMandRips(2).ListCount - 1
        strUpdate = "UPDATE " & ANOETeamUR & " " & _
                    "SET RECIPIENT_FLAG2 = 1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                    "AND USER_SEQ_ID = " & lstMandRips(2).ItemData(i)
        Conn.Execute (strUpdate)
    Next i
    '///// NOW EDIT CLIENT APPROVERS \\\\\'
    For i = 0 To lstClientApprovers.ListCount - 1
        strUpdate = "UPDATE " & ANOETeamUR & " " & _
                    "SET EXTCLIENTAPPROVER_FLAG = 1, " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                    "AND USER_SEQ_ID = " & lstClientApprovers.ItemData(i)
        Conn.Execute (strUpdate)
    Next i
    
    Conn.CommitTrans
    
    '///// NOW CLEAN UP \\\\\
    cmdSave.Enabled = False
    If bTeam Then cmdReplicateTeam.Enabled = True Else cmdReplicateTeam.Enabled = False
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdSaveClients_Click()
    Dim strInsert As String, strUpdate As String
    Dim rstL As ADODB.Recordset
    Dim sSourceFile As String, sDestinationFile As String, sChk As String, sMess As String
    Dim i As Integer, iC As Integer, iCheck As Integer, iM As Integer
    Dim lTEAM_ID As Long
    Dim bContinue As Boolean
    
    Select Case iTeamType
        Case 0
            If CheckExistingTeams Then
                Screen.MousePointer = 11
                
                sMess = "Current Email Notification Team replicated to the following:" & _
                            vbNewLine & vbNewLine
                
                Conn.BeginTrans
                On Error GoTo ErrorTrap
                For iC = 0 To lstClients.ListCount - 1
                    If lstClients.Selected(iC) = True And CLng(tBCC) <> lstClients.ItemData(iC) Then
                        sMess = sMess & vbTab & lstClients.List(iC) & vbNewLine
                        Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
                        lTEAM_ID = rstL.Fields("nextval")
                        rstL.Close: Set rstL = Nothing
                        strInsert = "INSERT INTO " & ANOETeam & " " & _
                                    "(TEAM_ID, AN8_CUNO, " & _
                                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                                    "VALUES " & _
                                    "(" & lTEAM_ID & ", " & lstClients.ItemData(iC) & ", " & _
                                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
                        Conn.Execute (strInsert)
                        For i = 0 To lstTeam.ListCount - 1
'''                            If lstTeam.Selected(i) = True Then iCheck = 1 Else iCheck = 0
                            strInsert = "INSERT INTO " & ANOETeamUR & " " & _
                                        "(TEAM_ID, USER_SEQ_ID, RECIPIENT_FLAG0, RECIPIENT_FLAG1, RECIPIENT_FLAG2, " & _
                                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                                        "VALUES " & _
                                        "(" & lTEAM_ID & ", " & lstTeam.ItemData(i) & ", 0, 0, 0, " & _
                                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
                            Conn.Execute (strInsert)
                        Next i
                        '///// NOW ADD MANDATORY RECIPIENTS \\\\\
                        For i = 0 To lstMandRips(0).ListCount - 1
                            strUpdate = "UPDATE " & ANOETeamUR & " " & _
                                        "SET RECIPIENT_FLAG0 = 1, " & _
                                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                                        "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                                        "AND USER_SEQ_ID = " & lstMandRips(0).ItemData(i)
                            Conn.Execute (strUpdate)
                        Next i
                        For i = 0 To lstMandRips(1).ListCount - 1
                            strUpdate = "UPDATE " & ANOETeamUR & " " & _
                                        "SET RECIPIENT_FLAG1 = 1, " & _
                                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                                        "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                                        "AND USER_SEQ_ID = " & lstMandRips(1).ItemData(i)
                            Conn.Execute (strUpdate)
                        Next i
                        For i = 0 To lstMandRips(2).ListCount - 1
                            strUpdate = "UPDATE " & ANOETeamUR & " " & _
                                        "SET RECIPIENT_FLAG2 = 1, " & _
                                        "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                                        "WHERE TEAM_ID = " & lTEAM_ID & " " & _
                                        "AND USER_SEQ_ID = " & lstMandRips(2).ItemData(i)
                            Conn.Execute (strUpdate)
                        Next i
                    End If
                    lstClients.Selected(iC) = False
                Next iC
                Conn.CommitTrans
                Screen.MousePointer = 0
                MsgBox sMess, vbInformation, "Team Setup Complete..."
            End If
        Case Else
            MsgBox "Still in development.", vbExclamation, "Sorry..."
    End Select
Exit Sub
ErrorTrap:
    Conn.RollbackTrans
    Screen.MousePointer = 0
    MsgBox "Error:  " & Err.Description, vbCritical, "Error Encountered..."
End Sub


Private Sub Form_Load()
        
    Screen.MousePointer = 11
    
    bGPJEmp = True
    bRepOpen = False
    
    '///// POP cboCUNO LIST (ALL CLIENTS) \\\\\
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    strSelect = "SELECT DISTINCT ABAN8, ABALPH " & _
                "FROM " & F0101 & " " & _
                "WHERE ABAT1 = 'C' " & _
                "ORDER BY UPPER(ABALPH)"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If Left(rst.Fields("ABALPH"), 1) <> "*" Then
            cboCUNO.AddItem UCase(Trim(rst.Fields("ABALPH")))
            cboCUNO.ItemData(cboCUNO.NewIndex) = rst.Fields("ABAN8")
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    If tBCC <> "" Then
        cboCUNO.Text = UCase(Trim(tFBCN))
        CheckForTeam
    End If
    
    '///// POP GPJ EMPLOYEE LIST \\\\\
    strSelect = "SELECT USER_STATUS,USER_SEQ_ID, NAME_LOGON, " & _
                "NAME_LAST, NAME_FIRST, EMAIL_ADDRESS " & _
                "FROM " & IGLUser & " " & _
                "WHERE USER_STATUS = 1 " & _
                "AND SUBSTR(EMPLOYER, 1, 3) = 'GPJ' " & _
                "AND EMAIL_ADDRESS IS NOT NULL " & _
                "ORDER BY NAME_LAST, NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstGPJEmps.AddItem Trim(rst.Fields("NAME_FIRST")) & " " & _
                    Trim(rst.Fields("NAME_LAST"))
        lstGPJEmps.ItemData(lstGPJEmps.NewIndex) = rst.Fields("USER_SEQ_ID")
        lstGPJEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
'''    '///// NOW, POP NON-GPJ LIST \\\\\
'''    strSelect = "SELECT USER_STATUS, USER_SEQ_ID, NAME_LOGON, " & _
'''                "NAME_LAST, NAME_FIRST, EMPLOYER, EMAIL_ADDRESS " & _
'''                "FROM " & IGLUser & " " & _
'''                "WHERE USER_STATUS = 1 " & _
'''                "AND SUBSTR(EMPLOYER, 1, 3) <> 'GPJ' " & _
'''                "AND EMAIL_ADDRESS IS NOT NULL " & _
'''                "ORDER BY NAME_LAST, NAME_FIRST"
'''    Set rst = Conn.Execute(strSelect)
'''    Do While Not rst.EOF
'''        lstNonGPJ.AddItem Trim(rst.Fields("NAME_FIRST")) & " " & _
'''                    Trim(rst.Fields("NAME_LAST")) & "  (" & Trim(rst.Fields("EMPLOYER")) & ")"
'''        lstNonGPJ.ItemData(lstNonGPJ.NewIndex) = rst.Fields("USER_SEQ_ID")
'''        lstNonGPJEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
'''        rst.MoveNext
'''    Loop
'''    rst.Close
'''    Set rst = Nothing
            
    '///// GET GROUPS \\\\\
    cboGroups.AddItem " "
    strSelect = "SELECT CUNO_GROUP_ID, CUNO_GROUP_DESC " & _
                "FROM " & IGLCGMas & " " & _
                "ORDER BY UPPER(CUNO_GROUP_DESC)"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboGroups.AddItem Trim(rst.Fields("CUNO_GROUP_DESC"))
        cboGroups.ItemData(cboGroups.NewIndex) = rst.Fields("CUNO_GROUP_ID")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
'''    Me.Width = 6975 '''6720
    lMeWidth = Me.Width
    Me.Height = 8790 '''8580
    Screen.MousePointer = 0

End Sub

Private Sub Form_Paint()
    Debug.Print "Form RePaint"
End Sub

Private Sub imgEmailHelp_Click()
    Dim EmailHelp As String
    EmailHelp = "GPJ Annotator Email-Notification Teams are ClientCode-based by default.  " & _
                "Teams can be ClientShow- or Project-based by selecting one of those options." & _
                vbNewLine & vbNewLine & "To add an Employee to a Team, " & _
                "either type their name in the noted field, and press <ENTER> " & _
                "on your keyboard, OR... From the lists on the left, " & _
                "select all members to be added to the " & _
                cboCUNO.Text & " Team.  Then, click 'ADD TO TEAM' " & _
                "to include the highlighted person(s) on the Team list." & _
                vbNewLine & vbNewLine & "To remove a Team member, " & _
                "select 'REMOVE', then click on the member on the Team list to remove."
    MsgBox EmailHelp, vbQuestion, "Email-Notification Team Setup Help"
End Sub

Private Sub imgRepTeamHelp_Click()
    Dim RepHelp As String
    RepHelp = "The 'Replicate Team' control is only available when the current Client " & _
            "has an associated Email Notification Team.  If one has not been set up, " & _
            "the control will be disabled." & vbCr & vbCr & _
            "To Replicate a Team, click on the 'Replicate Team' button.  " & _
            "This interface will expand to show a list of all Clients with Annotator Floor Plans.  " & _
            "Check all Clients you wish to copy the current Team to, and select 'SAVE'."
    MsgBox RepHelp, vbQuestion, "Team Replication Help"
End Sub

Private Sub lstClientApprovers_DragDrop(Source As Control, x As Single, y As Single)
    Dim i As Integer
    If iDrag > -1 Then
        If InStr(1, lstTeam.List(iDrag), " (") = 0 Then
            MsgBox "Only External Clients can be set in this option", vbExclamation, "Sorry..."
            Exit Sub
        Else
            For i = 0 To lstClientApprovers.ListCount - 1
                If lstClientApprovers.List(i) = lstTeam.List(iDrag) Then Exit Sub
            Next i
        End If
        lstClientApprovers.AddItem lstTeam.List(iDrag)
        lstClientApprovers.ItemData(lstClientApprovers.NewIndex) = lstTeam.ItemData(iDrag)
        SeeIfSaveable
    End If
End Sub

Private Sub lstClientApprovers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuClient, , fraTeam.Left + fraClientApprovers.Left + x, _
                    fraTeam.Top + fraClientApprovers.Top + y
    End If
End Sub

Private Sub lstGPJEmps_DblClick()
    cmdAdd_Click
End Sub

Private Sub lstGPJEmps_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bGPJEmp = True
    cmdAdd.Enabled = True
End Sub

Private Sub lstMandRips_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If iDrag > -1 Then
        lstMandRips(Index).AddItem lstTeam.List(iDrag)
        lstMandRips(Index).ItemData(lstMandRips(Index).NewIndex) = lstTeam.ItemData(iDrag)
        SeeIfSaveable
    End If
End Sub

Private Sub lstMandRips_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        iCurrMenu = Index
        PopupMenu mnuRemove
    End If
    SeeIfSaveable
End Sub

Private Sub lstNonGPJ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bGPJEmp = False
    cmdAdd.Enabled = True
End Sub

Private Sub lstTeam_Click()
    Dim iM As Integer, i As Integer
    Dim lUserID As Long
    
    Debug.Print lstTeam.ItemData(lstTeam.ListIndex)
    If bRemoving = True Then
        bRemoving = False
        cmdRemove.Caption = "Remove"
'''        lstTeamEmail.RemoveItem (lstTeam.ListIndex)
        lUserID = lstTeam.ItemData(lstTeam.ListIndex)
        For iM = 0 To 2
            For i = (lstMandRips(iM).ListCount - 1) To 0 Step -1
                If lstMandRips(iM).ItemData(i) = lUserID Then lstMandRips(iM).RemoveItem (i)
            Next i
        Next iM
        lstTeam.RemoveItem (lstTeam.ListIndex)
        
        lblInstruct2.Caption = "At minimum, one Team Member must be assigned to each of the " & _
                    "catagories on the right.  To assign them, left-mouse-drag them from the Team List above."
        If lstTeam.ListCount = 0 And bTeam = True Then
            cmdSave.Enabled = True
        Else
            SeeIfSaveable
        End If
    End If
End Sub

'''Private Sub lstTeam_ItemCheck(Item As Integer)
'''    Dim i As Integer
'''    For i = 0 To lstTeam.ListCount - 1
'''        If lstTeam.Selected(i) = True Then
'''            cmdSave.Enabled = True
'''            GoTo FoundOne
'''        End If
'''    Next
'''    cmdSave.Enabled = False
'''FoundOne:
'''End Sub

Private Sub lstTeam_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    iDrag = lstTeam.ListIndex
    iButton = Button
    Debug.Print "Button = " & Button
    Debug.Print iDrag
    Dim DY   ' Declare variable.
    DY = TextHeight("A")   ' Get height of one line.
    lblDrag.Move lstTeam.Left, lstTeam.Top + y - DY / 2, lstTeam.Width, DY
    lblDrag.Drag   ' Drag label outline.

End Sub

Private Sub mnuRemoveClient_Click()
    Dim i As Integer
    For i = (lstClientApprovers.ListCount - 1) To 0 Step -1
        If lstClientApprovers.Selected(i) = True Then lstClientApprovers.RemoveItem (i)
    Next i
End Sub

Private Sub mnuRemoveRecip_Click()
    Dim i As Integer
    For i = (lstMandRips(iCurrMenu).ListCount - 1) To 0 Step -1
        If lstMandRips(iCurrMenu).Selected(i) = True Then lstMandRips(iCurrMenu).RemoveItem (i)
    Next i
End Sub

Private Sub optTeamType_Click(Index As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    Screen.MousePointer = 11
    iTeamType = Index
    Select Case iTeamType
        Case 0
            fraSHCD.Visible = False
            fraCPRJ.Visible = False
            If cboCUNO.Text <> "" Then CheckForTeam
        Case 1
            cboSHCD.Clear
            lstTeam.Clear
            If tBCC <> "" Then
                '///// POP CLIENT SHOWS \\\\\
                strSelect = "SELECT CS.CSY56SHCD, SM.SHY56NAMA " & _
                            "FROM " & F5611 & " CS, " & F5601 & " SM " & _
                            "WHERE CS.CSY56CUNO = " & CLng(tBCC) & " " & _
                            "AND CS.CSY56SHYR = " & iCurrSHYR & " " & _
                            "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                            "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                            "ORDER BY UPPER(SM.SHY56NAMA)"
                Set rst = Conn.Execute(strSelect)
                Do While Not rst.EOF
                    cboSHCD.AddItem UCase(Trim(rst.Fields("SHY56NAMA")))
                    cboSHCD.ItemData(cboSHCD.NewIndex) = rst.Fields("CSY56SHCD")
                    rst.MoveNext
                Loop
                rst.Close
                Set rst = Nothing
                tSHCD = 0
            End If
            fraSHCD.Visible = True
            fraCPRJ.Visible = False
        Case 2
            cboCPRJ.Clear
            lstTeam.Clear
            If tBCC <> "" Then
                '///// POP ACTIVE PROJECTS \\\\\
                tCPRJ = ""
            End If
            fraSHCD.Visible = False
            fraCPRJ.Visible = True
    End Select
    
    If bRepOpen And lstTeam.ListCount <> 0 Then
        cmdReplicateTeam_Click
    ElseIf bRepOpen Then
        cmdCloseClients_Click
        cmdReplicateTeam.Enabled = False
    ElseIf lstTeam.ListCount = 0 Then
        cmdReplicateTeam.Enabled = False
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub txtGetName_Change()
    Dim i As Integer
    If bResetting = False Then
        For i = 0 To lstGPJEmps.ListCount - 1
            lstGPJEmps.Selected(i) = False
        Next i
        cmdAdd.Enabled = False
        If Len(Text1.Text) > 0 Then
            For i = 0 To lstGPJEmps.ListCount - 1
                If UCase(Left(lstGPJEmps.List(i), Len(txtGetName.Text))) = UCase(txtGetName.Text) Then
                    txtGetName.Text = lstGPJEmps.List(i)
                    lstGPJEmps.Selected(i) = True
                    cmdAdd.Enabled = True
                    cmdAdd.Default = True
                    GoTo FoundOne
                End If
            Next i
FoundOne:
            txtGetName.SelStart = TPos
            txtGetName.SelLength = Len(txtGetName.Text) - TPos
        Else
            txtGetName.Text = ""
        End If
    End If
End Sub

Private Sub txtGetName_KeyDown(KeyCode As Integer, Shift As Integer)
    bGPJEmp = True
End Sub

Private Sub txtGetName_KeyPress(KeyAscii As Integer)
    Debug.Print "Chr(" & KeyAscii & ")"
    If KeyAscii = 8 Then
        If Len(Text1.Text) > 0 Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    Else
        Text1.Text = Text1.Text & Chr(KeyAscii)
    End If
    TPos = Len(Text1.Text)
    If Len(Text1.Text) = 0 Then txtGetName.Text = ""
End Sub

Public Sub CheckForTeam()
    Dim sWho As String, sEMail As String, sEmployer As String
    Dim iVal(0 To 2) As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim lWhoID As Long
    Dim i As Integer
    
    bTeam = False
    lstTeam.Clear
    lstTeamEmail.Clear
    For i = 0 To 2
        lstMandRips(i).Clear
    Next i
    lstClientApprovers.Clear
    If optTeamType(0).Value = True Then iTeamType = 0
    If optTeamType(1).Value = True Then iTeamType = 1
    If optTeamType(2).Value = True Then iTeamType = 2
    
    Select Case iTeamType
        Case 0
            strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, " & _
                        "U.NAME_LOGON, U.EMAIL_ADDRESS, U.EMPLOYER, R.EXTCLIENTAPPROVER_FLAG, " & _
                        "R.RECIPIENT_FLAG0, R.RECIPIENT_FLAG1, R.RECIPIENT_FLAG2 " & _
                        "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                        "WHERE T.AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND T.AN8_SHCD IS NULL " & _
                        "AND T.MCU IS NULL " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Case 1
            strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, " & _
                        "U.NAME_LOGON, U.EMAIL_ADDRESS, U.EMPLOYER, R.EXTCLIENTAPPROVER_FLAG, " & _
                        "R.RECIPIENT_FLAG0, R.RECIPIENT_FLAG1, R.RECIPIENT_FLAG2 " & _
                        "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                        "WHERE T.AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND T.AN8_SHCD = " & tSHCD & " " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Case 2
            strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, " & _
                        "U.NAME_LOGON, U.EMAIL_ADDRESS, U.EMPLOYER, R.EXTCLIENTAPPROVER_FLAG, " & _
                        "R.RECIPIENT_FLAG0, R.RECIPIENT_FLAG1, R.RECIPIENT_FLAG2 " & _
                        "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                        "WHERE T.AN8_CUNO = " & CLng(tBCC) & " " & _
                        "AND T.MCU = '" & tCPRJ & "' " & _
                        "AND T.TEAM_ID = R.TEAM_ID " & _
                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    End Select
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        cmdReplicateTeam.Enabled = True
        bTeam = True
        Do While Not rst.EOF
            If UCase(Left(rst.Fields("EMPLOYER"), 3)) = "GPJ" Then
                sEmployer = ""
            Else
                sEmployer = "  (" & UCase(Trim(rst.Fields("EMPLOYER"))) & ")"
            End If
            sWho = Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")) & sEmployer
            lWhoID = rst.Fields("USER_SEQ_ID")
            If IsNull(rst.Fields("RECIPIENT_FLAG0")) Then iVal(0) = 0 Else iVal(0) = rst.Fields("RECIPIENT_FLAG0")
            If IsNull(rst.Fields("RECIPIENT_FLAG1")) Then iVal(1) = 0 Else iVal(1) = rst.Fields("RECIPIENT_FLAG1")
            If IsNull(rst.Fields("RECIPIENT_FLAG2")) Then iVal(2) = 0 Else iVal(2) = rst.Fields("RECIPIENT_FLAG2")
            If rst.Fields("EXTCLIENTAPPROVER_FLAG") > 0 Then
                lstClientApprovers.AddItem sWho
                lstClientApprovers.ItemData(lstClientApprovers.NewIndex) = lWhoID
            End If
            sEMail = Trim(rst.Fields("EMAIL_ADDRESS"))
            lstTeam.AddItem sWho
'''            lstTeam.Selected(lstTeam.NewIndex) = CBool(iVal(0))
            lstTeam.ItemData(lstTeam.NewIndex) = lWhoID
            lstTeamEmail.AddItem sEMail
            lstTeamEmail.ItemData(lstTeamEmail.NewIndex) = lWhoID
            
            For i = 0 To 2
                If iVal(i) = 1 Then
                    lstMandRips(i).AddItem sWho
                    lstMandRips(i).ItemData(lstMandRips(i).NewIndex) = lWhoID
                End If
            Next i
            rst.MoveNext
            
        Loop
    Else
        cmdReplicateTeam.Enabled = False
        If bRepOpen Then cmdCloseClients_Click
    End If
    rst.Close
    Set rst = Nothing
    
    lblInstruct2.Caption = "At minimum, one Team Member must be assigned to each of the " & _
                    "catagories on the right.  To assign them, left-mouse-drag them from the Team List above."
    
End Sub

Public Function AddToList(lst1 As ListBox) As Boolean
    Dim i As Integer, iC As Integer
    Dim bChanged As Boolean
    
    For i = 0 To lst1.ListCount - 1
        If lst1.Selected(i) = True Then
            For iC = 0 To lstTeam.ListCount - 1
                If lst1.List(i) = lstTeam.List(iC) Then GoTo foundacopy
            Next iC
            lstTeam.AddItem lst1.List(i)
            lstTeam.ItemData(lstTeam.NewIndex) = lst1.ItemData(i)
            bChanged = True
foundacopy:
            lst1.Selected(i) = False
        End If
    Next i
    AddToList = bChanged
End Function

Public Function CheckExistingTeams() As Boolean
    Dim iC As Integer
    Dim sMess As String, strSelect As String, strDelete As String, sTeamID As String
    Dim rst As ADODB.Recordset
    Dim Resp As VbMsgBoxResult
    Dim iTeamID As Long
    
    sMess = "": sTeamID = ""
    For iC = 0 To lstClients.ListCount - 1
        If lstClients.Selected(iC) = True And CLng(tBCC) <> lstClients.ItemData(iC) Then
            '///// CHECK FOR EXISTING TEAM \\\\\
            strSelect = "SELECT AN8_CUNO FROM " & ANOETeam & " " & _
                        "WHERE AN8_CUNO = " & lstClients.ItemData(iC)
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                sMess = "The following Client already has an associated Team:" & _
                            vbNewLine & vbNewLine & vbTab & lstClients.List(iC) & vbNewLine & vbNewLine & _
                            "Select 'YES' to create new Team (overwriting the existing)." & vbNewLine & _
                            "Select 'NO' to retain existing Team."
                Resp = MsgBox(sMess, vbYesNo, "Existing Team Found...")
                If Resp = vbYes Then
                    '///// GET TEAM ID \\\\\
                    rst.Close
                    strSelect = "SELECT TEAM_ID FROM " & ANOETeam & " " & _
                                "WHERE AN8_CUNO = " & lstClients.ItemData(iC)
                    Set rst = Conn.Execute(strSelect)
                    Do While Not rst.EOF
                        If sTeamID = "" Then
                            sTeamID = CStr(rst.Fields("TEAM_ID"))
                        Else
                            sTeamID = sTeamID & ", " & CStr(rst.Fields("TEAM_ID"))
                        End If
                        rst.MoveNext
                    Loop
                    '///// DELETE PREVIOUS TEAM \\\\\
                    strDelete = "DELETE FROM " & ANOETeamUR & " " & _
                                "WHERE TEAM_ID IN (" & sTeamID & ")"
                    Conn.Execute (strDelete)
                    strDelete = "DELETE FROM " & ANOETeam & " " & _
                                "WHERE AN8_CUNO = " & lstClients.ItemData(iC)
                    Conn.Execute (strDelete)
                Else
                    lstClients.Selected(iC) = False
                End If
'''                sMess = sMess & vbTab & lstClients.List(iC) & vbNewLine
            End If
            rst.Close
            Set rst = Nothing
        End If
    Next iC
    
    CheckExistingTeams = True
    
'''    If sMess <> "" Then
'''        sMess = "The following Clients already have associated Teams:" & _
'''                    vbNewLine & vbNewLine & sMess & vbNewLine & _
'''                    "Select 'YES' to create new Teams (overwriting the Teams listed above)." & vbNewLine & _
'''                    "Select 'NO' to cancel."
'''        Resp = MsgBox(sMess, vbYesNo, "Existing Teams Found...")
'''        If Resp = vbYes Then
'''            CheckExistingTeams = True
'''            '///// DELETE PREVIOUS TEAM \\\\\
'''
'''        Else
'''            CheckExistingTeams = False
'''        End If
'''    Else
'''        CheckExistingTeams = True
'''    End If
End Function

Public Sub GetSHCDTeam()
    
End Sub

Public Sub SeeIfSaveable()
    If lstTeam.ListCount > 0 And lstMandRips(0).ListCount > 0 And _
                lstMandRips(1).ListCount > 0 And lstMandRips(2).ListCount > 0 Then _
                cmdSave.Enabled = True Else cmdSave.Enabled = False
End Sub

Public Sub CheckForClients(tBCC As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    lstNonGPJ.Clear: lstNonGPJEmail.Clear
'''    strSelect = "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, " & _
'''                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME, " & _
'''                "US.EMPLOYER " & _
'''                "FROM IGL_USER_CUNO_R CR, IGL_CUNO_GROUP_R GR, " & _
'''                "IGL_USER US, IGL_USER_APP_R AR " & _
'''                "Where GR.AN8_CUNO = " & tBCC & " " & _
'''                "AND GR.CUNO_GROUP_ID = CR.CUNO_GROUP_ID " & _
'''                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
'''                "AND US.USER_STATUS > 0 " & _
'''                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
'''                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
'''                "AND AR.APP_ID = 1002 " & _
'''                "AND AR.PERMISSION_STATUS > 0 " & _
'''                "ORDER BY US.NAME_LAST"
                
    strSelect = "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, US.NAME_LAST, " & _
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
                "SELECT CR.USER_SEQ_ID, US.EMAIL_ADDRESS, US.NAME_LAST, " & _
                "TRIM(US.NAME_FIRST)||' '||TRIM(US.NAME_LAST) AS FULLNAME, US.EMPLOYER " & _
                "FROM IGLPROD.IGL_USER_CUNO_R CR, IGLPROD.IGL_USER US, IGLPROD.IGL_USER_APP_R AR " & _
                "Where CR.AN8_CUNO = " & tBCC & " " & _
                "AND CR.CUNO_GROUP_ID =0 " & _
                "AND CR.USER_SEQ_ID = US.USER_SEQ_ID " & _
                "AND US.USER_STATUS > 0 " & _
                "AND UPPER(US.EMPLOYER) NOT LIKE 'GPJ%' " & _
                "AND CR.USER_SEQ_ID = AR.USER_SEQ_ID " & _
                "AND AR.APP_ID = 1002 " & _
                "AND AR.PERMISSION_STATUS > 0 " & _
                "ORDER BY NAME_LAST"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        lstNonGPJ.AddItem UCase(Trim(rst.Fields("FULLNAME"))) & " (" & _
                    UCase(Trim(rst.Fields("EMPLOYER"))) & ")"
        lstNonGPJ.ItemData(lstNonGPJ.NewIndex) = rst.Fields("USER_SEQ_ID")
        lstNonGPJEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS"))
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub
