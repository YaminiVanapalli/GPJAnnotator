VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GPJ Annotator User Maintenance"
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
   Icon            =   "frmSecurityTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   180
      Visible         =   0   'False
      Width           =   1590
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4035
      Left            =   7620
      TabIndex        =   48
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      ExtentX         =   6800
      ExtentY         =   7117
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmSecurityTest.frx":030A
      Left            =   1560
      List            =   "frmSecurityTest.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   5085
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdTree 
      Caption         =   " Expand Tree"
      Height          =   555
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame fraUserMaint 
      Caption         =   "User Maintenance"
      Enabled         =   0   'False
      Height          =   4095
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   5475
      Begin VB.TextBox txtUserTypeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1875
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1500
         Width           =   5115
      End
      Begin VB.CommandButton cmdNotify 
         Caption         =   "Beta Notify..."
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   45
         Top             =   3405
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdNotify 
         Caption         =   "GPJ Notify..."
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   46
         Top             =   3720
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3180
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   3180
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   240
         Width           =   2115
      End
      Begin VB.CommandButton cmdResetPerm 
         Height          =   435
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3480
         Width           =   3855
      End
      Begin VB.ComboBox cboUserType 
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1140
         Width           =   5115
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         Height          =   195
         Left            =   1770
         TabIndex        =   13
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Password (minimum of 8 characters):"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   2940
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Base User Type:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   1665
      End
   End
   Begin MSComctlLib.TreeView tvw1 
      Height          =   8055
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   14208
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
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
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   3135
   End
   Begin VB.ComboBox cboPerm 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "View User Report..."
      Height          =   375
      Left            =   3660
      TabIndex        =   47
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdViewUsage 
      Caption         =   "View Usage                                          "
      Height          =   375
      Left            =   360
      TabIndex        =   52
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame fraPermMaint 
      Caption         =   "Permission Maintenance"
      Height          =   3015
      Left            =   180
      TabIndex        =   14
      Top             =   5400
      Width           =   5475
      Begin TabDlg.SSTab sst1 
         Height          =   2595
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   4577
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   882
         TabCaption(0)   =   "Add New Permission"
         TabPicture(0)   =   "frmSecurityTest.frx":030E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraAddNewPerm"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Edit Permission"
         TabPicture(1)   =   "frmSecurityTest.frx":032A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Add New User Type"
         TabPicture(2)   =   "frmSecurityTest.frx":0346
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label7"
         Tab(2).Control(1)=   "Label8"
         Tab(2).Control(2)=   "txtUserType"
         Tab(2).Control(3)=   "txtDescUserType(0)"
         Tab(2).Control(4)=   "cmdSaveUserType"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Edit User Type"
         TabPicture(3)   =   "frmSecurityTest.frx":0362
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label9"
         Tab(3).Control(1)=   "Label10"
         Tab(3).Control(2)=   "lblUID"
         Tab(3).Control(3)=   "cmdUpdateUserType"
         Tab(3).Control(4)=   "txtDescUserType(1)"
         Tab(3).Control(5)=   "cboEditUserType"
         Tab(3).ControlCount=   6
         Begin VB.ComboBox cboEditUserType 
            Height          =   315
            Left            =   -74820
            TabIndex        =   42
            Top             =   900
            Width           =   4875
         End
         Begin VB.TextBox txtDescUserType 
            Height          =   915
            Index           =   1
            Left            =   -74820
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   1500
            Width           =   3675
         End
         Begin VB.CommandButton cmdUpdateUserType 
            Caption         =   "  Update   User Type w/Current Permission"
            Height          =   1035
            Left            =   -71040
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1380
            Width           =   1095
         End
         Begin VB.CommandButton cmdSaveUserType 
            Caption         =   "Save New User Type w/Current Permission"
            Height          =   1035
            Left            =   -71040
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1380
            Width           =   1095
         End
         Begin VB.TextBox txtDescUserType 
            Height          =   915
            Index           =   0
            Left            =   -74820
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1500
            Width           =   3675
         End
         Begin VB.TextBox txtUserType 
            Height          =   285
            Left            =   -74820
            MaxLength       =   50
            TabIndex        =   33
            Top             =   900
            Width           =   4875
         End
         Begin VB.Frame Frame1 
            Caption         =   "Permission Details"
            Height          =   1815
            Left            =   -74880
            TabIndex        =   26
            Top             =   660
            Width           =   4995
            Begin VB.TextBox txtPermDesc 
               Height          =   285
               Left            =   960
               MaxLength       =   100
               TabIndex        =   28
               Top             =   720
               Width           =   3855
            End
            Begin VB.CommandButton cmdSavePermDesc 
               Caption         =   " Save New Permission Description"
               Height          =   555
               Left            =   2820
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   1140
               Width           =   2055
            End
            Begin VB.Label lblCurrNode 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   300
               Width           =   45
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Permission Description:"
               Height          =   390
               Left            =   60
               TabIndex        =   29
               Top             =   660
               Width           =   915
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraAddNewPerm 
            Caption         =   "Permission Details"
            Height          =   1815
            Left            =   120
            TabIndex        =   17
            Top             =   660
            Width           =   4995
            Begin VB.CommandButton cmdSave 
               Caption         =   " Save New Permission"
               Enabled         =   0   'False
               Height          =   555
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1140
               Width           =   1395
            End
            Begin VB.CommandButton cmdReset 
               Caption         =   " Cancel && Reset Permission Tree"
               Enabled         =   0   'False
               Height          =   555
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1140
               Width           =   1635
            End
            Begin VB.CommandButton cmdAddToTree 
               Caption         =   " Add Permission to Tree"
               Enabled         =   0   'False
               Height          =   555
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1140
               Width           =   1395
            End
            Begin VB.TextBox txtParent 
               Height          =   285
               Left            =   2640
               TabIndex        =   21
               Text            =   "<Select Parent Node Above>"
               Top             =   720
               Visible         =   0   'False
               Width           =   2235
            End
            Begin VB.OptionButton optParent 
               Caption         =   "Child Node"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   20
               Top             =   720
               Width           =   1155
            End
            Begin VB.OptionButton optParent 
               Caption         =   "Root Node"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   720
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.TextBox txtPerm 
               Height          =   285
               Left            =   1020
               MaxLength       =   100
               TabIndex        =   18
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Description:"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.Label lblUID 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   -70020
            TabIndex        =   43
            Top             =   660
            Width           =   45
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select User Type to Edit:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   41
            Top             =   660
            Width           =   1785
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description of User Type:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   40
            Top             =   1260
            Width           =   1830
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description of New User Type:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   35
            Top             =   1260
            Width           =   2190
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New User Type:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   34
            Top             =   660
            Width           =   1155
         End
      End
   End
   Begin VB.Label lblPrint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSecurityTest.frx":037E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   180
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   9990
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   12000
      TabIndex        =   44
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   195
      Left            =   3420
      TabIndex        =   32
      Top             =   180
      Width           =   840
   End
   Begin VB.Label lblShort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3420
      TabIndex        =   31
      ToolTipText     =   "Click to View User Log & Client Access Rights"
      Top             =   420
      Width           =   1380
   End
   Begin VB.Label lblPermValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   11685
      TabIndex        =   15
      Top             =   150
      Width           =   45
   End
   Begin VB.Label lblPermSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5880
      TabIndex        =   9
      Top             =   150
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   870
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuViewUser 
         Caption         =   "By UserName"
      End
      Begin VB.Menu mnuViewUserType 
         Caption         =   "By UserType"
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewClose 
         Caption         =   "Close Report"
      End
      Begin VB.Menu mnuDash02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTotal As String
Dim Calcing As Boolean, GettingParent As Boolean, bShowNodes As Boolean, bResetting As Boolean, _
            bSaveable As Boolean, bGPJUser As Boolean
Dim strNewNodeLevel As String, sPass As String
Dim CurrNode As String
Dim lUserID As Long




Private Sub cboEditUserType_Click()
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    
    If Trim(cboEditUserType.Text) <> "" Then
        lblUID = cboEditUserType.ItemData(cboEditUserType.ListIndex)
        strSelect = "SELECT USERTYPEDESC, " & _
                    "TO_CHAR(USERTYPEVALUE, '" & String(38, "9") & "') AS PERM " & _
                    "FROM " & ANOUserType & " " & _
                    "WHERE USERTYPEID = " & lblUID.Caption
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            txtDescUserType(1).Text = rst.Fields("USERTYPEDESC")
            sTotal = Trim(rst.Fields("PERM"))
            lblTotal = sTotal
            Call GenPerms(sTotal)
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub cboUser_Click()
    Dim strSelect As String ''''', sTotal As String
    Dim rst As ADODB.Recordset
    
    If Trim(cboUser.Text) <> "" Then
'''        tvw1.Enabled = False
        Debug.Print cboPerm.List(cboUser.ListIndex)
        lblPermSet = "CURRENT Permission Set for " & cboUser.Text
        lUserID = cboUser.ItemData(cboUser.ListIndex)
        strSelect = "SELECT NAME_LOGON FROM " & IGLUser & " " & _
                    "WHERE USER_SEQ_ID = " & lUserID
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            lblShort = UCase(Trim(rst.Fields("NAME_LOGON")))
        Else
            lblShort = ""
        End If
        rst.Close
        Set rst = Nothing
        
        '///// CHECK FOR A PASSWORD \\\\\
        strSelect = "SELECT PCODE FROM " & IGLUserAR & " " & _
                    "WHERE USER_SEQ_ID = " & lUserID & " " & _
                    "AND APP_ID = 1002"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            If Not IsNull(rst.Fields("PCODE")) Then
                txtPassword(0).Text = Trim(rst.Fields("PCODE"))
            Else
                txtPassword(0).Text = ""
            End If
        End If
        rst.Close
        Set rst = Nothing
        
'''''        sTotal = cboPerm.List(cboUser.ListIndex)
        If cboPerm.ItemData(cboUser.ListIndex) > 0 Then
            cboUserType.Text = cboPerm.List(cboUser.ListIndex)
        Else
            cboUserType.Text = " "
            sTotal = "0"
        End If
'''''        lblTotal = sTotal
'''        lblPermValue = Format(Total, "0,###")
'''        Call GenPerms(sTotal)
'''        Total = CDbl(lblTotal)
        bResetting = True
'''        cboUserType.Text = " "
        bResetting = False
'''        txtUserTypeDesc.text = ""
        cmdResetPerm.Enabled = True
        cmdResetPerm.Caption = "Reset Permission for " & cboUser.Text
        fraUserMaint.Enabled = True
        bSaveable = True
        cmdNotify(0).Visible = False: cmdNotify(1).Visible = False
    Else
        lUserID = 0
        ClearNodes
        fraUserMaint.Enabled = False
        bSaveable = False
        cmdNotify(0).Visible = False: cmdNotify(1).Visible = False
    End If
End Sub

Private Sub cboUserType_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If Not bResetting Then
        If Trim(cboUserType.Text) <> "" Then
            
            strSelect = "SELECT USERTYPEDESC, " & _
                        "TO_CHAR(USERTYPEVALUE, '" & String(38, "9") & "') AS PERM " & _
                        "FROM " & ANOUserType & " " & _
                        "WHERE USERTYPEID = " & cboUserType.ItemData(cboUserType.ListIndex)
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                txtUserTypeDesc.Text = Trim(rst.Fields("USERTYPEDESC"))
                sTotal = Trim(rst.Fields("PERM"))
    '''            lblPermValue = Format(Total, "0,###")
    '''            Text1.Text = Total
    '''            Total = Val(Text1.Text)
                Call GenPerms(sTotal)
                sTotal = Trim(rst.Fields("PERM"))
                lblTotal = sTotal
    '''            Command1_Click
            Else
                txtUserTypeDesc.Text = ""
'''                sTotal = "0"
'''                Call GenPerms(sTotal)
            End If
            rst.Close
            Set rst = Nothing
            lblPermSet = "PROPOSED Permission Set for " & cboUser.Text
'''''            tvw1.Enabled = False
            cmdResetPerm.Enabled = True
        ElseIf cboUserType.Text = " " Then
            txtUserTypeDesc.Text = ""
            cmdResetPerm.Enabled = False
            sTotal = "0"
            lblTotal = sTotal
            Call GenPerms(sTotal)
    '''''        lblPermValue = ""
        End If
    End If
End Sub

Private Sub cboViewUsage_Click()
    Dim strSelect As String
    Dim sSource As String, sName As String, sMess As String, _
                sShow As String, sNewShow As String, sType As String
    Dim rst As ADODB.Recordset
    Dim iOff As Integer, iCase As Integer, iLen As Integer, iQty As Integer
    
    
    iQty = 0
    If cboViewUsage.Text = "" Then Exit Sub
    
    Me.MousePointer = 11
    
    cmdViewUsage.Enabled = True
    iCase = cboViewUsage.ItemData(cboViewUsage.ListIndex)
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
                    If sMess = "" Then
                        sMess = sName & vbNewLine
                    Else
                        sMess = sMess & vbNewLine & sName & vbNewLine
                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  " & _
                            CStr(rst.Fields("SHYR")) & " - " & _
                            Trim(rst.Fields("SHOW"))
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sMess = sMess & vbTab & sShow & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
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
                    If sMess = "" Then
                        sMess = sName & vbNewLine
                    Else
                        sMess = sMess & vbNewLine & sName & vbNewLine
                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  [ " & _
                            Trim(rst.Fields("GDESC")) & " ]"
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sMess = sMess & vbTab & sShow & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
        
        Case 12, 13, 14, 15
            iOff = iCase - 12
            strSelect = "SELECT ADDUSER, COUNT(GID) AS GCOUNT " & _
                        "From ANNOTATOR.GFX_MASTER " & _
                        "WHERE TO_CHAR(ADDDTTM, 'DD-MON-YY') = '" & _
                        UCase(Format(DateAdd("d", (iCase - 12) * -1, Date), "DD-MMM-YY")) & "' " & _
                        "GROUP BY ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                iQty = iQty + rst.Fields("GCOUNT")
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
            iOff = iCase - 18
            strSelect = "SELECT DISTINCT LL.UPDUSER, AU.USERTYPE " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, IGLPROD.IGL_USER_APP_R US, ANNOTATOR.ANO_USERTYPE AU " & _
                        "WHERE TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "AND LL.USER_SEQ_ID = US.USER_SEQ_ID " & _
                        "AND US.APP_ID = 1002 " & _
                        "AND US.USER_PERMISSION_ID = AU.USERTYPEID " & _
                        "ORDER BY AU.USERTYPE, LL.UPDUSER"
            sMess = "": sType = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sType <> UCase(Trim(rst.Fields("USERTYPE"))) Then
                    sType = UCase(Trim(rst.Fields("USERTYPE")))
                    If sMess = "" Then
                        sMess = sType & vbNewLine
                    Else
                        sMess = sMess & vbNewLine & sType & vbNewLine
                    End If
                    sShow = ""
                End If
                iQty = iQty + 1
                sMess = sMess & vbTab & UCase(Trim(rst.Fields("UPDUSER"))) & vbNewLine
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
        Case 22, 23, 24, 25
            iOff = iCase - 22
            strSelect = "SELECT DISTINCT LL.ADDUSER " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL " & _
                        "WHERE LOCKREFSOURCE = 'DIL_OPEN' " & _
                        "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "ORDER BY LL.ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sMess = sMess & UCase(Trim(rst.Fields("ADDUSER"))) & vbNewLine
                iQty = iQty + 1
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
        Case 26, 27
            If iCase = 26 Then iOff = 7 Else iOff = 30
            strSelect = "SELECT DISTINCT LL.ADDUSER " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL " & _
                        "WHERE LL.LOCKREFSOURCE = 'DIL_OPEN' " & _
                        "AND LL.LOCKOPENDTTM >= SYSDATE-" & iOff & " " & _
                        "ORDER BY LL.ADDUSER"
            sMess = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                sMess = sMess & UCase(Trim(rst.Fields("ADDUSER"))) & vbNewLine
                iQty = iQty + 1
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
        Case 4, 5, 6, 7
            iOff = iCase - 4
            strSelect = "SELECT " & _
                        "LL.UPDUSER, C.ABALPH AS CLIENT, GM.GDESC " & _
                        "FROM ANNOTATOR.ANO_LOCKLOG LL, " & F0101 & " C, ANNOTATOR.GFX_MASTER GM " & _
                        "WHERE LOCKREFSOURCE = 'GFX_MASTER' " & _
                        "AND TO_CHAR(LL.LOCKOPENDTTM, 'DD-MON-YYYY') = TO_CHAR(SYSDATE-" & iOff & ", 'DD-MON-YYYY') " & _
                        "AND LL.LOCKREFID = GM.GID " & _
                        "AND GM.AN8_CUNO = C.ABAN8 " & _
                        "ORDER BY UPDUSER, CLIENT, GDESC"
            sName = "": sMess = "": sShow = ""
            Set rst = Conn.Execute(strSelect)
            Do While Not rst.EOF
                If sName <> UCase(Trim(rst.Fields("UPDUSER"))) Then
                    sName = UCase(Trim(rst.Fields("UPDUSER")))
                    iQty = iQty + 1
                    If sMess = "" Then
                        sMess = sName & vbNewLine
                    Else
                        sMess = sMess & vbNewLine & sName & vbNewLine
                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  [ " & _
                            Trim(rst.Fields("GDESC")) & " ]"
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sMess = sMess & vbTab & sShow & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing
            
        Case Is < 4
            iOff = iCase
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
                    If sMess = "" Then
                        sMess = sName & vbNewLine
                    Else
                        sMess = sMess & vbNewLine & sName & vbNewLine
                    End If
                    sShow = ""
                End If
                sNewShow = Trim(rst.Fields("CLIENT")) & "  " & _
                            CStr(rst.Fields("SHYR")) & " - " & _
                            Trim(rst.Fields("SHOW"))
                If sNewShow <> sShow Then
                    sShow = sNewShow
                    sMess = sMess & vbTab & sShow & vbNewLine
                End If
                rst.MoveNext
            Loop
            rst.Close: Set rst = Nothing

        
            
    End Select
    
    Me.MousePointer = 0
    
    With frmUsage
        .PassMess = sMess
        Select Case iCase
            Case 10: .PassTitle = "USAGE LOG: Floorplans opened during past 7 Days"
            Case 11: .PassTitle = "USAGE LOG: Graphics accessed during past 7 Days"
            Case 16: .PassTitle = "POSTING LOG: Quantity of Graphics posted during past 7 Days"
            Case 17: .PassTitle = "POSTING LOG: Quantity of Graphics posted during past 30 Days"
            Case 22, 23, 24, 25
                .PassTitle = "DIL USAGE LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
            Case 26: .PassTitle = "DIL USAGE LOG: Users accessing DIL during past 7 Days"
            Case 27: .PassTitle = "DIL USAGE LOG: Users accessing DIL during past 30 Days"
            Case Else
                .PassTitle = "USAGE LOG: " & Format(DateAdd("d", iOff * -1, Now), "DDDD, MMMM D, YYYY")
        End Select
        .PassQty = iQty
        .Show 1
    End With
    
'''    MsgBox sMess, vbInformation, "LOG: " & format(Now, "MMMM D, YYYY")
    
End Sub

Private Sub cmdAddToTree_Click()
    Dim nodX As Node
    Dim NewNode As String, sDesc As String, PNode As String
    Select Case strNewNodeLevel
        Case "A"
            NewNode = "N" & tvw1.Nodes.Count ''' strNewNodeLevel & tvw1.Nodes.Count
            sDesc = txtPerm.Text
            Set nodX = tvw1.Nodes.Add(, , NewNode, sDesc)
        Case Else
            NewNode = "N" & tvw1.Nodes.Count ''' strNewNodeLevel & tvw1.Nodes.Count
            sDesc = txtPerm.Text
            PNode = txtParent.Text
            Set nodX = tvw1.Nodes.Add(PNode, tvwChild, NewNode, sDesc)
            nodX.EnsureVisible
    End Select
    cmdReset.Enabled = True
    cmdSave.Enabled = True
'''    cmdCancel.Enabled = False
End Sub

Private Sub cmdNotify_Click(Index As Integer)
    Dim sMess As String
    Dim Resp As VbMsgBoxResult
    
    Select Case Index
        Case 0
            sMess = "Welcome!" & vbNewLine & vbNewLine & _
                        "You have been setup as a GPJ Annotator User.  " & _
                        "The Annotator will be available to you through your browser. " & _
                        "Click on the link below, or enter the following at the address line:" & vbNewLine & vbNewLine & _
                        vbTab & "http://gpjapps.gpjco.com" & vbNewLine & vbNewLine & _
                        "Once the URL is entered, you may save it in your 'Favorites' list, or create a Link to it.  " & _
                        "If you need assistance in doing this, contact the Help Desk." & vbNewLine & vbNewLine & _
                        "Use the following User Name and Password.  Note: There is a short delay during " & _
                        "login while the Annotator is authenticating itself to Novell.  Be patient.  " & _
                        "Once at the Annotator login screen, you may change your password.  " & _
                        "NOTE:  Changing your Annotator password, does not modify " & _
                        "your Global LINKS password." & vbNewLine & vbNewLine & _
                        vbTab & "User Name:" & vbTab & LCase(lblShort.Caption) & vbNewLine & _
                        vbTab & "Password:" & vbTab & sPass & vbNewLine & vbNewLine & _
                        "Once in the application, there is a selection of Help Files to instruct you " & _
                        "on its functionality.  If you experience any problems accessing " & _
                        "the GPJ Annotator, please contact the Help Desk and they will help you to resolve your issue." & _
                        vbNewLine & vbNewLine & "NOTE:  If your password is 'password', this is intended " & _
                        "as a temporary password only and you will be prompted to " & _
                        "change it before access is allowed.  Select 'Reset...' to change your password."
        Case 1
            If bGPJUser Then ''INTERNAL USER''
                sMess = "Welcome!" & vbNewLine & vbNewLine & _
                            "You have been setup as a GPJ Annotator User.  " & _
                            "The Annotator will be available to you from Global LINKS, " & _
                            "by selecting the 'Space Plan/Graphics' button on the opening page " & _
                            "(after initial Client selection)." & vbNewLine & vbNewLine & _
                            "Use the following User Name and Password.  Note: There may be a short delay during " & _
                            "login while the Annotator is authenticating itself to Novell.  Be patient." & _
                            vbNewLine & vbNewLine & _
                            vbTab & "User Name:" & vbTab & LCase(lblShort.Caption) & vbNewLine & _
                            vbTab & "Password:" & vbTab & "<Novell Password>" & vbNewLine & vbNewLine & _
                            "Once in the application, there is a selection of Help Files to instruct you " & _
                            "on its functionality.  If you experience any problems accessing " & _
                            "the GPJ Annotator, please contact the Help Desk and they will help you to resolve your issue." & _
                            vbNewLine & vbNewLine & vbNewLine & sLink
            Else ''EXTERNAL USER''
                sMess = "Welcome!" & vbNewLine & vbNewLine & _
                            "You have been setup as a GPJ Annotator User.  " & _
                            "The Annotator will be available to you from Global LINKS, " & _
                            "by selecting the 'Space Plan/Graphics' button on the opening page " & _
                            "(after initial Client selection)." & vbNewLine & vbNewLine & _
                            "Use the following User Name and Password.  Note: There may be a short delay during " & _
                            "login while the Annotator is authenticating itself to Novell.  Be patient.  " & _
                            "Once at the Annotator login screen, you may change your password.  " & _
                            "NOTE:  Changing your Annotator password, does not modify " & _
                            "your Global LINKS password." & vbNewLine & vbNewLine & _
                            vbTab & "User Name:" & vbTab & LCase(lblShort.Caption) & vbNewLine & _
                            vbTab & "Password:" & vbTab & sPass & vbNewLine & vbNewLine & _
                            "Once in the application, there is a selection of Help Files to instruct you " & _
                            "on its functionality.  If you experience any problems accessing " & _
                            "the GPJ Annotator, please contact the Help Desk and they will help you to resolve your issue." & _
                            vbNewLine & vbNewLine & "NOTE:  If your password is 'password', you will be prompted to " & _
                            "change it before access is allowed.  Select 'Reset...' to change your password." & _
                            vbNewLine & vbNewLine & vbNewLine
                If bICAUser Then
                    sMess = sMess & sGLLink
                Else
                    sMess = sMess & sLink
                End If
            End If
    End Select
    Resp = MsgBox(sMess, vbOKCancel, "Please Review Content...")
    If Resp = vbOK Then
        Screen.MousePointer = 11
        Call SendNotification(lUserID, sMess)
        cmdNotify(0).Visible = False: cmdNotify(1).Visible = False
        Screen.MousePointer = 0
    Else
        MsgBox "Notification Cancelled"
    End If

End Sub

Private Sub cmdOptions_Click()
    If mnuOptions.Visible = False Then _
        Me.PopupMenu mnuOptions, 0, cmdOptions.Left, cmdOptions.Top + cmdOptions.Height
End Sub

Private Sub cmdReport_Click()
    Dim sFile As String
    
    Me.MousePointer = 11
    sFile = GetUserList
    web1.Navigate sFile
    web1.Visible = True
    cmdOptions.Visible = True
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(True)
    If Not bENABLE_PRINTERS Then lblPrint.Visible = True
    '\\\\\ -------------------------------------------------------- /////
    
    Me.MousePointer = 0

End Sub

Private Sub cmdReset_Click()
    tvw1.Nodes.Clear
    txtPerm.Text = ""
    optParent(0).Value = True
    PopTree
    cmdTree.Caption = " Expand Tree"
    cmdReset.Enabled = False
    cmdSave.Enabled = False
'''    cmdCancel.Enabled = True
End Sub

Private Sub cmdResetPerm_Click()
    Dim strUpdate As String, strSelect As String
    Dim rst As ADODB.Recordset
    
    If bSaveable Then
        On Error Resume Next
        If txtPassword(1).Text <> "" And txtPassword(0).Text = txtPassword(1).Text Then
            strUpdate = "UPDATE " & IGLUserAR & " " & _
                        "SET USER_PERMISSION_ID = " & cboUserType.ItemData(cboUserType.ListIndex) & ", " & _
                        "PCODE = '" & Trim(txtPassword(1).Text) & "', " & _
                        "UPDUSER = '" & Left(DeGlitch(LogName), 20) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                        "WHERE APP_ID = 1002 " & _
                        "AND USER_SEQ_ID = " & lUserID
        Else
            strUpdate = "UPDATE " & IGLUserAR & " " & _
                        "SET USER_PERMISSION_ID = " & cboUserType.ItemData(cboUserType.ListIndex) & ", " & _
                        "UPDUSER = '" & Left(DeGlitch(LogName), 20) & "', " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                        "WHERE APP_ID = 1002 " & _
                        "AND USER_SEQ_ID = " & lUserID
        End If
        Conn.Execute (strUpdate)
        If Err Then
            MsgBox "Error:  " & Err.Description, vbCritical, "An error has occurred..."
        Else
            ''CHECK IF USER IS GPJ OR EXTERNAL USER''
            bICAUser = False
            strSelect = "SELECT U.EMPLOYER, A.USER_PERMISSION_ID " & _
                        "FROM " & IGLUser & " U, " & IGLUserAR & " A " & _
                        "WHERE U.USER_SEQ_ID = " & lUserID & " " & _
                        "AND U.USER_SEQ_ID = A.USER_SEQ_ID " & _
                        "AND A.APP_ID = 1002"
            Set rst = Conn.Execute(strSelect)
            If Not rst.EOF Then
                If UCase(Left(rst.Fields("EMPLOYER"), 3)) = "GPJ" Then
                    bGPJUser = True
                Else
                    bGPJUser = False
                End If
            Else
                bGPJUser = False
                If rst.Fields("USER_PERMISSION_ID") = 6553 Then bICAUser = True
            End If
            rst.Close: Set rst = Nothing
            
            cboPerm.RemoveItem (cboUser.ListIndex)
            cboPerm.AddItem cboUserType.List(cboUserType.ListIndex), cboUser.ListIndex
            cboPerm.ItemData(cboUser.ListIndex) = cboUserType.ItemData(cboUserType.ListIndex)
            sPass = txtPassword(0).Text
            txtPassword(0).Text = ""
            txtPassword(1).Text = ""
            MsgBox "Data has been updated for " & cboUser.Text & ".", vbInformation, "Save Complete..."
            cmdNotify(0).Visible = True: cmdNotify(1).Visible = True
        End If
        If UserID = lUserID Then Call frmStartUp.SetPermissionBools(sTotal)
    End If
End Sub

Private Sub cmdSave_Click()
    Dim strInsert As String, sParent As String
    
    If txtParent.Text = "<Select Parent Node Above>" Then sParent = "" Else sParent = txtParent.Text
    strInsert = "INSERT INTO " & ANOPerm & " " & _
                "(NODEID, NODELEVEL, NODEPARENT, NODEDESC, " & _
                "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                "VALUES " & _
                "(" & tvw1.Nodes.Count - 1 & ", '" & strNewNodeLevel & "', '" & _
                Trim(sParent) & "', '" & DeGlitch(txtPerm.Text) & "', " & _
                "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
'''    fraAddNewPerm.Visible = False
    txtPerm.Text = ""
    optParent(0).Value = True
'''    cmdCancel.Enabled = True
End Sub

Private Sub cmdSavePermDesc_Click()
    Dim strUpdate As String
    
    Debug.Print "Current NodeKey = " & CurrNode
    Debug.Print "Current Description = " & tvw1.Nodes(CurrNode).Text
    If Trim(txtPermDesc.Text) <> "" Then
        On Error Resume Next
        strUpdate = "UPDATE " & ANOPerm & " " & _
                    "SET NODEDESC = '" & DeGlitch(txtPermDesc.Text) & "', " & _
                    "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                    "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                    "WHERE NODEID = " & Mid(CurrNode, 2)
        Conn.Execute (strUpdate)
        If Err Then
            MsgBox "Error:  " & Err.Description, vbCritical, "An error has occurred..."
        Else
            tvw1.Nodes(CurrNode).Text = txtPermDesc.Text
            lblCurrNode = "Key: " & CurrNode & "   Desc: " & txtPermDesc.Text
        End If
    End If
End Sub

Private Sub cmdSaveUserType_Click()
    Dim strInsert As String, strSelect As String
    Dim rstL As ADODB.Recordset
    Dim lUID As Long
    
'''    CalcTotal
'''    If sTotal > "0" Then
        Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
        lUID = rstL.Fields("nextval")
        rstL.Close: Set rstL = Nothing
        strInsert = "INSERT INTO " & ANOUserType & " " & _
                    "(USERTYPEID, USERTYPE, " & _
                    "USERTYPEDESC, USERTYPEVALUE, " & _
                    "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                    "VALUES " & _
                    "(" & lUID & ", '" & DeGlitch(Trim(txtUserType.Text)) & "', " & _
                    "'" & DeGlitch(Trim(txtDescUserType(0).Text)) & "', 0, " & _
                    "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, '" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
        Conn.Execute (strInsert)
        Call WriteUserTypePermToOracle(lUID)
        
        txtUserType.Text = ""
        PopUserTypes
'''    End If
End Sub

'''Private Sub Command1_Click()
'''    Total = Val(Text1.Text)
'''    Call GenPerms(Total)
'''End Sub

'''Private Sub Command2_Click()
'''    Dim i As Integer
'''    For i = tvw1.Nodes.Count To 1 Step -1
'''        tvw1.Nodes(i).Checked = False
'''    Next i
'''    For i = 0 To opt1.Count - 1
'''        opt1(i).Value = False
'''    Next i
'''    Total = 0
''''''''    Me.Caption = "Total: " & Format(Total, "#,##0")
'''    Text1.Text = ""
'''End Sub

Private Sub cmdTree_Click()
    Dim i As Integer
    Dim bool As Boolean
    tvw1.Visible = False
    If cmdTree.Caption = " Expand Tree" Then
        bool = True
        cmdTree.Caption = " Colapse Tree"
    Else
        bool = False
        cmdTree.Caption = " Expand Tree"
    End If
    For i = 1 To tvw1.Nodes.Count
        tvw1.Nodes(i).Expanded = bool
    Next i
    tvw1.Visible = True
    tvw1.Nodes(1).EnsureVisible
End Sub

Private Sub PopTree()
    Dim nodX As Node
    Dim i As Integer
    Dim rst As ADODB.Recordset
    Dim strSelect As String, sDesc As String
    Dim ANode As String, BNode As String, PNode As String
    Calcing = False
    
    
    tvw1.Nodes.Clear
    ANode = "": BNode = "": PNode = ""
    
    strSelect = "SELECT * FROM " & ANOPerm & " ORDER BY NODELEVEL, UPPER(NODEDESC), NODEID"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case rst.Fields("NODELEVEL")
            Case "A"
                ANode = "N" & rst.Fields("NODEID") ''''' rst.Fields("NODELEVEL") & rst.Fields("NODEID")
                If bShowNodes Then
                    sDesc = "[ " & rst.Fields("NODEID") & " ]  " & Trim(rst.Fields("NODEDESC"))
                Else
                    sDesc = Trim(rst.Fields("NODEDESC"))
                End If
                Set nodX = tvw1.Nodes.Add(, , ANode, sDesc)
            Case Else
                BNode = "N" & rst.Fields("NODEID") ''''' rst.Fields("NODELEVEL") & rst.Fields("NODEID")
                If bShowNodes Then
                    sDesc = "[ " & rst.Fields("NODEID") & " ]  " & Trim(rst.Fields("NODEDESC"))
                Else
                    sDesc = Trim(rst.Fields("NODEDESC"))
                End If
'''                sDesc = "[ " & rst.FIELDS("NODEID") & " ]  " & Trim(rst.FIELDS("NODEDESC"))
'''                sDesc = Trim(rst.FIELDS("NODEDESC"))
                PNode = Trim(rst.Fields("NODEPARENT"))
                Set nodX = tvw1.Nodes.Add(PNode, tvwChild, BNode, sDesc)
        End Select
       rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub cmdUpdateUserType_Click()
    Dim strUpdate As String
    On Error Resume Next
    strUpdate = "UPDATE " & ANOUserType & " " & _
                "SET USERTYPE = '" & DeGlitch(Trim(cboEditUserType.Text)) & "', " & _
                "USERTYPEDESC = '" & DeGlitch(Trim(txtDescUserType(1).Text)) & "', " & _
                "USERTYPEVALUE = 0, " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 24)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE USERTYPEID = " & lblUID.Caption
    Conn.Execute (strUpdate)
    If Err Then
        MsgBox "Error:  " & Err.Description, vbCritical, "An error has occurred..."
    Else
        Call WriteUserTypePermToOracle(CLng(lblUID.Caption))
        MsgBox "User Type Updated.", vbExclamation, "Save Complete..."
    End If
End Sub

Private Sub cmdViewUsage_Click()
'''    cboViewUsage_Click
    frmUsageGannt.Show 1, Me
End Sub

Private Sub Form_DblClick()
'''    Dim iVal As Integer
'''    iVal = Printer.Orientation
'''    Printer.Orientation = 2
'''    Me.PrintForm
'''    Printer.EndDoc
'''    Printer.Orientation = iVal
    
     '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    If bDo_Printer_Check Then bDo_Printer_Check = Check_Printers(True)
    If Not bENABLE_PRINTERS Then GoTo ErrorHandler
    '\\\\\ -------------------------------------------------------- /////
    
    Screen.MousePointer = 11
    On Error GoTo ErrorHandler
    Printer.Orientation = 2
    PrintForm
    Printer.EndDoc
    bSaveable = False
    Screen.MousePointer = 0
Exit Sub
ErrorHandler:
    Screen.MousePointer = 0
    MsgBox "Temporarily, this form cannot be printed until the correct Printer Drivers " & _
                "have been installed on the Server.", vbExclamation, "Sorry..."
End Sub

Private Sub Form_Load()
    Dim ConnStr As String
    
'''    strHTMLPath = "\\DETMSFS01\GPJAnnotator\Support\HTML\"
    bShowNodes = False
    bResetting = False
    strNewNodeLevel = "A"
    sst1.Tab = 0
    sTotal = "0"
    lblTotal = sTotal
    PopUsers
    PopUserTypes
    PopTree
    If bPerm(46) Then fraPermMaint.Visible = True Else fraPermMaint.Visible = False
'''    If bPerm(54) Then cmdReport.Visible = True Else cmdReport.Visible = False
    web1.Top = 180: web1.Left = 180
'''    fraAddNewPerm.Top = 840: fraAddNewPerm.Left = 180
    
    With cboViewUsage
        .AddItem "Floorplans - Today"
        .ItemData(.NewIndex) = 0
        .AddItem "Floorplans - Yesterday"
        .ItemData(.NewIndex) = 1
        .AddItem "Floorplans - 2 Days Ago"
        .ItemData(.NewIndex) = 2
        .AddItem "Floorplans - 3 Days Ago"
        .ItemData(.NewIndex) = 3
        .AddItem "Floorplans - For Week"
        .ItemData(.NewIndex) = 10
        
        .AddItem "Graphics - Today"
        .ItemData(.NewIndex) = 4
        .AddItem "Graphics - Yesterday"
        .ItemData(.NewIndex) = 5
        .AddItem "Graphics - 2 Days Ago"
        .ItemData(.NewIndex) = 6
        .AddItem "Graphics - 3 Days Ago"
        .ItemData(.NewIndex) = 7
        .AddItem "Graphics - For Week"
        .ItemData(.NewIndex) = 11
        
        .AddItem "Posted Gfx - Today"
        .ItemData(.NewIndex) = 12
        .AddItem "Posted Gfx - Yesterday"
        .ItemData(.NewIndex) = 13
        .AddItem "Posted Gfx - 2 Days Ago"
        .ItemData(.NewIndex) = 14
        .AddItem "Posted Gfx - 3 Days Ago"
        .ItemData(.NewIndex) = 15
        .AddItem "Posted Gfx - For Week"
        .ItemData(.NewIndex) = 16
        .AddItem "Posted Gfx - For Month"
        .ItemData(.NewIndex) = 17
        
        .AddItem "AnnoUsers - Today"
        .ItemData(.NewIndex) = 18
        .AddItem "AnnoUsers - Yesterday"
        .ItemData(.NewIndex) = 19
        .AddItem "AnnoUsers - 2 Days Ago"
        .ItemData(.NewIndex) = 20
        .AddItem "AnnoUsers - 3 Days Ago"
        .ItemData(.NewIndex) = 21
        
        .AddItem "DIL Users - Today"
        .ItemData(.NewIndex) = 22
        .AddItem "DIL Users - Yesterday"
        .ItemData(.NewIndex) = 23
        .AddItem "DIL Users - 2 Days Ago"
        .ItemData(.NewIndex) = 24
        .AddItem "DIL Users - 3 Days Ago"
        .ItemData(.NewIndex) = 25
        .AddItem "DIL Users - For Week"
        .ItemData(.NewIndex) = 26
        .AddItem "DIL Users - For Month"
        .ItemData(.NewIndex) = 27
    End With
        
End Sub

Private Sub Form_Resize()
    Select Case Me.WindowState
        Case 0
            tvw1.Width = 5835
            tvw1.Height = 8055
        Case 2
            tvw1.Width = Me.Width - tvw1.Left - 285
            tvw1.Height = Me.Height - tvw1.Top - 585
    End Select
    
    web1.Width = Me.ScaleWidth - 360
    
    web1.Height = Me.ScaleHeight - 360
    cmdOptions.Top = web1.Top + 120
    cmdOptions.Left = web1.Left + web1.Width - 360 - cmdOptions.Width
    
End Sub

'''Private Sub Form_Unload(Cancel As Integer)
'''    Dim sChk As String
    
'''    sChk = Dir(strHTMLPath & "Users.htm", vbNormal)
'''    If sChk <> "" Then Kill strHTMLPath & "Users.htm"
'''End Sub

Private Sub lblShort_Click()
    If lblShort.Caption <> "" Then
        frmUserLog.PassUser = cboUser.Text
        frmUserLog.PassUserID = lUserID
        frmUserLog.Show 1
    End If
End Sub

Private Sub mnuViewClose_Click()
    Dim sChk As String
    web1.Visible = False
    cmdOptions.Visible = False
    
    '///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
    lblPrint.Visible = False
    '\\\\\ -------------------------------------------------------- /////
    
    sChk = Dir(strHTMLPath & "Users.htm", vbNormal)
    If sChk <> "" Then Kill strHTMLPath & "Users.htm"
End Sub

Private Sub mnuViewUser_Click()
    Dim sFile As String
    
    Me.MousePointer = 11
    sFile = GetUserList
    web1.Navigate sFile
    web1.Visible = True
    cmdOptions.Visible = True
    mnuViewUser.Checked = True
    mnuViewUserType.Checked = False
    Me.MousePointer = 0
End Sub

Private Sub mnuViewUserType_Click()
    Dim sFile As String
    
    Me.MousePointer = 11
    sFile = GetUserTypeList
    web1.Navigate sFile
    web1.Visible = True
    cmdOptions.Visible = True
    mnuViewUser.Checked = False
    mnuViewUserType.Checked = True
    Me.MousePointer = 0
End Sub


Private Sub optParent_Click(Index As Integer)
    Select Case Index
        Case 0
            txtParent.Text = ""
            txtParent.Visible = False
            GettingParent = False
'''''            tvw1.Enabled = False
        Case 1
            txtParent.Text = "<Select Parent Node Above>"
            txtParent.Visible = True
            GettingParent = True
'''''            tvw1.Enabled = True
    End Select
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    Select Case sst1.Tab
        Case 0
            cboUser.Text = " "
            cboUserType.Text = " "
            sTotal = "0"
            lblTotal = sTotal
            lblPermValue = ""
            txtPassword(0).Text = ""
'''            fraAddNewPerm.Visible = True
            strNewNodeLevel = "A"
'''            txtPerm.SetFocus
'''''            tvw1.Enabled = False
        Case 1
            If CurrNode <> "" Then
                txtPermDesc.Text = tvw1.Nodes(CurrNode).Text
            Else
                txtPermDesc.Text = ""
            End If
'''''            tvw1.Enabled = True
        Case 2, 3
'''''            tvw1.Enabled = True
    End Select
    
End Sub

'''Private Sub tvw1_AfterLabelEdit(Cancel As Integer, NewString As String)
'''    Dim strUpdate As String
'''    Dim Resp As VbMsgBoxResult
'''    Resp = MsgBox("Are you certain you want to change the Permission Description?" & _
'''                vbCr & vbCr & "New Description:  " & NewString, vbYesNoCancel, "Confirming...")
'''    If Resp = vbYes Then
'''        strUpdate = "UPDATE " & ANOPerm & " " & _
'''                    "SET NODEDESC = '" & NewString & "' " & _
'''                    "WHERE NODEID = " & CInt(Mid(CurrNode, 2))
'''        Conn.Execute (strUpdate)
'''    Else
'''        Cancel = True
'''    End If
'''    tvw1.LabelEdit = tvwManual
'''End Sub

Private Sub tvw1_DblClick()
    Dim bool As Boolean
    Dim i As Integer
    
    If bShowNodes Then bShowNodes = False Else bShowNodes = True
    tvw1.Visible = False
    PopTree
    Call GenPerms(sTotal)
    If cmdTree.Caption = " Colapse Tree" Then
        bool = True
    Else
        bool = False
    End If
    For i = 1 To tvw1.Nodes.Count
        tvw1.Nodes(i).Expanded = bool
    Next i
    tvw1.Visible = True
End Sub

Private Sub TVW1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim iStart As Integer, i As Integer, iCnt As Integer, iLI As Integer
    Dim NodeChild As Node
    Dim Total2 As Double
    Dim lErr As Long
    
    lErr = LockWindowUpdate(Me.hwnd)
    
    
    Debug.Print Node.Key
'''''    tvw1.Visible = False
    If Calcing = False Then
        Call ResetNodes(Node)
    
''''''''    On Error GoTo OutOfParents
''''''''    If Calcing = False Then
''''''''        If Node.Checked = True Then
'''''''''''            Total = Total + (2 ^ (CLng(Mid(Node.Key, 2))))
''''''''            Debug.Print "Sibling:  " & Node.FirstSibling.Index
''''''''            If Node.FirstSibling.Index <> 1 Then
''''''''                If Node.Parent.Checked = False Then
''''''''                    Node.Parent.Checked = True
'''''''''''                    Total = Total + (2 ^ (CLng(Mid(Node.Parent.Key, 2))))
''''''''                    If Node.Parent.Parent.Checked = False Then
''''''''                        Node.Parent.Parent.Checked = True
'''''''''''                        Total = Total + (2 ^ (CLng(Mid(Node.Parent.Parent.Key, 2))))
''''''''                        If Node.Parent.Parent.Parent.Checked = False Then
''''''''                            Node.Parent.Parent.Parent.Checked = True
'''''''''''                            Total = Total + (2 ^ (CLng(Mid(Node.Parent.Parent.Parent.Key, 2))))
''''''''                            If Node.Parent.Parent.Parent.Parent.Checked = False Then
''''''''                                Node.Parent.Parent.Parent.Parent.Checked = True
'''''''''''                                Total = Total + (2 ^ (CLng(Mid(Node.Parent.Parent.Parent.Parent.Key, 2))))
''''''''                                If Node.Parent.Parent.Parent.Parent.Parent.Checked = False Then
''''''''                                    Node.Parent.Parent.Parent.Parent.Parent.Checked = True
'''''''''''                                    Total = Total + (2 ^ (CLng(Mid(Node.Parent.Parent.Parent.Parent.Parent.Key, 2))))
''''''''                                    If Node.Parent.Parent.Parent.Parent.Parent.Parent.Checked = False Then
''''''''                                        Node.Parent.Parent.Parent.Parent.Parent.Parent.Checked = True
'''''''''''                                        Total = Total + (2 ^ (CLng(Mid(Node.Parent.Parent.Parent.Parent.Parent.Parent.Key, 2))))
''''''''                                    End If
''''''''                                End If
''''''''                            End If
''''''''                        End If
''''''''                    End If
''''''''                End If
''''''''            End If
''''''''OutOfParents:
''''''''            Node.Expanded = True
''''''''        Else
'''''''''''            Total = Total - (2 ^ (CLng(Mid(Node.Key, 2))))
''''''''            If Node.Children > 0 Then
''''''''                iCnt = Node.Children
''''''''                If iCnt <> 0 Then
''''''''                    Set NodeChild = Node.Child.FirstSibling
''''''''                    iLI = NodeChild.LastSibling.Index
''''''''                    If NodeChild.Checked = True Then
''''''''                        NodeChild.Checked = False
'''''''''''                        Total = Total - (2 ^ (CLng(Mid(NodeChild.Key, 2))))
''''''''                    End If
''''''''                    If NodeChild.Index <> iLI Then
''''''''                        Do
''''''''                            Set NodeChild = NodeChild.Next
''''''''                            If NodeChild.Checked = True Then
''''''''                                NodeChild.Checked = False
'''''''''''                                Total = Total - (2 ^ (CLng(Mid(NodeChild.Key, 2))))
''''''''                            End If
''''''''                       Loop Until NodeChild.Index = iLI
''''''''                    End If
''''''''                End If
''''''''
'''''''''''                iStart = Node.Index
'''''''''''                i = 1
'''''''''''                Do While iCnt > 0
'''''''''''                    If tvw1.Nodes(iStart + i).Parent.Key = Node.Key Then
'''''''''''                        If tvw1.Nodes(iStart + i).Checked = True Then
'''''''''''                            tvw1.Nodes(iStart + i).Checked = False
'''''''''''                            Total = Total - (2 ^ (CLng(Mid(tvw1.Nodes(iStart + i).Key, 2))))
'''''''''''                        End If
'''''''''''                        iCnt = iCnt - 1
'''''''''''                    End If
'''''''''''                    i = i + 1
'''''''''''                Loop
''''''''            End If
''''''''        End If
'''''''''''''        lblPermValue = Format(Total, "#,##0")
'''''''''''''        Me.Caption = "Total: " & Format(Total, "#,##0")
        
'''        Total = 0
'''        lblTotal = Total
'''        For i = 1 To tvw1.Nodes.Count
'''            If tvw1.Nodes(i).Checked = True Then Total = Total + (2 ^ (CDbl(Mid(tvw1.Nodes(i).Key, 2))))
'''        Next i
'''        lblPermValue = format(Total, "#,##0")
'''        lblTotal = Total
'''        Debug.Print "Total = " & format(Total, "#,##0")
        
        '///// ALERT IF TURNING OFF PERMISSION TO USE THIS FORM \\\\\
        If Node.Key = "N41" And Node.Checked = False And lUserID = UserID Then
            MsgBox "BE AWARE:  You are not able to save this permission level for yourself, " & _
                        "If you were able to, you would not be able to re-enter this screen after closing." & _
                        vbNewLine & vbNewLine & "Another User with this permission level " & _
                        "is required to reset this permission for you.", vbCritical, "Resetting Check..."
            bSaveable = False
        Else
            bSaveable = True
        End If
    End If
    If bSaveable = False Then Call ResetIt
    lErr = LockWindowUpdate(0)
'''''    tvw1.Visible = True
End Sub

Public Sub GenPerms(sTotal As String)
    Dim i As Integer, iInt As Integer
    Dim sNode As String, strSelect As String, subT As String
    Dim iCnt As Integer
    Dim rst As ADODB.Recordset
    
    lblPermValue = sTotal
    Calcing = True
    
    For i = tvw1.Nodes.Count To 1 Step -1
        tvw1.Nodes(i).Checked = False
    Next i
    
    subT = sTotal
    For i = tvw1.Nodes.Count To 1 Step -1
        strSelect = "SELECT TO_CHAR(MOD(" & sTotal & ", POWER(2, " & i - 1 & ")), '" & String(38, "9") & "') FROM DUAL"
        Set rst = Conn.Execute(strSelect)
        subT = Trim(rst.Fields(0))
        sNode = CStr("N" & i - 1)
        If subT = sTotal Then
            tvw1.Nodes(sNode).Checked = False
        Else
            tvw1.Nodes(sNode).Checked = True
            sTotal = subT
        End If
    Next i
    
    
'''    For i = tvw1.Nodes.Count To 1 Step -1
'''        iInt = Int(subT / (2 ^ (i - 1)))
'''        sNode = CStr("N" & i - 1)
'''        tvw1.Nodes(sNode).Checked = CBool(iInt * -1)
'''        subT = subT - ((Int(subT / (2 ^ (i - 1)))) * (2 ^ (i - 1)))
'''    Next i
    Calcing = False
'''    lblPermValue = Format(Total, "0,###")
End Sub

Private Sub tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, iStr As Integer
    Debug.Print Node.Index
    Debug.Print Node.Key
    CurrNode = Node.Key
    lblCurrNode = "Key: " & Node.Key & "   Desc: " & Node.Text
    txtPermDesc.Text = Node.Text
    If GettingParent = True Then
        txtParent.Text = Node.Key
        i = 0: iStr = 1
        Do While InStr(iStr, Node.FullPath, "\") <> 0
            i = i + 1
            iStr = InStr(iStr, Node.FullPath, "\") + 1
        Loop
        strNewNodeLevel = Chr(65 + i + 1)
        '''Node.FullPath
        '''strNewNodeLevel = Chr(Asc(Left(Node.Key, 1)) + 1)
        Debug.Print "New Level = " & strNewNodeLevel
    End If
End Sub

Private Sub txtDescUserType_Change(Index As Integer)
    Select Case Index
        Case 0
            If txtUserType.Text <> "" And txtDescUserType(0).Text <> "" Then
                cmdSaveUserType.Enabled = True
            Else
                cmdSaveUserType.Enabled = False
            End If
    End Select
End Sub

'''Private Sub txtDescUserType_Change()
'''End Sub

Private Sub txtPassword_Change(Index As Integer)
    Select Case Index
        Case 0
            If Len(txtPassword(0).Text) >= 8 And Len(txtPassword(0).Text) <= 16 Then
                txtPassword(1).Enabled = True
'''                lblPassword.Visible = True
            Else
                txtPassword(1).Enabled = False
'''                lblPassword.Visible = False
            End If
        Case 1
            If txtPassword(0).Text = txtPassword(1).Text Then
                cmdResetPerm.Enabled = True
            Else
                cmdResetPerm.Enabled = False
            End If
    End Select
End Sub

Private Sub txtPassword_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            If KeyAscii = 13 Then txtPassword(1).SetFocus
        Case 1
            If KeyAscii = 13 And txtPassword(0).Text = txtPassword(1).Text Then
                cmdResetPerm.Enabled = True
            Else
                cmdResetPerm.Enabled = False
            End If
    End Select
End Sub

Private Sub txtPerm_Change()
    If Trim(txtPerm.Text) <> "" Then
        cmdAddToTree.Enabled = True
    Else
        cmdAddToTree.Enabled = False
    End If
End Sub

Public Sub PopUsers()
    Dim rst As ADODB.Recordset
    Dim strSelect As String
    Dim dPerm As String
    
    cboUser.Clear
    cboUser.AddItem " "
    cboPerm.AddItem " "
'''''    strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, " & _
'''''                "TO_CHAR(A.USER_PERMISSION_ID, '" & String(38, "9") & "') AS PERM " & _
'''''                "FROM " & IGLUserAR & " A, " & IGLUser & " U " & _
'''''                "WHERE A.APP_ID = 1002 " & _
'''''                "AND A.PERMISSION_STATUS = 1 " & _
'''''                "AND A.USER_SEQ_ID = U.USER_SEQ_ID " & _
'''''                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    strSelect = "SELECT U.USER_SEQ_ID, U.NAME_LAST, U.NAME_FIRST, " & _
                "NVL(A.USER_PERMISSION_ID, 0) AS PERMID, NVL(UT.USERTYPE, ' ') AS UTYPE " & _
                "FROM " & IGLUserAR & " A, " & IGLUser & " U, " & ANOUserType & " UT " & _
                "WHERE A.APP_ID = 1002 " & _
                "AND A.PERMISSION_STATUS = 1 " & _
                "AND U.USER_STATUS = 1 " & _
                "AND A.USER_PERMISSION_ID = UT.USERTYPEID (+) " & _
                "AND A.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"

    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
'''''        If IsNull(rst.FIELDS("PERM")) Then dPerm = "0" Else dPerm = Trim(rst.FIELDS("PERM"))
        cboUser.AddItem UCase(Trim(rst.Fields("NAME_FIRST"))) & " " & UCase(Trim(rst.Fields("NAME_LAST")))
        cboUser.ItemData(cboUser.NewIndex) = rst.Fields("USER_SEQ_ID")
        cboPerm.AddItem Trim(rst.Fields("UTYPE"))
        cboPerm.ItemData(cboPerm.NewIndex) = rst.Fields("PERMID")
'''''        cboUser.ItemData(cboUser.NewIndex) = dPerm
'''''        Debug.Print Trim(rst.FIELDS("NAME_FIRST")) & " " & Trim(rst.FIELDS("NAME_LAST")) & _
'''''                    " (" & rst.FIELDS("USER_SEQ_ID") & ") -- " & dPerm
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub PopUserTypes()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    cboUserType.Clear
    cboEditUserType.Clear
    cboUserType.AddItem " "
    strSelect = "SELECT * FROM " & ANOUserType & " ORDER BY USERTYPE"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        cboUserType.AddItem Trim(rst.Fields("USERTYPE"))
        cboUserType.ItemData(cboUserType.NewIndex) = rst.Fields("USERTYPEID")
        cboEditUserType.AddItem Trim(rst.Fields("USERTYPE"))
        cboEditUserType.ItemData(cboEditUserType.NewIndex) = rst.Fields("USERTYPEID")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub ClearNodes()
    Dim i As Integer
    For i = 1 To tvw1.Nodes.Count
        tvw1.Nodes.Item(i).Checked = False
    Next i
    lblPermSet = ""
    lblPermValue = ""
    sTotal = "0"
    lblTotal = sTotal
End Sub

'''Function DeGlitch(sName As String) As String
'''    Dim iStr As Integer
'''    iStr = 1
'''    Do While InStr(iStr, sName, "'") <> 0
'''        sName = Left(sName, InStr(iStr, sName, "'")) & "'" & Mid(sName, InStr(iStr, sName, "'") + 1)
'''        iStr = InStr(iStr, sName, "'") + 2
'''    Loop
'''    DeGlitch = sName
'''End Function

Private Sub txtUserType_Change()
    If txtUserType.Text <> "" And txtDescUserType(0).Text <> "" Then
        cmdSaveUserType.Enabled = True
    Else
        cmdSaveUserType.Enabled = False
    End If
End Sub

Public Sub ResetIt()
    tvw1.Nodes("N41").Checked = True
End Sub

Public Sub ResetNodes(ByVal Node As MSComctlLib.Node)
    Dim nod1 As Node, nod2 As Node, nod3 As Node, nod4 As Node, nod5 As Node
    Dim i As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
    
    Debug.Print Node.Key
    On Error GoTo OutOfNodes
    If Node.Checked = False Then
        If Node.Children > 0 Then
            i1 = 0
            Do While i1 < Node.Children
                If i1 = 0 Then Set nod1 = Node.Child Else Set nod1 = nod1.Next
                nod1.Checked = False
                If nod1.Children > 0 Then
                    i2 = 0
                    Do While i2 < nod1.Children
                        If i2 = 0 Then Set nod2 = nod1.Child Else Set nod2 = nod2.Next
                        nod2.Checked = False
                        If nod2.Children > 0 Then
                            i3 = 0
                            Do While i3 < nod2.Children
                                If i3 = 0 Then Set nod3 = nod2.Child Else Set nod3 = nod3.Next
                                nod3.Checked = False
                                If nod3.Children > 0 Then
                                    i4 = 0
                                    Do While i4 < nod3.Children
                                        If i4 = 0 Then Set nod4 = nod3.Child Else Set nod4 = nod4.Next
                                        nod4.Checked = False
                                        If nod4.Children > 0 Then
                                            i5 = 0
                                            Do While i5 < nod4.Children
                                                If i5 = 0 Then Set nod5 = nod4.Child Else Set nod5 = nod5.Next
                                                nod5.Checked = False
                                                i5 = i5 + 1
                                            Loop
                                        End If
                                        i4 = i4 + 1
                                    Loop
                                End If
                                i3 = i3 + 1
                            Loop
                        End If
                        i2 = i2 + 1
                    Loop
                End If
                i1 = i1 + 1
            Loop
        End If
    ElseIf Node.Checked = True Then
        Node.Parent.Checked = True
        Node.Parent.Parent.Checked = True
        Node.Parent.Parent.Parent.Checked = True
        Node.Parent.Parent.Parent.Parent.Checked = True
        Node.Parent.Parent.Parent.Parent.Parent.Checked = True
        Node.Parent.Parent.Parent.Parent.Parent.Parent.Checked = True
        Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Checked = True
    End If
OutOfNodes:

End Sub

Public Sub SendNotification(lUser As Long, sMess As String)
    Dim MessHdr As String, sAddress As String
    Dim i As Integer, iAdd As Integer
    Dim sList As String, sIntro As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    sAddress = ""
    MessHdr = "GPJ Annotator Access"
    strSelect = "SELECT EMAIL_ADDRESS FROM " & IGLUser & " " & _
                "WHERE USER_SEQ_ID = " & lUser
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        If IsNull(rst.Fields("EMAIL_ADDRESS")) Then
            sAddress = ""
        Else
            sAddress = Trim(rst.Fields("EMAIL_ADDRESS"))
        End If
    End If
    rst.Close: Set rst = Nothing
    If sAddress = "" Then GoTo GetOut
        
    '///// EXECUTE E-MAIL \\\\\
''    Dim myNotes As New Domino.NotesSession
''    Dim myDB As New Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    
''''    myNotes.Initialize
'''    On Error Resume Next
'''    If sNOTESID = "GANNOTAT" Then
'''        myNotes.Initialize (sNOTESPASSWORD)
'''    Else
'''        If sNOTESPASSWORD = "" Then
'''            ''GET PASSWORD''
'''TryPWAgain:
'''            frmGetPassword.Show 1, Me
'''            Select Case sNOTESPASSWORD
'''                Case "_CANCEL"
'''                    sNOTESPASSWORD = ""
'''                    MsgBox "No email will be sent", vbExclamation, "User Canceled..."
'''                    Set myNotes = Nothing
'''                    Set myDB = Nothing
'''                Case Else
'''                    Err.Clear
'''                    myNotes.Initialize (sNOTESPASSWORD)
'''                    If Err Then
'''                        Err.Clear
'''                        GoTo TryPWAgain
'''                    End If
'''            End Select
'''        Else
'''            myNotes.Initialize (sNOTESPASSWORD)
'''        End If
'''    End If
    
    
    
    Dim MailMan As New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
    
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
    
    Dim Email As New ChilkatEmail2
    
    Email.AddTo sAddress, sAddress
        
    Email.FromAddress = LogAddress
    Email.fromName = LogName
    
    
'    If Not bCitrix Then
'        ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
'        Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
'        Dim myDB As Object '' LOTUS.NotesDatabase
'
'
'        On Error Resume Next
'        Set myNotes = GetObject(, "Notes.NotesSession")
'
'        If Err Then
'            Err.Clear
'            Set myNotes = CreateObject("Notes.NotesSession")
'            If Err Then
'                MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
'                GoTo GetOut
'            End If
'        End If
'        On Error GoTo 0
'        Set myDB = myNotes.GetDatabase("", "")
'        myDB.OPENMAIL
'        Set myDoc = myDB.CreateDocument
'
'    Else
'        ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
'        Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
'        Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
'
'
'        myDom.Initialize (sGAnnoPW)
'        'Set myDomDB = myDom.GetDatabase("detsrv1/det/GPJNotes", "mail\gannotat.nsf")
'        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'        Set myDoc = myDomDB.CreateDocument
'
'        Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'    End If
    
    Email.subject = MessHdr
    Email.Body = sMess
    
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
'    Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
'    Set myRichText = myDoc.CreateRichTextItem("Body")
'    myRichText.AppendText sMess ''' & vbNewLine & vbNewLine & vbNewLine & sLink
'    myDoc.AppendItemValue "SENDTO", sAddress
''''    myDoc.SaveMessageOnSend = True
'
'    Call myDoc.Send(False, sAddress)
'
'    If Err Then
'        MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                    vbExclamation, "Error Encountered"
'        Err = 0
'    End If
    
GetOut:
'    Set myReply = Nothing
'    Set myRichText = Nothing
'    Set myItem = Nothing
'    Set myDoc = Nothing
'
'    If bCitrix Then
'        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
'        If Not myDom Is Nothing Then Set myDom = Nothing
'    Else
'        If Not myDB Is Nothing Then Set myDB = Nothing
'        If Not myNotes Is Nothing Then Set myNotes = Nothing
'    End If

End Sub

'''''Public Function WritePermToOracle()
'''''    Dim i As Integer, iNode As Integer
'''''    Dim strUpdate As String, strSelect As String
'''''    Dim rst As ADODB.Recordset
'''''
'''''    For i = 1 To tvw1.Nodes.Count
'''''        If tvw1.Nodes(i).Checked = True Then
'''''            iNode = CInt(Mid(tvw1.Nodes(i).Key, 2))
'''''            strUpdate = "UPDATE " & IGLUserAR & " " & _
'''''                        "SET USER_PERMISSION_ID = USER_PERMISSION_ID + POWER(2, " & iNode & "), " & _
'''''                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
'''''                        "WHERE APP_ID = 1002 " & _
'''''                        "AND USER_SEQ_ID = " & lUserID
'''''            Conn.Execute (strUpdate)
'''''        End If
'''''    Next i
'''''
'''''    strSelect = "SELECT TO_CHAR(USER_PERMISSION_ID, '" & String(38, "9") & "') AS PERM " & _
'''''                "FROM " & IGLUserAR & " " & _
'''''                "WHERE APP_ID = 1002 " & _
'''''                "AND USER_SEQ_ID = " & lUserID
'''''    Set rst = Conn.Execute(strSelect)
'''''    If Not rst.EOF Then
'''''        sTotal = Trim(rst.FIELDS("PERM"))
'''''    Else
'''''        sTotal = "0"
'''''    End If
'''''    rst.Close: Set rst = Nothing
'''''End Function

Public Function WriteUserTypePermToOracle(lUID As Long)
    Dim i As Integer, iNode As Integer
    Dim strUpdate As String
    
    For i = 1 To tvw1.Nodes.Count
        If tvw1.Nodes(i).Checked = True Then
            iNode = CInt(Mid(tvw1.Nodes(i).Key, 2))
            strUpdate = "UPDATE " & ANOUserType & " " & _
                        "SET USERTYPEVALUE = USERTYPEVALUE + POWER(2, " & iNode & "), " & _
                        "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                        "WHERE USERTYPEID = " & lUID
            Conn.Execute (strUpdate)
        End If
    Next i
    
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
        
        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
                    "(SELECT COUNT(*) AS FP_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'DWG_MASTER') LF, " & _
                    "(SELECT COUNT(*) AS GFX_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'GFX_MASTER') LG"
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
        
        strSelect = "SELECT LF.FP_COUNT, LG.GFX_COUNT FROM " & _
                    "(SELECT COUNT(*) AS FP_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'DWG_MASTER') LF, " & _
                    "(SELECT COUNT(*) AS GFX_COUNT " & _
                    "FROM ANNOTATOR.ANO_LOCKLOG WHERE USER_SEQ_ID = " & rst.Fields("USER_SEQ_ID") & " " & _
                    "AND LOCKREFSOURCE = 'GFX_MASTER') LG"
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

Private Sub web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim strSelect As String, sList As String, sMess As String, sName As String
    Dim rst As ADODB.Recordset
    Dim iChk As Integer, iDash As Integer, iHtm As Integer
    Dim lUser As Long
    
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



