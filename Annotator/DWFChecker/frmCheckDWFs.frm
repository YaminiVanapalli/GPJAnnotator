VERSION 5.00
Begin VB.Form frmCheckDWFs 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Move 'em"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   300
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   7080
      ItemData        =   "frmCheckDWFs.frx":0000
      Left            =   3240
      List            =   "frmCheckDWFs.frx":0002
      TabIndex        =   4
      Top             =   900
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check 'em"
      Height          =   495
      Left            =   1860
      TabIndex        =   2
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get 'em"
      Height          =   495
      Left            =   540
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7080
      ItemData        =   "frmCheckDWFs.frx":0004
      Left            =   540
      List            =   "frmCheckDWFs.frx":0006
      TabIndex        =   0
      Top             =   900
      Width           =   2535
   End
   Begin VB.Label lblMoved 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4380
      TabIndex        =   7
      Top             =   600
      Width           =   105
   End
   Begin VB.Label lblCntBad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   8040
      Width           =   105
   End
   Begin VB.Label lblCnt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   8040
      Width           =   105
   End
End
Attribute VB_Name = "frmCheckDWFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New ADODB.Connection
Dim sPath As String


Private Sub Command1_Click()
    Dim sChk As String
    
    List1.Visible = False
    sChk = Dir(sPath & "*.dwf", vbNormal)
    Do While sChk <> ""
        List1.AddItem Left(sChk, Len(sChk) - 4)
'''        List1.Selected(List1.NewIndex) = True
'''        List1.Refresh
'''        lblCnt = List1.ListCount
'''        lblCnt.Refresh
        sChk = Dir
    Loop
    List1.Visible = True
    lblCnt = List1.ListCount
End Sub

Private Sub Command2_Click()
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Long
    
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
        List1.Refresh
        strSelect = "SELECT DWFID FROM ANNOTATOR.DWG_DWF " & _
                    "WHERE DWFID = " & List1.List(i)
        Set rst = Conn.Execute(strSelect)
        If rst.EOF Then
            List2.AddItem List1.List(i) & ".dwf"
            List2.Selected(List2.NewIndex) = True
            List2.Refresh
            lblCntBad = "Not Found:  " & List2.ListCount
            lblCntBad.Refresh
        End If
        rst.Close
    Next
    Set rst = Nothing
                    
End Sub

Private Sub Command3_Click()
    Dim i As Long, iCnt As Long
    Dim OldName, NewName
    
    iCnt = 0
    For i = 0 To List2.ListCount - 1
        OldName = sPath & List2.List(i)
        NewName = sPath & "MOVED\" & List2.List(i)
        Name OldName As NewName
        iCnt = iCnt + 1
        lblMoved = iCnt & " Moved"
        lblMoved.Refresh
    Next i
End Sub

Private Sub Form_Load()
    Dim ConnStr As String
    
    
    ConnStr = "DSN=JDE;UID=FPAPP;PWD=FPAPP"
    Conn.Open (ConnStr)
    sPath = "\\Detnovfs2\Data1\GPJAnnotator\Floorplans\"
    
        
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub
