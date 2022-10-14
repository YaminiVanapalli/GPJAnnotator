VERSION 5.00
Begin VB.Form frmEmailEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5325
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      Begin VB.Label lblHdr 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Delete unwanted past Email Addresses, highlight those you wish to clear and click 'Delete Selections'."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   180
         TabIndex        =   6
         Top             =   150
         Width           =   4665
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   4200
      Width           =   1515
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selections"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.OptionButton optView 
      Caption         =   "Show my Email Address History for all Clients"
      Height          =   435
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton optView 
      Caption         =   "Show my Email Address History for current Client"
      Height          =   435
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.ListBox lstEmail 
      Height          =   2790
      ItemData        =   "frmEmailEdit.frx":0000
      Left            =   120
      List            =   "frmEmailEdit.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   1260
      Width           =   5055
   End
End
Attribute VB_Name = "frmEmailEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pUID As Long, pBCC As Long

Public Property Get PassUID() As Long
    PassUID = pUID
End Property
Public Property Let PassUID(ByVal vNewValue As Long)
    pUID = vNewValue
End Property

Public Property Get PassBCC() As Long
    PassBCC = pBCC
End Property
Public Property Let PassBCC(ByVal vNewValue As Long)
    pBCC = vNewValue
End Property


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    Dim strDelete As String
    
    Conn.BeginTrans
    For i = lstEmail.ListCount - 1 To 0 Step -1
        If lstEmail.Selected(i) Then
            If lstEmail.ItemData(i) > 0 Then
                strDelete = "DELETE FROM ANNOTATOR.ANO_EMAIL_ADDRESS " & _
                            "WHERE EMAIL_ID = " & lstEmail.ItemData(i)
                Conn.Execute (strDelete)
                lstEmail.RemoveItem (i)
            Else
                lstEmail.Selected(i) = False
            End If
        End If
    Next i
    Conn.CommitTrans
End Sub

Private Sub Form_Load()
    Me.Caption = FBCN
    optView(0).Value = True
'''    lblHdr.Caption = "To Delete unwanted past Email Addresses, highlight " & _
'''                "those you wish to clear and click 'Delete Selections'."
End Sub

Private Sub optView_Click(Index As Integer)
    Select Case Index
        Case 0: Call GetHistory(pUID, pBCC)
        Case 1: Call GetHistory(pUID, 0)
    End Select
End Sub

Public Sub GetHistory(tUID As Long, tBCC As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim lCUNO As Long
    
    lstEmail.Clear
    lCUNO = 0
    Select Case tBCC
        Case 0
            strSelect = "SELECT AE.EMAIL_ID, AE.EMAIL_ADDRESS, AE.AN8_CUNO, AB.ABALPH " & _
                        "FROM ANNOTATOR.ANO_EMAIL_ADDRESS AE, " & F0101 & " AB " & _
                        "WHERE AE.USER_SEQ_ID = " & tUID & " " & _
                        "AND AE.EMAIL_ADDRESS IS NOT NULL " & _
                        "AND AE.AN8_CUNO = AB.ABAN8 " & _
                        "ORDER BY UPPER(ABALPH), UPPER(EMAIL_ADDRESS)"
        Case Else
            strSelect = "SELECT AE.EMAIL_ID, AE.EMAIL_ADDRESS, AE.AN8_CUNO, AB.ABALPH " & _
                        "FROM ANNOTATOR.ANO_EMAIL_ADDRESS AE, " & F0101 & " AB " & _
                        "WHERE AE.USER_SEQ_ID = " & tUID & " " & _
                        "AND AE.EMAIL_ADDRESS IS NOT NULL " & _
                        "AND AE.AN8_CUNO = " & tBCC & " " & _
                        "AND AE.AN8_CUNO = AB.ABAN8 " & _
                        "ORDER BY UPPER(EMAIL_ADDRESS)"
    End Select
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        If lCUNO <> rst.Fields("AN8_CUNO") Then
            lCUNO = rst.Fields("AN8_CUNO")
            If lstEmail.ListCount > 0 Then lstEmail.AddItem "-----"
        End If
        lstEmail.AddItem Trim(rst.Fields("EMAIL_ADDRESS")) & " - " & _
                    "[" & Trim(rst.Fields("ABALPH")) & "]"
        lstEmail.ItemData(lstEmail.NewIndex) = rst.Fields("EMAIL_ID")
        
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
End Sub
