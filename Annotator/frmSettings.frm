VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   45
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   2160
      TabIndex        =   6
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   435
      Left            =   720
      TabIndex        =   5
      Top             =   1980
      Width           =   1275
   End
   Begin VB.PictureBox picSettings 
      BackColor       =   &H00666666&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   60
      ScaleHeight     =   1740
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   60
      Width           =   4020
      Begin VB.Frame Frame1 
         BackColor       =   &H00666666&
         Caption         =   "Exit Settings..."
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   3825
         Begin VB.OptionButton optClientSave 
            BackColor       =   &H00666666&
            Caption         =   "Do not Save a default Client"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   4
            Top             =   1140
            Width           =   3735
         End
         Begin VB.OptionButton optClientSave 
            BackColor       =   &H00666666&
            Caption         =   "Retain 'SAW' as your default Client"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Index           =   1
            Left            =   60
            TabIndex        =   3
            Top             =   660
            Width           =   3735
         End
         Begin VB.OptionButton optClientSave 
            BackColor       =   &H00666666&
            Caption         =   "Set 'SAW' as your default Client"
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Width           =   3735
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iOpt As Integer

Dim pBCC As Long, pBCC_Def As Long
Dim pFBCN As String, pFBCN_Def As String
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

Public Property Get PassBCC_DEF() As Long
    PassBCC_DEF = pBCC_Def
End Property
Public Property Let PassBCC_DEF(ByVal vNewValue As Long)
    pBCC_Def = vNewValue
End Property

Public Property Get PassFBCN_DEF() As String
    PassFBCN_DEF = pFBCN_Def
End Property

Public Property Let PassFBCN_DEF(ByVal vNewValue As String)
    pFBCN_Def = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property

Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property




Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strDelete As String, strInsert As String
    
    Select Case iOpt
        Case -1
            MsgBox "No option has been selected", vbExclamation, "Hey..."
            Exit Sub
        Case 0
            strDelete = "DELETE FROM " & ANOUPref & " " & _
                        "WHERE APP_ID = 1002 " & _
                        "AND USER_SEQ_ID = " & UserID
            Conn.Execute (strDelete)
            strInsert = "INSERT INTO " & ANOUPref & " " & _
                        "(USER_SEQ_ID, APP_ID, AN8_CUNO, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & UserID & ", 1002, " & pBCC & ", " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            defCUNO = pBCC
            defFBCN = pFBCN
        Case 1
            strDelete = "DELETE FROM " & ANOUPref & " " & _
                        "WHERE APP_ID = 1002 " & _
                        "AND USER_SEQ_ID = " & UserID
            Conn.Execute (strDelete)
            strInsert = "INSERT INTO " & ANOUPref & " " & _
                        "(USER_SEQ_ID, APP_ID, AN8_CUNO, " & _
                        "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
                        "VALUES " & _
                        "(" & UserID & ", 1002, " & pBCC_Def & ", " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
                        "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
            Conn.Execute (strInsert)
            defCUNO = pBCC_Def
            defFBCN = pFBCN_Def
        Case 2
            strDelete = "DELETE FROM " & ANOUPref & " " & _
                        "WHERE APP_ID = 1002 " & _
                        "AND USER_SEQ_ID = " & UserID
            Conn.Execute (strDelete)
            defCUNO = 0
            defFBCN = ""
    End Select
    
    Unload Me
End Sub

Private Sub Form_Load()
    iOpt = -1
    Select Case pFrom
        Case "GH"
            With frmGraphics
                Me.Left = .Left + ((.Width - .ScaleWidth) / 2)
                
                Me.Top = .Top + ((.Height - .ScaleHeight) - ((.Width - .ScaleWidth) / 2)) + _
                            .picMenu2.Top + .imgFullSize.Top + .imgFullSize.Height
            End With
        Case "FP"
            With frmAnnotator
                Me.Left = .Left + ((.Width - .ScaleWidth) / 2) + .lblSettings.Left + _
                            .lblSettings.Width - Me.Width
                Me.Top = .Top + ((.Height - .ScaleHeight) - ((.Width - .ScaleWidth) / 2)) + _
                            .lblSettings.Top + .lblSettings.Height
            End With
    End Select
    If pBCC_Def = 0 Then
        optClientSave(1).Enabled = False
        optClientSave(2).Enabled = False
    End If
    If pBCC = 0 Then
        optClientSave(0).Enabled = False
    End If
    
    optClientSave(0).Caption = Replace(optClientSave(0).Caption, "SAW", UCase(pFBCN))
    optClientSave(1).Caption = Replace(optClientSave(1).Caption, "SAW", UCase(pFBCN_Def))
End Sub

Private Sub optClientSave_Click(Index As Integer)
    iOpt = Index
    cmdSave.Enabled = True
End Sub
