VERSION 5.00
Begin VB.Form frmPrinterDrivers 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   30
   ClientTop       =   0
   ClientWidth     =   10425
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
   ScaleHeight     =   3975
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      IMEMode         =   3  'DISABLE
      Left            =   7920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2400
      Width           =   2250
   End
   Begin VB.CheckBox chkIssue 
      BackColor       =   &H00000000&
      Caption         =   "Required Printers are missing."
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   8
      Top             =   2820
      Width           =   1995
   End
   Begin VB.CheckBox chkIssue 
      BackColor       =   &H00000000&
      Caption         =   "Poor print quality."
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   5940
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox chkIssue 
      BackColor       =   &H00000000&
      Caption         =   "Printing is disabled."
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   5940
      TabIndex        =   6
      Top             =   2220
      Width           =   1935
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Enabled         =   0   'False
      Height          =   435
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1635
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel/Close"
      Height          =   435
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1635
   End
   Begin VB.TextBox txtDrivers 
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
      IMEMode         =   3  'DISABLE
      Left            =   5940
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1260
      Width           =   4230
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment (optional):"
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   10
      Top             =   2160
      Width           =   1470
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter requested Driver names (Case-Sensitive):"
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   1
      Left            =   5940
      TabIndex        =   5
      Top             =   1020
      Width           =   3450
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000F&
      Height          =   195
      Index           =   0
      Left            =   5940
      TabIndex        =   3
      Top             =   180
      Width           =   4230
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMess 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   720
      Width           =   5475
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPrinterDrivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub cmdSubmit_Click()
    Dim sDriver As String, sIssues As String, sComment As String
    Dim i As Integer
    
    sDriver = txtDrivers.Text & vbNewLine
    
    sIssues = ""
    For i = 0 To chkIssue.Count - 1
        If chkIssue(i).value = 1 Then
            sIssues = sIssues & UCase(chkIssue(i).Caption) & vbNewLine
        End If
    Next i
    
    If txtComment.Text <> "" Then
        sIssues = sIssues & vbNewLine & "User Comment:" & vbNewLine & _
                    Trim(txtComment.Text) & vbNewLine
    End If
    
    ''/// DRIVER EMAIL ALERT TO HELP DESK \\\''
    Call DriverAlert(sDriver, sIssues, LogName, LogAddress, Now)
    
    
End Sub

Private Sub Form_Load()
    Dim result As Long
    Dim i As Integer

    
'    Me.Top = frmStartUp.Top + (frmStartUp.Height - frmStartUp.ScaleHeight) + _
'                (frmStartUp.imgBadge.Top - 360)
'    Me.Left = frmStartUp.Left + ((frmStartUp.Width - frmStartUp.ScaleWidth) / 2) + _
'                ((frmStartUp.ScaleWidth - Me.Width) / 2)
    
'''    txtDrivers.BackColor = lColor
    lblMess.ForeColor = lColor
    For i = 0 To lbl1.Count - 1
        lbl1(i).ForeColor = lColor
    Next i
    For i = 0 To chkIssue.Count - 1
        chkIssue(i).ForeColor = lColor
    Next i
    
'''    cmdChange.BackColor = lColor
'''    cmdClose.BackColor = lColor
'''    cmdOK.BackColor = lColor
    
    lblMess.Caption = "In order to Print, the Annotator's Application Server " & _
                "must have the drivers for your Printer installed." & vbNewLine & vbNewLine & _
                "If your printing has been disabled, due to the absence of an " & _
                "appropriate driver, or if you are experiencing poor print quality, " & _
                "enter the names of your missing drivers in the box at right, " & _
                "check the issue(s) that applies, and 'Submit'.  We will then install " & _
                "the requested drivers, and alert you of their availability." & _
                vbNewLine & vbNewLine & _
                "NOTE: The Annotator should make all of your configured Printers " & _
                "available.  If any required Printers are not displaying, you should " & _
                "submit those drivers, as well."
    lbl1(0).Caption = "In Windows, the driver name can be found by opening the 'Printers' " & _
                "menu, right-clicking the Printer you want accessible in the Annotator, " & _
                "and selecting 'Properties'."
    
End Sub

Private Sub txtDrivers_Change()
    If txtDrivers.Text <> "" Then cmdSubmit.Enabled = True Else cmdSubmit.Enabled = False
End Sub
