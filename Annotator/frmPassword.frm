VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3510
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
   ScaleHeight     =   3510
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2325
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2820
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel/Close"
      Height          =   492
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox txtPW2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   3012
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change password"
      Height          =   492
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox txtPW1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   3012
   End
   Begin VB.TextBox txtExistingPW 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   660
      Width           =   3012
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   6420
      TabIndex        =   6
      Text            =   "swesterh"
      Top             =   3960
      Width           =   3012
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   6420
      TabIndex        =   5
      Text            =   "GPJCO_TREE"
      Top             =   3540
      Width           =   3012
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verify NewPassword:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   225
      Index           =   2
      Left            =   6300
      TabIndex        =   10
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   225
      Index           =   1
      Left            =   6300
      TabIndex        =   9
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   225
      Index           =   0
      Left            =   6300
      TabIndex        =   8
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label lblMess 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   240
      Left            =   330
      TabIndex        =   7
      Top             =   480
      Width           =   4785
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'================================================================
'Copyright ® 2001 Novell, Inc.  All Rights Reserved.
'
'  With respect to this file, Novell hereby grants to Developer a
'  royalty-free, non-exclusive license to include this sample code
'  and derivative binaries in its product. Novell grants to Developer
'  worldwide distribution rights to market, distribute or sell this
'  sample code file and derivative binaries as a component of
'  Developer 's product(s).  Novell shall have no obligations to
'  Developer or Developer's customers with respect to this code.
'
'DISCLAIMER:
'
'  Novell disclaims and excludes any and all express, implied, and
'  statutory warranties, including, without limitation, warranties
'  of good title, warranties against infringement, and the implied
'  warranties of merchantibility and fitness for a particular purpose.
'  Novell does not warrant that the software will satisfy customer's
'  requirements or that the licensed works are without defect or error
'  or that the operation of the software will be uninterrupted.
'  Novell makes no warranties respecting any technical services or
'  support tools provided under the agreement, and disclaims all other
'  warranties, including the implied warranties of merchantability and
'  fitness for a particular purpose.
'
'================================================================
'
' Project:  VBPASSWD
'
'    Desc: Sample code which demonstrates how to use Novell DS DLL calls
'          in VB application for password verification and/or modification.
'
'          There is no need to be already authenticated to the NDS
'          for password verification/modification in order to run this sample code,
'          but sufficient access rights must be granted for changing the password.
'
'          To keep code as simple as possible only basic error handling is implemented.
'
'   Programmers:
'
'   Ini       Who                 Firm
'   ------------------------------------------------------------------
'   RLE       Rostislav Letos     Novell DeveloperNet Labs
'
'
'   History:
'
'   When              Who     What
'   ------------------------------------------------------------------
'   2001  January     RLE     Initial code
'=====================================================================

'Following declarations has been taken from Novell Libraries for Visual Basic
' available at http://developer.novell.com

Private Const DS_ROOT_NAME = "[Root]"
   
Private Const DCK_FLAGS = 1
Private Const DCK_CONFIDENCE = 2
Private Const DCK_NAME_CONTEXT = 3
Private Const DCK_TRANSPORT_TYPE = 4
Private Const DCK_REFERRAL_SCOPE = 5
Private Const DCK_LAST_CONNECTION = 8
Private Const DCK_LAST_SERVER_ADDRESS = 9
Private Const DCK_LAST_ADDRESS_USED = 10
Private Const DCK_TREE_NAME = 11
Private Const DCK_DSI_FLAGS = 12
Private Const DCK_NAME_FORM = 13

Private Const DCV_DEREF_ALIASES = &H1
Private Const DCV_XLATE_STRINGS = &H2
Private Const DCV_TYPELESS_NAMES = &H4
Private Const DCV_ASYNC_MODE = &H8
Private Const DCV_CANONICALIZE_NAMES = &H10
Private Const DCV_DEREF_BASE_CLASS = &H40
Private Const DCV_DISALLOW_REFERRALS = &H80
Private Const DCV_LOW_CONF = 0
Private Const DCV_MED_CONF = 1
Private Const DCV_HIGH_CONF = 2

Private Const NW_MAX_SERVER_NAME_LEN = 49
Private Const NW_MAX_TREE_NAME_LEN = 33

Private Const NDS_PASSWORD = 1        'Selects the NDS password
Private Const NT_PASSWORD = 2         'Selects the NT password in the NDS database
 
Private Declare Function NWCallsInit Lib "calwin32" _
    (reserved1 As Byte, reserved2 As Byte) As Long

Private Declare Function NWDSCreateContextHandle Lib "netwin32" _
    (context As Long) As Long

Private Declare Function NWDSGetContext Lib "netwin32" _
    (ByVal context As Long, ByVal key As Long, _
     ByVal value As Long) As Long
     
Private Declare Function NWDSSetContext Lib "netwin32" _
    (ByVal context As Long, ByVal key As Long, _
     ByVal value As Long) As Long

Private Declare Function NWDSFreeContext Lib "netwin32" _
    (ByVal context As Long) As Long

Private Declare Function NWDSVerifyObjectPassword Lib "netwin32" _
    (ByVal context As Long, ByVal optionsFlag As Long, _
     ByVal objectName As String, ByVal password As String) As Long

Private Declare Function NWDSChangeObjectPassword Lib "netwin32" _
    (ByVal context As Long, ByVal optionsFlag As Long, _
     ByVal objectName As String, ByVal oldPassword As String, _
     ByVal newPassword As String) As Long

Private Declare Function NWDSGenerateObjectKeyPair Lib "netwin32" _
    (ByVal contextHandle As Long, ByVal objectName As String, _
     ByVal objectPassword As String, ByVal optionsFlag As Long) As Long

' It`s a good idea to have the following option switched ON
'   to avoid unexpectable results with VB Variant types
'   whenever we need to call DLL functions
Option Explicit

Private Sub ByteArrayToString(src() As Byte, dest As String)
Dim i As Integer
    i = 0
    dest = ""
    While src(i) <> 0
        dest = dest + Chr(src(i))
        i = i + 1
    Wend
End Sub

Private Sub StringToByteArray(src As String, dest() As Byte)
Dim i As Integer
' Following For-Next loop  should run to 0x0 char only
'   but we do not care if it runs longer
    For i = 0 To Len(src) - 1
        dest(i) = CByte(Asc(Mid(src, i + 1, 1)))
    Next i
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

'''Private Sub Command1_Click()
'''Dim retCode As Long, context As Long
'''Dim anyName As String
'''Dim byteName(255) As Byte
'''
'''    If Text1.Text = "" Or Text2.Text = "" Or txtexistingpw.Text = "" Then
'''        Exit Sub
'''    End If
'''
'''    Screen.MousePointer = vbHourglass
'''    retCode = NWDSCreateContextHandle(context)
'''
'''' Here we set one of the context variables (name context) to the [Root]
'''' This will require fully distinguished userID to be passed as a parameter
'''    anyName = DS_ROOT_NAME + Chr(0)
'''    Call StringToByteArray(anyName, byteName)
'''    retCode = NWDSSetContext(context, DCK_NAME_CONTEXT, VarPtr(byteName(0)))
'''
'''' The good idea is to set explicitly DS tree name in case you are in
'''' multiple-trees environment
'''    anyName = Text1.Text + Chr(0)
'''    Call StringToByteArray(anyName, byteName)
'''    retCode = NWDSSetContext(context, DCK_TREE_NAME, VarPtr(byteName(0)))
'''
'''' Now we can verify userID/password pair...
'''    retCode = NWDSVerifyObjectPassword(context, 0, Text2.Text + Chr(0), txtexistingpw.Text + Chr(0))
'''    If retCode = 0 Then
'''        MsgBox "UserID/password pair is correct !", vbInformation
'''    Else
'''        MsgBox "Password verification FAILED ! Error=" + Str(retCode), vbCritical
'''    End If
'''
'''' Some clean-ups
'''    retCode = NWDSFreeContext(context)
'''
'''    Screen.MousePointer = vbDefault
'''
'''End Sub

Private Sub cmdChange_Click()
    Dim retCode As Long, context As Long
    Dim anyName As String
    Dim byteName(255) As Byte
    Dim strUpdate As String
    
    
    If Text1.Text = "" Or Text2.Text = "" Or txtPW1.Text = "" Or txtPW2.Text = "" Then
        MsgBox "Both the 'New Password' and the 'Verify New Password' boxes must be filled in!", _
                    vbExclamation, "Incomplete..."
        Exit Sub
    End If
    
    If txtPW1.Text <> txtPW2.Text Then
        MsgBox "New Password strings are not identical!", _
                    vbExclamation, "Failure..."
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    retCode = NWDSCreateContextHandle(context)
    
' Here we set one of the context variables (name context) to the [Root]
' This will require fully distinguished userID to be passed as a parameter
    anyName = DS_ROOT_NAME + Chr(0)
    Call StringToByteArray(anyName, byteName)
    retCode = NWDSSetContext(context, DCK_NAME_CONTEXT, VarPtr(byteName(0)))
    
' The good idea is to set explicitly DS tree name in case you are in
' multiple-trees environment
    anyName = Text1.Text + Chr(0)
    Call StringToByteArray(anyName, byteName)
    retCode = NWDSSetContext(context, DCK_TREE_NAME, VarPtr(byteName(0)))
   
' Now we change user's password
    If txtExistingPW.Text = "" Then
        retCode = NWDSGenerateObjectKeyPair(context, Text2.Text + Chr(0), txtPW1.Text + Chr(0), NDS_PASSWORD)
'Following line can also be used in order to synchronize new password in NDS4NT environment
'        retCode = NWDSGenerateObjectKeyPair(context, Text2.Text + Chr(0), txtpw1.Text + Chr(0), NT_PASSWORD)
    Else
        retCode = NWDSChangeObjectPassword(context, NDS_PASSWORD, Text2.Text + Chr(0), txtExistingPW.Text + Chr(0), txtPW1.Text + Chr(0))
'Following line can also be used in order to synchronize new password in NDS4NT environment
'        retCode = NWDSChangeObjectPassword(context, NT_PASSWORD, Text2.Text + Chr(0), txtexistingpw.Text + Chr(0), txtpw1.Text + Chr(0))
    End If
    
    If retCode = 0 Then
'''        Dim oDomain As IADsDomain ''' As Object
        Dim oUser As IADsUser ''' As Object
'''        Set oDomain = GetObject("WinNT://63.79.167.166")
        Set oUser = GetObject("WinNT://DETCTX02/" & Shortname)
        oUser.SetPassword txtPW1.Text
        oUser.SetInfo
        Set oUser = Nothing
'''        Set oDomain = Nothing
        
        MsgBox "Password has been changed.", vbInformation, "Successful..."
        
    Else
        Select Case retCode
            Case 669
                MsgBox "The 'Existing Password' entered does not match the current Password.", _
                            vbExclamation, "Failure..."
            Case 216
                MsgBox "The Password change Failed!  Passwords must be a minimum of 8 characters.", _
                            vbExclamation, "Failure..."
            Case 672
                MsgBox "Unable to change Password, based on 'Existing Password' provided!", _
                            vbExclamation, "Failure..."
            Case Else
                MsgBox "Password change FAILED!" & vbNewLine & "Error=" + Str(retCode), _
                            vbExclamation, "Failure..."
        End Select
        ' Some clean-ups
        retCode = NWDSFreeContext(context)
        Screen.MousePointer = vbDefault
        Exit Sub
        
    End If
    
' Some clean-ups
    retCode = NWDSFreeContext(context)
    
    
    ''NOW UPDATE PASSWORD IN ORACLE''
    strUpdate = "UPDATE " & IGLUserAR & " " & _
                "SET PCODE = '" & txtPW1.Text & "', " & _
                "UPDUSER = '" & DeGlitch(Left(LogName, 20)) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE USER_SEQ_ID = " & UserID & " " & _
                "AND APP_ID = 1002"
    Conn.Execute (strUpdate)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim result As Long
    Dim i As Integer

' All API calls need following initialization,
'  otherwise they will return errors !
    result = NWCallsInit(0, 0)
    If result <> 0 Then
        MsgBox "Unable to initialize !"
        Unload Me
    End If
    
'''    bGPJ = False   '''FOR TESTING ONLY
    
'    Me.Top = frmStartUp.Top + (frmStartUp.Height - frmStartUp.ScaleHeight) + _
'                (frmStartUp.imgBadge.Top - 180)
'    Me.Left = frmStartUp.Left + ((frmStartUp.Width - frmStartUp.ScaleWidth) / 2) + _
'                ((frmStartUp.ScaleWidth - Me.Width) / 2)
    
    txtExistingPW.BackColor = lColor
    txtPW1.BackColor = lColor
    txtPW2.BackColor = lColor
    lblMess.ForeColor = lColor
    For i = 0 To lbl1.Count - 1
        lbl1(i).ForeColor = lColor
    Next i
    cmdChange.BackColor = lColor
    cmdClose.BackColor = lColor
    cmdOK.BackColor = lColor
    
    Select Case bGPJ
        Case True
            lblMess.Caption = "The Annotator provides means for non-GPJ Users, only, " & _
                        "to reset their passwords." & vbNewLine & vbNewLine & _
                        "For internal GPJ users, since your Annotator password and " & _
                        "your Novell password are the same, you will need to change " & _
                        "your password through Novell.  Once changed, the new password " & _
                        "will immediately be available for Annotator Login."
                        
            cmdOK.Visible = True
            Me.Width = 5535
        Case False
            lblMess.Caption = "To change your password:  First enter your existing password, " & _
                        "followed by your new password, which must be re-entered in the bottom box. " & _
                        vbNewLine & vbNewLine & _
                        "Passwords must be a minimum of 8 characters (not to exceed 16), " & _
                        "and are NOT case-sensitive." & _
                        vbNewLine & vbNewLine & _
                        "Be aware, this process may take up to two minutes, as the new password " & _
                        "is passed on to the Server."
            cmdOK.Visible = False
            Me.Width = 10515
    End Select
            
    Text2.Text = Shortname & ".External"
End Sub
