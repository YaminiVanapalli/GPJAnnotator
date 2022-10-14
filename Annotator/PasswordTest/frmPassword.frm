VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3525
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
   ScaleHeight     =   3525
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel/Close"
      Height          =   492
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   8
      Top             =   1920
      Width           =   3012
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change password"
      Height          =   492
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   6
      Top             =   1320
      Width           =   3012
   End
   Begin VB.TextBox Text3 
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
      Top             =   660
      Width           =   3012
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   9600
      TabIndex        =   1
      Text            =   "rletos.dev"
      Top             =   2760
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.TextBox frmPassword 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Text            =   "GPJCO_TREE"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verify new password:"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6360
      TabIndex        =   9
      Top             =   1680
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New password:"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6360
      TabIndex        =   5
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existing password:"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6360
      TabIndex        =   4
      Top             =   420
      Width           =   1350
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   4905
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

'''Private Sub Command1_Click()
'''Dim retCode As Long, context As Long
'''Dim anyName As String
'''Dim byteName(255) As Byte
'''
'''    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
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
'''    retCode = NWDSVerifyObjectPassword(context, 0, Text2.Text + Chr(0), Text3.Text + Chr(0))
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

Private Sub Command2_Click()
Dim retCode As Long, context As Long
Dim anyName As String
Dim byteName(255) As Byte

    If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
        Exit Sub
    End If
    
    If Text4.Text <> Text5.Text Then
        MsgBox "Entered new password strings are not identical !", vbCritical
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
    If Text3.Text = "" Then
        retCode = NWDSGenerateObjectKeyPair(context, Text2.Text + Chr(0), Text4.Text + Chr(0), NDS_PASSWORD)
'Following line can also be used in order to synchronize new password in NDS4NT environment
'        retCode = NWDSGenerateObjectKeyPair(context, Text2.Text + Chr(0), Text4.Text + Chr(0), NT_PASSWORD)
    Else
        retCode = NWDSChangeObjectPassword(context, NDS_PASSWORD, Text2.Text + Chr(0), Text3.Text + Chr(0), Text4.Text + Chr(0))
'Following line can also be used in order to synchronize new password in NDS4NT environment
'        retCode = NWDSChangeObjectPassword(context, NT_PASSWORD, Text2.Text + Chr(0), Text3.Text + Chr(0), Text4.Text + Chr(0))
    End If
    If retCode = 0 Then
        MsgBox "Password has been changed !", vbInformation
    Else
        MsgBox "Password change FAILED ! Error=" + Str(retCode), vbCritical
    End If
    
' Some clean-ups
    retCode = NWDSFreeContext(context)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim result As Long
' All API calls need following initialization,
'  otherwise they will return errors !
    result = NWCallsInit(0, 0)
    If result <> 0 Then
        MsgBox "Unable to initialize !"
        Unload Me
    End If
    
    Label1.Caption = "The Annotator provides you a means to reset your password.  " & _
                "Be aware, for internal GPJ users, since your Annotator password and " & _
                "your Novell password are the same, this will reset your Novell password, " & _
                "but it will NOT reset the password on your computer." & vbNewLine & vbNewLine & _
                "GPJ Users:  The next time you login to your computer, you will be prompted that " & _
                "the two passwords are out of synch, and force you to update it."
End Sub
