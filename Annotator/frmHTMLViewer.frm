VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmHTMLViewer 
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHTMLViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBlank 
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   8599
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
   Begin VB.Menu mnuMenu 
      Caption         =   "Options"
      Begin VB.Menu mnuDownload 
         Caption         =   "Download..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSendALink 
         Caption         =   "Send-A-Link..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGFXData 
         Caption         =   "View File Data..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Document"
      End
   End
End
Attribute VB_Name = "frmHTMLViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iCount As Integer
Public pDownloadPath As String

Dim bDownloading As Boolean
Dim sPath As String

Dim pFile As String, pHdr As String, pFrom As String, pDownloadFile As String
Dim pSDID As Long, pGID As Long

Public Property Get PassFile() As String
    PassFile = pFile
End Property
Public Property Let PassFile(ByVal vNewValue As String)
    pFile = vNewValue
End Property

Public Property Get PassDFile() As String
    PassDFile = pDownloadFile
End Property
Public Property Let PassDFile(ByVal vNewValue As String)
    pDownloadFile = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = pHdr
End Property

Public Property Let PassHDR(ByVal vNewValue As String)
    pHdr = vNewValue
End Property

Public Property Get PassFrom() As String
    PassFrom = pFrom
End Property
Public Property Let PassFrom(ByVal vNewValue As String)
    pFrom = vNewValue
End Property

Public Property Get PassSDID() As Long
    PassSDID = pSDID
End Property
Public Property Let PassSDID(ByVal vNewValue As Long)
    pSDID = vNewValue
End Property

Public Property Get PassGID() As Long
    PassGID = pGID
End Property
Public Property Let PassGID(ByVal vNewValue As Long)
    pGID = vNewValue
End Property

Private Sub Form_Activate()
    On Error Resume Next
    If web1.Visible Then web1.SetFocus
End Sub

Private Sub Form_Load()
    web1.Top = 120: web1.Left = 120
''    wmp1.Top = 120: wmp1.Left = 120
''    xpdf1.Top = 120: xpdf1.Left = 120
    
'''    cmdBlank.Top = web1.Top
'''    cmdBlank.Left = web1.Left
    
    
'    If InStr(1, UCase(pFile), ".AVI") > 0 Or InStr(1, UCase(pFile), ".MPG") > 0 Then
'        wmp1.URL = pFile
'        wmp1.Visible = True
'        mnuDownload.Visible = True
'        mnuSendALink.Visible = True
'        mnuGFXData.Visible = True
'    Else
'        web1.Navigate2 pFile
'        web1.Visible = True
        
    
    
    
'    If InStr(1, UCase(pFile), ".AVI") = 0 And InStr(1, UCase(pFile), ".MPG") = 0 _
'                And InStr(1, UCase(pFile), ".PDF") = 0 Then
    If InStr(1, UCase(pFile), ".AVI") = 0 And InStr(1, UCase(pFile), ".MPG") = 0 Then
                ''And InStr(1, UCase(pFile), ".PDF") = 0 Then
        web1.Navigate2 pFile
        web1.Visible = True
'        If InStr(1, UCase(pFile), ".PDF") Then cmdBlank.Visible = True Else cmdBlank.Visible = False
''    ElseIf InStr(1, UCase(pFile), ".PDF") <> 0 Then
''        xpdf1.loadFile pFile
''        xpdf1.Visible = True
''    Else
''        wmp1.URL = pFile
''        wmp1.Visible = True
    End If
    
    Me.Caption = pHdr
    
    If pDownloadFile = "" Then mnuDownload.Visible = False Else mnuDownload.Visible = True
    
    If pFrom = "frmVersions" Or pFrom = "frmGraphics" Then
        Me.WindowState = 2
        mnuPrint.Visible = True
    
    ElseIf pFrom = "frmDownloadCart" Then
        mnuDownload.Visible = False
        mnuSendALink.Visible = False '' True
        If UCase(Left(Right(pDownloadFile, 3), 2)) = "PP" _
                    Or UCase(Left(Right(pDownloadFile, 3), 2)) = "AV" _
                    Or UCase(Left(Right(pDownloadFile, 3), 2)) = "MP" _
                    Or UCase(Left(Right(pDownloadFile, 3), 2)) = "MO" Then
            Me.WindowState = 0
            If UCase(Left(Right(pDownloadFile, 3), 2)) = "MO" Then
                Me.Height = 8550
                Me.Width = 9960
            End If
            mnuPrint.Visible = False
        Else
            Me.WindowState = 2
            mnuPrint.Visible = True
        End If
        
    ElseIf pFrom = "frmDIL" Then
        If pDownloadFile = "" Then
            Me.WindowState = 2
        Else
            If UCase(Left(Right(pDownloadFile, 3), 2)) = "PP" _
                        Or UCase(Left(Right(pDownloadFile, 3), 2)) = "AV" _
                        Or UCase(Left(Right(pDownloadFile, 3), 2)) = "MP" _
                        Or UCase(Left(Right(pDownloadFile, 3), 2)) = "MO" Then
                Me.WindowState = 0
                If UCase(Left(Right(pDownloadFile, 3), 2)) = "MO" Then
                    Me.Height = 8550
                    Me.Width = 9960
                End If
                mnuPrint.Visible = False
                mnuSendALink.Visible = True
            Else
                Me.WindowState = 2
                mnuPrint.Visible = True
            End If
        End If
        
    ElseIf pFrom = "frmComments" Then
        mnuPrint.Visible = True
    
    ElseIf pFrom = "frmEmailFile" Then
        mnuPrint.Visible = True
    
    End If
    
    
    If pGID = 0 Then Me.mnuGFXData.Visible = False Else mnuGFXData.Visible = True
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    web1.Width = Me.ScaleWidth - 240
    web1.Height = Me.ScaleHeight - 240
''    wmp1.Width = web1.Width
''    wmp1.Height = web1.Height
''    xpdf1.Width = web1.Width
''    xpdf1.Height = web1.Height
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuDownload_Click()
    Dim strSelect As String, sTemp As String, sFolder As String, sChk As String
    Dim rst As ADODB.Recordset
    Dim fs As New Scripting.FileSystemObject
    
'    Set fs = CreateObject("Scripting.FileSystemObject")
    
    pDownloadPath = ""
    frmBrowse.PassFrom = Me.Name
    frmBrowse.Show 1, Me
    
'''    If shlShell Is Nothing Then
'''        Set shlShell = New Shell32.Shell
'''    End If
'''
''''''    If sDownloadRootFolder = "" Then
'''        Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
'''                    "Select Folder to download Graphic into:", _
'''                    BIF_RETURNONLYFSDIRS)
''''''    Else
''''''        Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, _
''''''                    "Select Folder to download Graphic into:", _
''''''                    BIF_RETURNONLYFSDIRS, sDownloadRootFolder)
''''''    End If
    
'''    If shlFolder Is Nothing Then
    If pDownloadPath = "" Then
        Exit Sub
    Else
        Screen.MousePointer = 11
        On Error GoTo BadFile
'''        sFolder = shlFolder.Items.Item.Path
        sFolder = pDownloadPath '' shlFolder.Items.Item.Path
        
        If UCase(Left(sFolder, 1)) = "C" Then
            Screen.MousePointer = 0
            MsgBox "You do not have rights to download files onto one of " & _
                        "the Citrix Server drives." & vbNewLine & vbNewLine & _
                        "Please, select another location.", vbExclamation, "Invalid Location..."
            Exit Sub
        End If
        
        Err = 0
        On Error GoTo ErrorTrap
        
'''        sDownloadRootFolder = sFolder
        sPath = sFolder & "\" & pDownloadFile
        
'''        strSelect = "SELECT GPATH, GDESC, GFORMAT, AN8_CUNO " & _
'''                    "FROM " & GFXMas & " " & _
'''                    "WHERE GID = " & lGID
'''        Set rst = Conn.Execute(strSelect)
'''        If Not rst.EOF Then
'''            sPath = sFolder & "\" & pDownloadFile
'''        Else
'''            rst.Close: Set rst = Nothing
'''            Screen.MousePointer = 0
'''            MsgBox "Error:  File Not Found", vbExclamation, "File Not Copied..."
'''            Exit Sub
'''        End If
'        bDownloading = True
''        web1.Visible = False
'        web1.Navigate2 "about:Downloading Document [" & pDownloadFile & "]:  File will closed at completion."
        

        frmDownloadProgress.PassSRCFILE = pFile
        frmDownloadProgress.PassDESFILE = sPath
        frmDownloadProgress.Show 1, Me
        
'        FileCopy pFile, sPath
'''        rst.Close: Set rst = Nothing
        
        Screen.MousePointer = 0
'        MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
'        Unload Me
    End If
    
Exit Sub
ErrorTrap:
    Screen.MousePointer = 0
'''    rst.Close: Set rst = Nothing
    MsgBox "Error:  " & Err.Description, vbExclamation, "File Not Copied..."
Exit Sub
BadFile:
    Screen.MousePointer = 0
    MsgBox "The location you chose is not valid.  If you chose " & _
                "'DESKTOP', be aware the desktop in this interface is the " & _
                "Citrix Servers Desktop and you do not have rights to " & _
                "download files there." & vbNewLine & vbNewLine & _
                "Please, choose another folder.", vbExclamation, "Invalid Location..."
    Err.Clear
End Sub

Private Sub mnuGFXData_Click()
    Dim strSelect As String
    
    strSelect = "SELECT * " & _
                "FROM " & GFXMas & " " & _
                "WHERE GID = " & lGID
    Call GetGFXData(strSelect, "msgbox")
End Sub

Private Sub mnuPrint_Click()
    If web1.Visible Then
        On Error Resume Next
        web1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
''    ElseIf xpdf1.Visible Then
''        xpdf1.printWithDialog
    End If
'''    web1.SetFocus
'''''    web1.Refresh
'''    SendKeys "^p", True ''' & vbCr
'''    web1.SetFocus
''''    web1.Refresh
End Sub


'Private Sub web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'    If bDownloading Then
'        bDownloading = False
'        FileCopy pFile, sPath
'        MsgBox "File Copied to " & sPath, vbInformation, "File Download Successful..."
'        Unload Me
'    End If
'End Sub

Private Sub mnuSendALink_Click()

End Sub
