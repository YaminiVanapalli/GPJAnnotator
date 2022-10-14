Attribute VB_Name = "modAnnotator"
Option Explicit

'///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Const NAB_NAME = "names.nsf"    'Domino Directory name
Public Const LKUP_SRVR = "GPJCO/GPJNotes"      'main Domino server

Public bENABLE_PRINTERS As Boolean, bPRINTER_ALERT As Boolean
'\\\\\------------------------------------/////

Public bDebug As Boolean, bHideRN As Boolean

Public strMailSrvr As String, strMailFile As String

'''Public sDownloadRootFolder As String
Public defCUNO As Long
Public defFBCN As String
Public defSIN As String, defEditSIN As String

Public sNOTESID As String, sNOTESPASSWORD As String

Public bGfxMandRecip As Boolean

Public iAppConn As Integer
Public bCitrix As Boolean
Public sTempPath As String, sGIPath As String

Public ConnStr As String
Public CunoList As String
Public BCC As String
Public SHCD As Long
Public SHYR As Integer
Public FBCN As String
Public BCN As String
Public SHNM As String
Public sPDFFile As String
Public iPDFPage As Integer
Public Shortname As String, LogName As String, _
            LogAddress As String, LogFirstName As String, NotesServer As String, _
            UserType As String
Public sOUser As String
Public UserID As Long
Public bBos As Boolean, bGPJNotes As Boolean, bGPJ As Boolean  ''''', bTeam As Boolean
Public lGID As Long
Public sCGDesc As String
Public bAnnoOpen As Boolean, bGfxOpen As Boolean, bConstOpen As Boolean, _
            bCommentsOpen As Boolean
Public AppWindowState As Integer
Public ConnOpen As Boolean
Public Conn As Object '''ADODB.Connection
Public iCurrGType As Integer
Public iCurrSHYR As Integer
Public CurrSelect(0 To 3) As String
Public bPerm() As Boolean
Public STZ As String
Public bClientAll_Enabled As Boolean, bTeamMember As Boolean, bGFXReviewer As Boolean, _
            bDo_Printer_Check As Boolean, bICAUser As Boolean
Public strCunoList As String
Public strHTMLPath As String, sSupDocPath As String
Public lColor As Long '', lColor2 As Long
Public lGeo_Fore As Long, lGeo_Back As Long, lGeo_Bright As Long, lGeo_Dark As Long
Public sLink As String, sGLLink As String, sLink_Disclaimer As String
Public sPrinter As String
Public sPassInValue As String
Public bPassIn As Boolean '''''''''', bWatching As Boolean
Public iPassIn As Integer ''1=floorplan; 2=graphic''
Public lOpenID As Long
Public sLinkID As String
Public PassDate As Date
Public lOpenInViewer As Long

Public StatusGID As Long
Dim GFXAddress() As String, GFXMandRecip() As String, GFXMandRecipFull() As String
Public sCC As String

Public GfxType(1 To 4) As String
Public imageX As Single, imageY As Single, spaceX As Single, spaceY As Single
Public iRes As Integer, iRows As Integer, iCols As Integer, iView As Integer

Public sVPath As String, sBMPPath As String

Public sIFile As String, sSFile As String

Public sDWGZip As String

Public lIconX As Long, lIconY As Long
Public iIconSize As Integer
Public Const IconSize_0x = 1200
Public Const IconSize_0y = 900
Public Const IconSize_1x = 1600
Public Const IconSize_1y = 1200
Public Const IconSize_2x = 2800
Public Const IconSize_2y = 2100

Public Const sGAnnoPW As String = "XPD8Notes"

Public Const DWGShow = "ANNOTATOR.DWG_SHOW"
Public Const DWGElt = "ANNOTATOR.DWG_ELEMENT"
Public Const DWGMas = "ANNOTATOR.DWG_MASTER"
Public Const DWGSht = "ANNOTATOR.DWG_SHEET"
Public Const DWGDwf = "ANNOTATOR.DWG_DWF"
Public Const DWGSeq = "ANNOTATOR.DWG_SEQ"

Public Const ANOETeam = "ANNOTATOR.ANO_EMAIL_TEAM"
Public Const ANOETeamUR = "ANNOTATOR.ANO_EMAIL_TEAM_USER_R"
Public Const ANOLockLog = "ANNOTATOR.ANO_LOCKLOG"
Public Const ANOComment = "ANNOTATOR.ANO_COMMENT"
Public Const ANOUserType = "ANNOTATOR.ANO_USERTYPE"
Public Const ANOPerm = "ANNOTATOR.ANO_PERM"
Public Const ANOHelp = "ANNOTATOR.ANO_HELPFILE"
Public Const ANOSeq = "ANNOTATOR.ANO_SEQ"
Public Const ANOSuppAll = "ANNOTATORVIEW_SUPPLIER_ALL"
Public Const ANOSuppAct = "ANNOTATORVIEW_SUPPLIER_ACTIVE"
Public Const ANOSession = "ANNOTATOR.ANO_SESSION"
Public Const ANOUPref = "ANNOTATOR.ANO_USER_PREFERENCE"

'///// ADDED 06-SEP-2001 FOR PRINTER RECOGNITION CHANGES \\\\\
Public Const ANODriver = "ANNOTATOR.ANO_PRINTER_DRIVER"
'\\\\\---------------------------------------------------------/////

'''
'''BAAN.TTIPCS905615
Public Const F5601 = "PRODDTA.F5601"
'''BAAN.TTIPCS907615
Public Const F5611 = "PRODDTA.F5611"
Public Const F5620 = "PRODDTA.F5620"
'''BAAN.TTIPCS020615
Public Const F0006 = "PRODDTA.F0006"
Public Const F0005 = "PRODCTL.F0005"
'''BAAN.TTCCOM010615
Public Const F0101 = "PRODDTA.F0101"
'''BAAN.TTCCOM001615
Public Const F060116_View = "PRODDTA.F060116_view"
'''Public Const F060116 = "PRODDTA.F060116" ''' replaced with F060116_View NET-410
'''BAAN.TTCCOM020615
'''BAAN.TTIIGL001615
Public Const F0115 = "PRODDTA.F0115"
Public Const F0116 = "PRODDTA.F0116"
Public Const F0401 = "PRODDTA.F0401"
'''
'''BAAN.TTIIGL030615
Public Const IGLKit = "IGLPROD.IGL_KIT"
'''BAAN.TTIIGL110615
Public Const IGLElt = "IGLPROD.IGL_ELEMENT"
'''BAAN.TTIIGL120615
Public Const IGLEltX = "IGLPROD.IGL_ELEMENT_EXTENSION"
'''BAAN.TTIIGL130615
Public Const IGLPart = "IGLPROD.IGL_PART"
'''BAAN.TTIIGL040615
Public Const AQUAKitU = "AQUA.AQUA_KIT_USE"
'''BAAN.TTIIGL210615
Public Const AQUAEltU = "AQUA.AQUA_ELEMENT_USE"
'''BAAN.TTIIGL230615
Public Const AQUAPartU = "AQUA.AQUA_PART_USE"
'''BAAN.TTIIGL007615

''Public Const IGLEmpDay = "IGL_EMPLOYEE_DAY"
''CHANGE TO THIS WHEN MOVING TO SS IN APM''
Public Const IGLEmpDay = "SFDC.SFDCVIEW_EMPLOYEE_DAY"

'''BAAN.TTIIGL005615
''Public Const IGLEmpTask = "IGL_EMPLOYEE_TASK"
Public Const IGLEmpTask = "SFDC.SFDCVIEW_EMPLOYEE_TASK"

'''BAAN.TTIIGL000615
Public Const IGLRef = "IGLPROD.IGL_REF"
'''
Public Const IGLSeq = "IGLPROD.IGL_SEQ"

Public Const GFXSeq = "ANNOTATOR.GFX_SEQ"
Public Const GFXShow = "ANNOTATOR.GFX_SHOW"
Public Const GFXMas = "ANNOTATOR.GFX_MASTER"
Public Const GFXElt = "ANNOTATOR.GFX_ELEMENT"
Public Const GFXRed = "ANNOTATOR.GFX_REDLINE"
Public Const GFXFolder = "ANNOTATOR.GFX_FOLDER"

Public Const IGLUser = "IGLPROD.IGL_USER"
Public Const IGLCGR = "IGLPROD.IGL_CUNO_GROUP_R"
Public Const IGLCGMas = "IGLPROD.IGL_CUNO_GROUP_MASTER"
Public Const IGLUserAR = "IGLPROD.IGL_USER_APP_R"
Public Const IGLUserCR = "IGLPROD.IGL_USER_CUNO_R"

Public Const SRAHallMas = "IGLPROD.SRA_HALLMASTER"
Public Const SRACliHall = "IGLPROD.SRA_CLIENTHALL"
Public Const SRAEase = "IGLPROD.SRA_EASEMENT"
Public Const SRAHallRes = "IGLPROD.SRA_SHOWHALLRESTRICTION"
Public Const SRAEngCodeReq = "IGLPROD.SRA_ENGCODEREQUIREMENT"
Public Const SRAEngCodeConf = "IGLPROD.SRA_ENGCODECONFIRMATION"
Public Const SRASeq = "IGLPROD.SRA_SEQ"

'''BAAN.TTCCOM020615
'''BAAN.TTIIGL005615

Public shlShell As Shell32.Shell
Public shlFolder As Shell32.Folder
Public Const BIF_RETURNONLYFSDIRS = &H1

Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
        Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)

Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory _
        As String, ByVal nShowCmd As Long) As Long
     
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
        "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
        
Declare Function SetEnvironmentVariable& Lib "kernel32" Alias _
        "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String)
        
Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Declare Function GetFileSize& Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim p As KBDLLHOOKSTRUCT

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim fEatKeystroke As Boolean
   
   If (nCode = HC_ACTION) Then
      If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
         CopyMemory p, ByVal lParam, Len(p)
         fEatKeystroke = _
            p.vkCode = 119
'''            p.vkCode = VK_LWIN Or _
'''            p.vkCode = VK_RWIN Or _
'''            p.vkCode = VK_APPS Or _
'''            p.vkCode = VK_CONTROL Or _
'''            p.vkCode = VK_SHIFT Or _
'''            p.vkCode = VK_MENU Or _
'''            ((GetKeyState(VK_CONTROL) And &H8000) <> 0) Or _
'''            ((p.flags And LLKHF_ALTDOWN) <> 0)
        End If
    End If
    
    If fEatKeystroke Then
        LowLevelKeyboardProc = -1
    Else
        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If
End Function


Public Function DblAmp(txt As String) As String
    Dim Pos As Integer
    Pos = 1
    Do While Pos <> 0
        Pos = InStr(Pos, txt, "&")
        If Pos <> 0 Then
            txt = Left(txt, Pos - 1) & Chr(38) & Mid(txt, Pos)
            Pos = Pos + 2
        End If
    Loop
    DblAmp = txt
End Function

Public Function GetBCN(tmpBCC As String)
    Dim rstCN As ADODB.Recordset
    Dim strSelect As String
    
    '****'Conn.Open ALREADY****
    strSelect = "SELECT ABALPH FROM " & F0101 & " " & _
                "WHERE ABAN8 = " & CLng(tmpBCC)
    Set rstCN = Conn.Execute(strSelect)
    If Not rstCN.EOF Then
        GetBCN = UCase(Trim(rstCN.Fields("ABALPH")))
    Else
        GetBCN = ""
    End If
    rstCN.Close
    Set rstCN = Nothing
End Function

Public Function GetSHNM(tSHCD As Long, tSHYR As Integer)
    Dim rstSN As ADODB.Recordset
    Dim strSelect As String
    
    '****'Conn.Open ALREADY****
    strSelect = "SELECT SHY56NAMA FROM " & F5601 & " " & _
                "WHERE SHY56SHCD = " & tSHCD & " " & _
                "AND SHY56SHYR = " & tSHYR
    Set rstSN = Conn.Execute(strSelect)
    If Not rstSN.EOF Then
        SHNM = UCase(Trim(rstSN.Fields("SHY56NAMA")))
    Else
        SHNM = ""
    End If
    rstSN.Close
    Set rstSN = Nothing
End Function

Public Sub GetShowClients(combo As ComboBox, iSHYR As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT CS.CSY56CUNO, C.ABALPH " & _
                    "FROM " & F5611 & " CS, " & F0101 & " C " & _
                    "WHERE CS.CSY56SHCD > 0 " & _
                    "AND CS.CSY56SHYR = " & iSHYR & " " & _
                    "AND CS.CSY56CUNO > 0 " & _
                    "AND CS.CSY56CUNO = C.ABAN8 " & _
                    "ORDER BY UPPER(ABALPH)"
    Else
        strSelect = "SELECT DISTINCT CS.CSY56CUNO, C.ABALPH " & _
                "FROM " & F5611 & " CS, " & F0101 & " C " & _
                "WHERE CS.CSY56SHCD > 0 " & _
                "AND CS.CSY56SHYR = " & iSHYR & " " & _
                "AND CS.CSY56CUNO IN (" & strCunoList & ") " & _
                "AND CS.CSY56CUNO = C.ABAN8 " & _
                "ORDER BY UPPER(ABALPH)"
    End If
    Set rst = Conn.Execute(strSelect)
    combo.Clear
    Do While Not rst.EOF
        If Left(rst.Fields("ABALPH"), 2) <> "**" Then
            combo.AddItem UCase(Trim(rst.Fields("ABALPH")))
            combo.ItemData(combo.NewIndex) = rst.Fields("CSY56CUNO")
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub GetShows(combo As ComboBox, iSHYR As Integer, sCUNO As String)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT CS.CSY56SHCD, SM.SHY56NAMA " & _
                "FROM " & F5611 & " CS, " & F5601 & " SM " & _
                "WHERE CS.CSY56SHYR = " & iSHYR & " " & _
                "AND CS.CSY56CUNO = " & CLng(sCUNO) & " " & _
                "AND CS.CSY56SHCD = SM.SHY56SHCD " & _
                "AND CS.CSY56SHYR = SM.SHY56SHYR " & _
                "ORDER BY UPPER(SHY56NAMA)"
    Set rst = Conn.Execute(strSelect)
    combo.Clear
    Do While Not rst.EOF
        combo.AddItem UCase(Trim(rst.Fields("SHY56NAMA")))
        combo.ItemData(combo.NewIndex) = rst.Fields("CSY56SHCD")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub PopClientsWithInventory(combo As ComboBox)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT K.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM " & IGLKit & " K, " & F0101 & " C " & _
                    "WHERE K.KITID > 0 " & _
                    "AND K.KSTATUS > 0 " & _
                    "AND K.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY CLIENT"
    Else
        strSelect = "SELECT DISTINCT K.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM " & IGLKit & " K, " & F0101 & " C " & _
                    "WHERE K.KITID > 0 " & _
                    "AND K.AN8_CUNO IN (" & strCunoList & ") " & _
                    "AND K.KSTATUS > 0 " & _
                    "AND K.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY CLIENT"
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        combo.AddItem UCase(Trim(rst.Fields("ABALPH")))
        combo.ItemData(combo.NewIndex) = rst.Fields("AN8_CUNO")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Public Sub PopClientsWithEngProjects(combo As ComboBox)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    If bClientAll_Enabled Then
        strSelect = "SELECT DISTINCT K.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM " & IGLKit & " K, " & F0101 & " C " & _
                    "WHERE K.KITID > 0 " & _
                    "AND K.KSTATUS > 0 " & _
                    "AND K.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "UNION " & _
                    "SELECT DISTINCT EP.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM ANNOTATOR.ENG_PROGRAM EP, " & F0101 & " C " & _
                    "WHERE EP.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY CLIENT"
    Else
        strSelect = "SELECT DISTINCT K.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM " & IGLKit & " K, " & F0101 & " C " & _
                    "WHERE K.KITID > 0 " & _
                    "AND K.AN8_CUNO IN (" & strCunoList & ") " & _
                    "AND K.KSTATUS > 0 " & _
                    "AND K.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "UNION " & _
                    "SELECT DISTINCT EP.AN8_CUNO, C.ABALPH, UPPER(C.ABALPH) AS CLIENT " & _
                    "FROM ANNOTATOR.ENG_PROGRAM EP, " & F0101 & " C " & _
                    "WHERE EP.AN8_CUNO IN (" & strCunoList & ") " & _
                    "AND EP.AN8_CUNO = C.ABAN8 " & _
                    "AND C.ABAT1 = 'C' " & _
                    "ORDER BY CLIENT"
    End If
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        combo.AddItem UCase(Trim(rst.Fields("ABALPH")))
        combo.ItemData(combo.NewIndex) = rst.Fields("AN8_CUNO")
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing

End Sub

Function DeGlitch(sName As String) As String
    Dim iStr As Integer
    iStr = 1
    Do While InStr(iStr, sName, "'") <> 0
        sName = Left(sName, InStr(iStr, sName, "'")) & "'" & Mid(sName, InStr(iStr, sName, "'") + 1)
        iStr = InStr(iStr, sName, "'") + 2
    Loop
    DeGlitch = sName
End Function

Public Function CheckForTeam(tmpBCC As String, tmpSHCD As Long, frm As Form) As Boolean
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim bTeam As Boolean
    
    bTeamMember = False
    '///// FIRST, CHECK FOR CLIENT-SHOW TEAM \\\\\
    strSelect = "SELECT U.NAME_FIRST, U.NAME_LAST " & _
                "From " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                "WHERE T.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND T.AN8_SHCD = " & tmpSHCD & " " & _
                "AND T.MCU IS NULL " & _
                "AND T.TEAM_ID = R.TEAM_ID " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        '///// IF NOT FOUND, CHECK FOR CLIENT TEAM \\\\\
        strSelect = "SELECT U.NAME_FIRST, U.NAME_LAST " & _
                    "From " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                    "WHERE T.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                    "AND T.AN8_SHCD IS NULL " & _
                    "AND T.MCU IS NULL " & _
                    "AND T.TEAM_ID = R.TEAM_ID " & _
                    "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                    "AND U.USER_STATUS > 0"
        Set rst = Conn.Execute(strSelect)
    End If
    
    If Not rst.EOF Then
        bTeam = True
        Do While Not rst.EOF
            If StrConv(Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST")), vbProperCase) = LogName Then
                bTeamMember = True
                Debug.Print "I am a member of the team."
                GoTo FoundMe
            End If
            rst.MoveNext
        Loop
FoundMe:
        If frm.Name = "frmAnnotator" Then
            frm.lblReds.Caption = ""
            frm.mnuComments1.Enabled = True
        End If
    Else
        bTeam = False
        If frm.Name = "frmAnnotator" Then
            frm.lblReds.Caption = "No Email Team has been setup for " & _
                        frm.lblClient.Caption & " (Comments Interface is disabled)"
            frm.lblReds.Visible = True
            frm.mnuComments1.Enabled = False
        End If
    End If
    rst.Close
    Set rst = Nothing
    
    If frm.Name = "frmAnnotator" Then
        frm.imgTeam.ToolTipText = "Select to Edit " & frm.lblClient.Caption & " email notification team"
        If bPerm(0) Then frm.imgTeam.Visible = True
    End If
    
    If Not bTeamMember Then Debug.Print "I am NOT a member of the team."
    
    If bTeam = False And frm.mnuRedlining.Enabled = True Then frm.mnuRedlining.Enabled = False
    
    If frm.Name = "frmGraphics" Then
        If bTeamMember And bPerm(25) Then
            frmGraphics.mnuAssign.Visible = True
        Else
            frmGraphics.mnuAssign.Visible = False
        End If
    End If
    
'''''    If bTeam Then frm.mnuRedlining.Enabled = True Else frm.mnuRedlining.Enabled = False
    CheckForTeam = bTeam
End Function

Public Sub DriverAlert(sDriver As String, sIssues As String, sUser As String, sEMail As String, dDate As Date)
    Dim MessBody As String, MessHdr As String
    Dim Address As String
    Dim i As Integer
    
    ''Address = "Annotator Support Team"
    Address = "Steve.Westerholm@Project.com"
    
'''    Address = "swesterh@gpjco.com" ''FOR NOW, CHANGE TO SUPPORT TEAM''
    MessHdr = "Annotator Driver Request from " & UCase(sUser)
    MessBody = sUser & " has requested that the following Printer Driver(s) " & _
                "be made available on the Annotator.  " & vbNewLine & vbNewLine & _
                "Driver Name(s):" & vbNewLine & sDriver & vbNewLine & _
                "Issues Checked:" & vbNewLine & sIssues & vbNewLine & _
                "Once the Driver is available, please notify " & sUser & " (mailto:" & _
                sEMail & ")." & vbNewLine & _
                vbNewLine & _
                "Date & Time of Original Request:  " & Format(dDate, "dddd, mmmm d, yyyy") & _
                " @ " & Format(dDate, "h:nn ampm")
                
        
    '///// EXECUTE E-MAIL \\\\\
'''''    Dim myNotes As New Domino.NotesSession
'''''    Dim myDB As New Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    
    
    
    Dim MailMan As New ChilkatMailMan2 '' New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
        
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
        
    Dim Email As New ChilkatEmail2
    
    Email.AddTo Address, Address
        
    Email.FromAddress = LogAddress
    Email.fromName = LogName
            
    Email.subject = MessHdr
    Email.Body = MessBody
    
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
    
    
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
'        myDom.Initialize (sGAnnoPW)
'        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'        Set myDoc = myDomDB.CreateDocument
'
'        Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'    End If
'
'    Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
'    Set myRichText = myDoc.CreateRichTextItem("Body")
'    myRichText.AppendText MessBody
'    myDoc.AppendItemValue "SENDTO", Address
''''    myDoc.SaveMessageOnSend = True
'
''''    End With
'    On Error Resume Next
'    Call myDoc.Send(False, Address)
'    If Err Then
'        MsgBox "Printing has been disabled due to the absence of the correct Printer Driver " & _
'                    "on the Application Server, and an error has occurred during the automated " & _
'                    "request for the Driver to be made available.  Please contact the Help Desk " & _
'                    "to inform them of the error." & vbNewLine & vbNewLine & _
'                    "ERROR: " & Err.Description, vbExclamation, "Error Encountered..."
'        Err = 0
'        GoTo GetOut
'    Else
'        MsgBox "Request has been submitted.", vbExclamation, "Complete..."
'        With frmPrinterDrivers
'            .txtDrivers.Text = ""
'            For i = 0 To .chkIssue.Count - 1
'                .chkIssue(i).Value = 0
'            Next i
'        End With
'    End If
'
'GetOut:
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

Public Sub RedAlert(iType As Integer, sHDR As String, tmpBCC As String, tmpSHCD As Long)
    Dim MessBody As String, MessHdr As String
    Dim i As Integer, iAdd As Integer, iVal As Integer
    Dim tYear As String, tClient As String, tShow As String, sWho As String, _
                sEMail As String, sList As String, sInternet As String
    Dim iStr As Integer, iEnd As Integer
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim sColumn As String
    
    
    frmRedAlert.Show 1
    
    '///// EXECUTE E-MAIL \\\\\
'''''    Dim myNotes As New Domino.NotesSession
'''''    Dim myDB As New Domino.NotesDatabase
'    Dim myItem  As Object ''' NOTESITEM
'    Dim myDoc As Object ''' NOTESDOCUMENT
'    Dim myRichText As Object ' NOTESRICHTEXTITEM
'    Dim myReply  As Object ''' NOTESITEM
    Dim Address() As String
    
    
    
    Dim MailMan As New ChilkatMailMan2
    MailMan.UnlockComponent "MMZLLAMAILQ_fyMcFdWtpR9o"
    
    MailMan.SmtpSsl = 1
    MailMan.SmtpPort = 465
    MailMan.SmtpUsername = "smtp@project.com"
    MailMan.SmtpPassword = "Tosa5550"
    MailMan.SmtpHost = "smtp.gmail.com"
    
    Dim Email As New ChilkatEmail2
       
    
    '///// FIRST, GET TEAM \\\\\
    '///// SEE IF CLIENT-SHOW TEAM EXISTS \\\\\
    Select Case iType
        Case 0: sColumn = "RECIPIENT_FLAG0"
        Case 1: sColumn = "RECIPIENT_FLAG1"
        Case 2: sColumn = "RECIPIENT_FLAG2"
    End Select
    
    strSelect = "SELECT U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
                "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                "WHERE T.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND T.AN8_SHCD = " & tmpSHCD & " " & _
                "AND T.MCU IS NULL " & _
                "AND T.TEAM_ID = R.TEAM_ID " & _
                "AND R." & sColumn & " = 1 " & _
                "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                "AND U.USER_STATUS > 0 " & _
                "ORDER BY U.NAME_LAST, U.NAME_FIRST"
    Set rst = Conn.Execute(strSelect)
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        strSelect = "SELECT U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
                    "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
                    "WHERE T.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                    "AND T.AN8_SHCD IS NULL " & _
                    "AND T.MCU IS NULL " & _
                    "AND T.TEAM_ID = R.TEAM_ID " & _
                    "AND R." & sColumn & " = 1 " & _
                    "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
                    "AND U.USER_STATUS > 0 " & _
                    "ORDER BY U.NAME_LAST, U.NAME_FIRST"
        Set rst = Conn.Execute(strSelect)
    End If
    iAdd = 1: sList = ""
    Do While Not rst.EOF
        ReDim Preserve Address(iAdd)
        Address(iAdd - 1) = Trim(rst.Fields("EMAIL_ADDRESS"))
        sList = sList & vbTab & Trim(rst.Fields("NAME_FIRST")) & " " & Trim(rst.Fields("NAME_LAST"))
        
        Email.AddTo Address(iAdd - 1), Address(iAdd - 1)
        iAdd = iAdd + 1
    
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    
    MessHdr = sHDR & " Redlines"
    MessBody = "Redline Annotations have been drawn and saved by " & _
                LogName & " for " & sHDR & "." & vbNewLine & vbNewLine & _
                "The following Team members are being alerted through this email:" & _
                vbNewLine & vbNewLine & sList & vbNewLine & vbNewLine & _
                "Please forward to those, not included in this Email, that you feel should be alerted." & _
                vbCr
        
    Email.subject = MessHdr
    Email.Body = MessBody
     
    Email.FromAddress = LogAddress
    Email.fromName = LogName
    
    Dim Success As Integer
    Success = MailMan.SendEmail(Email)
    If (Success = 0) Then
        MsgBox MailMan.LastErrorText
    End If
    
    
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
'        myDom.Initialize (sGAnnoPW)
'        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
'        Set myDoc = myDomDB.CreateDocument
'
'        Call myDoc.ReplaceItemValue("Principal", LogName)
'        Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
'    End If
'
'    Set myItem = myDoc.AppendItemValue("Subject", MessHdr)
'    Set myRichText = myDoc.CreateRichTextItem("Body")
'    With myRichText
'        .AppendText MessBody & vbNewLine & vbNewLine & sLink
'        .AddNewLine 2
'        .AppendText LogName
'    End With
'    myDoc.AppendItemValue "SENDTO", Address(i)
''''    myDoc.SaveMessageOnSend = True
'
'    On Error Resume Next
'    For i = 0 To iAdd - 2
'        Call myDoc.Send(False, Address(i))
'    Next i
'    If Err Then
'        MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
'                    vbExclamation, "Error Encountered"
'        Err = 0
'        GoTo GetOut
'    End If
'
'GetOut:
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

Public Function IGLToJDEDate(dDate As String) As Long
    Dim sJDECent As String, sJDEYear As String, sJDEDay As String
    Dim iIGLYear As Integer
    
    iIGLYear = CInt(Format(dDate, "yyyy"))
    sJDEYear = CStr(Right(iIGLYear, 2))
    If iIGLYear > 1999 Then sJDECent = "1" Else sJDECent = "0"
    sJDEDay = Right("000" & Format(dDate, "y"), 3)
    IGLToJDEDate = CLng(sJDECent & sJDEYear & sJDEDay)
End Function

Public Function Check_Printers(bChk As Boolean) As Boolean
    Dim sRegPrinter As String, sMess As String
    Dim processes As CVector
    Dim i As Long
    Dim bCTXLOGON As Boolean
    
    On Error Resume Next
    bPRINTER_ALERT = True
    sMess = "NOTE:  The Annotator has not found a compatible Printer.  Printing is disabled.  " & _
                "See the 'Printer Drivers...' interface on the 'Options' menu ('Key' icon)."
    sPrinter = Printer.DeviceName
    
    If sPrinter = "" Or Err > 0 Then
        If bChk Then
            bCTXLOGON = False
            Set processes = CreateProcessList
            For i = 1 To processes.Last
                If UCase(processes(i).EXEName) = "CTXLOGON" Then
                    ''CTXLOGIN IS STILL RUNNING''
                    MsgBox "Your logon is still in the process of auto-creating " & _
                                "your printers.  Depending on the quantity of " & _
                                "printers you have setup, this make take up to " & _
                                "3 or 4 minutes after login." & vbNewLine & vbNewLine & _
                                "Please, wait a moment before trying to print again.", vbExclamation, _
                                "Printers are currently being setup..."
                    bCTXLOGON = True
                    bPRINTER_ALERT = False
                    Exit For
                End If
            Next
            Set processes = Nothing
            
            
            If Not bCTXLOGON Then
                ''CTXLOGON IS NOT RUNNING, BUT THERE IS NO PRINTER''
                ''DOUBLE-CHECK TO SEE IF A PRINTER WAS SETUP DURING THIS PROCEDURE''
                Err.Clear
                sPrinter = Printer.DeviceName
                If sPrinter = "" Or Err > 0 Then
                    MsgBox "The Annotator has not found any printers with compatible " & _
                                "printer drivers on your computer." & vbNewLine & vbNewLine & _
                                "In order to enable printing you will need to submit " & _
                                "the name of the driver your default printer is using.  " & _
                                "To do this, go to your Printers interface and right-click " & _
                                "on your default printer icon.  Select 'Properties'.  " & _
                                "The driver name should be listed here.  You will need to go " & _
                                "to the Annotator's 'Printer Drivers...' interface on the " & _
                                "StartUp screen (the 'Key' icon), enter the driver name, " & _
                                "check the printer problem option, and click 'Submit'." & _
                                vbNewLine & vbNewLine & _
                                "If you wish, you can contact the GPJ Helpdesk to assist you.", _
                                vbInformation, "No compatible printer was found..."
                            
                    ''LET USER KNOW THAT THERE IS A PRINTING ISSUE''
                    bENABLE_PRINTERS = False
                    frmStartUp.lblDriver = sMess
                    Check_Printers = False
                Else
                    ''CTXLOGON IS STILL RUNNING''
                    ''PRINTING IS OK''
                    bENABLE_PRINTERS = True
                    frmStartUp.lblDriver = ""
                    Check_Printers = False
                End If
            Else
                ''LET USER KNOW THAT THERE IS A PRINTING ISSUE''
                bENABLE_PRINTERS = False
                frmStartUp.lblDriver = sMess
                Check_Printers = True
            End If
        Else
            ''LET USER KNOW THAT THERE IS A PRINTING ISSUE''
            bENABLE_PRINTERS = False
            frmStartUp.lblDriver = sMess
            Check_Printers = True
            
        End If
        
        
    Else
        bENABLE_PRINTERS = True
        frmStartUp.lblDriver = ""
        Check_Printers = False
'''        bDo_Printer_Check = False
    End If
End Function

Public Sub OpenConn(sConnStr As String)
    Set Conn = Nothing
    Set Conn = New ADODB.Connection
    Conn.Open (sConnStr)
    ConnOpen = True
End Sub

Public Sub GetDrawings(frm As Form, tFileType As String, tDWGID As Long, _
            tSHYR As Integer, tSHCD As Long)
    Dim strSelect As String, sCheck As String, sPDF As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT NVL(DWFDESC, 'UNKNOWN') AS DWFDESC, DWFID, DWFPATH " & _
                "FROM " & DWGDwf & " " & _
                "WHERE DWGID = " & tDWGID & " " & _
                "AND DWFSTATUS > 0 " & _
                "ORDER BY DWFDESC"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case tFileType
            Case "DWF"
                frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                frm.lstPaths.AddItem Trim(rst.Fields("DWFPATH"))
                frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(frm.lstPaths.List(frm.lstPaths.NewIndex))
            Case "PDF"
                sCheck = UCase(Trim(rst.Fields("DWFPATH")))
                sCheck = Left(sCheck, Len(sCheck) - 3) & "pdf"
                sCheck = Dir(sCheck)
                If sCheck <> "" Then
                    frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                    sPDF = UCase(Trim(rst.Fields("DWFPATH")))
                    sPDF = Left(sPDF, Len(sPDF) - 3) & "pdf"
                    frm.lstPaths.AddItem sPDF
                    frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(sPDF)
                End If
        End Select
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    

    strSelect = "SELECT M.DWGID, M.DWGTYPE, DWF.DWFPATH, DWF.DWFDESC " & _
                "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "Where SHO.SHYR = " & tSHYR & " " & _
                "AND SHO.AN8_SHCD = " & tSHCD & " " & _
                "AND SHO.DWGID = M.DWGID " & _
                "AND M.DWGTYPE IN (3, 4) " & _
                "AND M.DSTATUS > 0 " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND M.DWGID = DWF.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "ORDER BY M.DWGTYPE"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case tFileType
            Case "DWF"
                Select Case rst.Fields("DWGTYPE")
                    Case 3
                        If bPerm(4) Then
                            frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                            frm.lstPaths.AddItem Trim(rst.Fields("DWFPATH"))
                            frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(frm.lstPaths.List(frm.lstPaths.NewIndex))
                        End If
                    Case 4
                        If bPerm(5) Then
                            frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                            frm.lstPaths.AddItem Trim(rst.Fields("DWFPATH"))
                            frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(frm.lstPaths.List(frm.lstPaths.NewIndex))
                        End If
                End Select
            Case "PDF"
                Select Case rst.Fields("DWGTYPE")
                    Case 3
                        If bPerm(4) Then
                            sCheck = UCase(Trim(rst.Fields("DWFPATH")))
                            sCheck = Left(sCheck, Len(sCheck) - 3) & "pdf"
                            sCheck = Dir(sCheck)
                            If sCheck <> "" Then
                                frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                                sPDF = UCase(Trim(rst.Fields("DWFPATH")))
                                sPDF = Left(sPDF, Len(sPDF) - 3) & "pdf"
                                frm.lstPaths.AddItem sPDF
                                frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(sPDF)
                            End If
                        End If
                    Case 4
                        If bPerm(5) Then
                            sCheck = UCase(Trim(rst.Fields("DWFPATH")))
                            sCheck = Left(sCheck, Len(sCheck) - 3) & "pdf"
                            sCheck = Dir(sCheck)
                            If sCheck <> "" Then
                                frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                                sPDF = UCase(Trim(rst.Fields("DWFPATH")))
                                sPDF = Left(sPDF, Len(sPDF) - 3) & "pdf"
                                frm.lstPaths.AddItem sPDF
                                frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(sPDF)
                            End If
                        End If
                End Select
            
        End Select
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Public Sub GetShowPlans(frm As Form, tFileType As String, tSHYR As Integer, _
            tSHCD As Long)
    Dim strSelect As String, sCheck As String, sPDF As String
    Dim rst As ADODB.Recordset
    
    strSelect = "SELECT M.DWGID, M.DWGTYPE, DWF.DWFPATH, DWF.DWFDESC " & _
                "FROM " & DWGShow & " SHO, " & DWGMas & " M, " & DWGSht & " SHT, " & DWGDwf & " DWF " & _
                "Where SHO.SHYR = " & tSHYR & " " & _
                "AND SHO.AN8_SHCD = " & tSHCD & " " & _
                "AND SHO.DWGID = M.DWGID " & _
                "AND M.DWGTYPE IN (3, 4) " & _
                "AND M.DSTATUS > 0 " & _
                "AND M.DWGID = SHT.DWGID " & _
                "AND M.DWGID = DWF.DWGID " & _
                "AND SHT.DWGID = DWF.DWGID " & _
                "AND SHT.SHTID = DWF.SHTID " & _
                "ORDER BY M.DWGTYPE"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Select Case tFileType
            Case "DWF"
                frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                frm.lstPaths.AddItem Trim(rst.Fields("DWFPATH"))
                frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(frm.lstPaths.List(frm.lstPaths.NewIndex))
            Case "PDF"
                If rst.Fields("DWGTYPE") = 3 Or _
                            (rst.Fields("DWGTYPE") = 4 And bPerm(5) = True) Then
                    sCheck = UCase(Trim(rst.Fields("DWFPATH")))
                    sCheck = Left(sCheck, Len(sCheck) - 3) & "pdf"
                    sCheck = Dir(sCheck)
                    If sCheck <> "" Then
                        frm.lstPlans.AddItem UCase(Trim(rst.Fields("DWFDESC")))
                        sPDF = UCase(Trim(rst.Fields("DWFPATH")))
                        sPDF = Left(sPDF, Len(sPDF) - 3) & "pdf"
                        frm.lstPaths.AddItem sPDF
                        frm.lstPaths.ItemData(frm.lstPaths.NewIndex) = FileLen(sPDF)
                    End If
                End If
        End Select
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Public Function InsertComment(lGID As Long, Index As Integer, sNote As String) As Integer
    Dim strInsert As String
    Dim rstL As ADODB.Recordset
    Dim lCOMMID As Long
    Dim sComm As String
    
    On Error Resume Next
    
    Select Case Index
        Case 0: sComm = "Graphic Status reset to 'INTERNAL DRAFT' by " & LogName & "."
        Case 1: sComm = "Graphic Status reset to 'CLIENT DRAFT' by " & LogName & "."
        Case 2: sComm = "Graphic 'APPROVED' by " & LogName & "."
        Case 3: sComm = "Graphic Cancelled by " & LogName & "."
        Case 4: sComm = "Graphic 'RETURNED FOR CHANGES' by " & LogName & "."
    End Select
    If sNote <> "" Then sComm = sComm & " (Comment: " & sNote & ")"
    
    '///// GET NEW COMMID \\\\\
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    lCOMMID = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing

    strInsert = "INSERT INTO " & ANOComment & " " & _
            "(COMMID, REFID, REFSOURCE, ANO_COMMENT, " & _
            "COMMUSER, COMMDATE, COMMSTATUS, " & _
            "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
            "VALUES " & _
            "(" & lCOMMID & ", " & lGID & ", 'GFX_MASTER', '" & DeGlitch(sComm) & "', " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
    
    InsertComment = Err.Number
    
End Function

Public Function CreateMessage(strGList As String, sCC As String, sShortname As String, _
            sComm As String) As String
    Dim strSelect As String, sMess As String, _
                sDefault As String, sComment As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iStatus As Integer
    Dim sGType(1 To 4) As String
    Dim sGStatus(2 To 30) As String
    
    '///// FILE TYPE VARIABLES \\\\\
    sGType(1) = "Photo File"
    sGType(2) = "Graphic File"
    sGType(3) = "Graphic Layout"
    sGType(4) = "Presentation File"
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(2) = "CANCELED"
    sGStatus(3) = "CANCELED"
    sGStatus(4) = "CANCELED"
    sGStatus(10) = "INTERNAL DRAFT"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(30) = "APPROVED"
    sGStatus(27) = "RETURNED FOR CHANGES"
    
    sDefault = "Those Graphics posted with an 'INTERNAL DRAFT' status, need to be reviewed, " & _
                "and their status advanced to 'CLIENT DRAFT' or 'APPROVED', before " & _
                "they will be available for viewing by non-GPJ personnel.  Graphics noted " & _
                "with a 'RETURNED FOR CHANGES' status, have been reviewed and NOT Approved " & _
                "for the Commented Reason."
'''    sList = ""
'''    For i = 0 To lstFiles.ListCount - 1
'''        If lstFiles.Selected(i) = True Then
'''            If sList = "" Then sList = CStr(lstFiles.ItemData(i)) _
'''                        Else sList = sList & ", " & CStr(lstFiles.ItemData(i))
'''        End If
'''    Next i
    
    If strGList <> "" Then
'        sCC = "cc:" & GetGMRs
        
'''        For i = 0 To lstTeam.ListCount - 1
'''            If lstTeam.Selected(i) = True Then
'''                sCC = sCC & vbTab & lstTeam.List(i) & vbNewLine
'''            End If
'''        Next i
'''        If chkReceive.value = 1 Then sCC = sCC & vbTab & UCase(LogName) & vbNewLine
'        sCC = sCC & vbNewLine & "From:" & vbTab & UCase(LogName) & vbNewLine
        
        sMess = "": iStatus = -1
        strSelect = "SELECT AB.ABALPH AS CLIENT, GM.GID, GM.GDESC, GM.GTYPE, " & _
                    "GM.GAPPROVER_ID, (TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) AS FNAME, " & _
                    "GM.GSTATUS, GM.ADDDTTM " & _
                    "FROM " & GFXMas & " GM, " & F0101 & " AB, IGLPROD.IGL_USER U " & _
                    "WHERE GM.GID IN (" & strGList & ") " & _
                    "AND GM.GAPPROVER_ID = U.USER_SEQ_ID " & _
                    "AND GM.AN8_CUNO = AB.ABAN8 " & _
                    "ORDER BY AB.ABALPH, GM.GSTATUS, GM.GTYPE, GM.GDESC"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            sMess = sMess & UCase(Trim(rst.Fields("CLIENT"))) & " GRAPHIC FILES:" & vbNewLine
            Do While Not rst.EOF
                If iStatus <> rst.Fields("GSTATUS") Then
                    iStatus = rst.Fields("GSTATUS")
                    Select Case iStatus
                        Case 2, 3, 4
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (The following Files have been Deleted)" & vbNewLine
                        Case 10
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Available for Viewing by internal GPJ Only)" & vbNewLine
                        Case 20
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Released for Client Viewing, Awaiting Final Approval)" & vbNewLine
                        Case 27
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Reviewed and Returned for Changes)" & vbNewLine
                        Case 30
                            sMess = sMess & vbNewLine & Space(4) & sGStatus(iStatus) & _
                                        "  (Reviewed and Approved for Production)" & vbNewLine
                    End Select
                End If
                Select Case iStatus
                    Case 2, 3, 4
                        sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                    UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                    vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                    vbNewLine
                    Case 27
'''                        sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
'''                                    UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
'''                                    vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
'''                                    vbTab & " -- Comment:  " & Trim(sComm) & vbNewLine & _
'''                                    vbNewLine '''REMOVE THIS LINE'''& _
'''                                    vbtab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
''''                                    "?name_logon=" & LCASE(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
''''                                    vbNewLine
                        If rst.Fields("GAPPROVER_ID") > 0 Then
                            sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                        UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                        vbTab & " -- Assigned Approver:  " & Trim(rst.Fields("FNAME")) & vbNewLine & _
                                        vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                        vbTab & " -- Comment:  " & Trim(sComm) & vbNewLine & _
                                        vbTab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
                                        "?name_logon=" & LCase(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
                                        vbNewLine
                        Else
                            sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                        UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                        vbTab & " -- Assigned Approver:  None" & vbNewLine & _
                                        vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                        vbTab & " -- Comment:  " & Trim(sComm) & vbNewLine & _
                                        vbTab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
                                        "?name_logon=" & LCase(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
                                        vbNewLine
                        End If
                    Case Else
                        If rst.Fields("GAPPROVER_ID") > 0 Then
                            sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                        UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                        vbTab & " -- Assigned Approver:  " & Trim(rst.Fields("FNAME")) & vbNewLine & _
                                        vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                        vbTab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
                                        "?name_logon=" & LCase(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
                                        vbNewLine
                        Else
                            sMess = sMess & vbTab & "Posted " & sGType(rst.Fields("GTYPE")) & ":  " & _
                                        UCase(Trim(rst.Fields("GDESC"))) & vbNewLine & _
                                        vbTab & " -- Assigned Approver:  None" & vbNewLine & _
                                        vbTab & " -- Status:  " & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                                        vbTab & " -- Link:  http://gpjapps02.gpjco.com/LinksToAnno.asp" & _
                                        "?name_logon=" & LCase(sShortname) & "&gid=" & rst.Fields("GID") & vbNewLine & _
                                        vbNewLine
                        End If
                End Select

                rst.MoveNext
            Loop
        End If
        rst.Close: Set rst = Nothing
        
        If sComm <> "" And iStatus <> 27 Then
            sComment = vbNewLine & vbNewLine & "POSTERS COMMENTS (" & LogName & "):" & _
                        vbNewLine & Trim(sComm)
        Else
            sComment = ""
        End If
        
        CreateMessage = "cc:" & sCC & vbNewLine & vbNewLine & _
                    "The following Graphic Files have been posted to the GPJ Annotator, " & _
                    "and are ready for your review.  " & sDefault & sComment & _
                    vbNewLine & vbNewLine & sLink & _
                    vbNewLine & vbNewLine & vbNewLine & sMess
    Else
        CreateMessage = ""
    End If
End Function


'''Public Function GetGMRs(iType As Integer, lBCC As Long, iStatus As Integer) As String
'''    Dim strSelect As String, sColumn As String
'''    Dim rst As ADODB.Recordset
'''    Dim iAdd As Integer
'''
'''    sCC = ""
'''    Select Case iType
'''        Case 0: sColumn = "RECIPIENT_FLAG0"
'''        Case 1: sColumn = "RECIPIENT_FLAG1"
'''        Case 2: sColumn = "RECIPIENT_FLAG2"
'''    End Select
'''    Select Case iStatus
'''        Case 10 ''INTERNAL GPJ ONLY''
'''            strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
'''                        "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
'''                        "WHERE T.AN8_CUNO = " & lBCC & " " & _
'''                        "AND T.AN8_SHCD IS NULL " & _
'''                        "AND T.MCU IS NULL " & _
'''                        "AND T.TEAM_ID = R.TEAM_ID " & _
'''                        "AND R." & sColumn & " = 1 " & _
'''                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
'''                        "AND UPPER(SUBSTR(U.EMPLOYER, 1, 3)) = 'GPJ' " & _
'''                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
'''        Case Else ''AVAILABLE TO ALL''
'''            strSelect = "SELECT U.NAME_LOGON, U.NAME_LAST, U.NAME_FIRST, U.EMAIL_ADDRESS " & _
'''                        "FROM " & ANOETeam & " T, " & ANOETeamUR & " R, " & IGLUser & " U " & _
'''                        "WHERE T.AN8_CUNO = " & lBCC & " " & _
'''                        "AND T.AN8_SHCD IS NULL " & _
'''                        "AND T.MCU IS NULL " & _
'''                        "AND T.TEAM_ID = R.TEAM_ID " & _
'''                        "AND R." & sColumn & " = 1 " & _
'''                        "AND R.USER_SEQ_ID = U.USER_SEQ_ID " & _
'''                        "ORDER BY U.NAME_LAST, U.NAME_FIRST"
'''    End Select
'''    Set rst = Conn.Execute(strSelect)
'''
'''    iAdd = -1: sCC = ""
'''    Do While Not rst.EOF
'''        iAdd = iAdd + 1
'''        ReDim Preserve GFXAddress(iAdd)
'''        ReDim Preserve GFXMandRecip(iAdd)
'''        GFXAddress(iAdd) = Trim(rst.Fields("EMAIL_ADDRESS"))
'''        GFXMandRecip(iAdd) = LCase(Trim(rst.Fields("NAME_LOGON")))
'''        sCC = sCC & vbTab & Trim(rst.Fields("NAME_FIRST")) & " " & _
'''                    Trim(rst.Fields("NAME_LAST")) & vbNewLine
'''        rst.MoveNext
'''    Loop
'''    rst.Close
'''    Set rst = Nothing
'''
'''    GetGMRs = sCC
'''End Function

''Public Sub SendGFXEmail(tBCC As Long, tFBCN As String, sGList As String, sComm As String, _
''            iStatus As Integer)
''    Dim sMess As String, sList1 As String, sList2 As String, sList3 As String, _
''                sHDR As String, strUpdate As String, strDelete As String, strSelect As String
''    Dim rst As ADODB.Recordset
'''''''    Dim myNotes As New Domino.NotesSession
'''''''    Dim myDB As New Domino.NotesDatabase
''    Dim myItem  As Object ''' NOTESITEM
''    Dim myReply  As Object ''' NOTESITEM
''    Dim myDoc As Object ''' NOTESDOCUMENT
''    Dim myRichText As Object ''' NOTESRICHTEXTITEM
''    Dim i As Integer, iAdd As Integer, iRecips As Integer
'''''    Dim sNote As String
''
''    Screen.MousePointer = 11
'''''    sNote = vbNewLine & vbNewLine & _
'''''                "NOTE:  The direct access hyperlinks contained in " & _
'''''                "this email have been specifically created for the " & _
'''''                "original receiver's access only.  Please, do not " & _
'''''                "forward this document to others.  If you have been " & _
'''''                "forwarded this document, the link will not function " & _
'''''                "for you, unless you as able to login as the " & _
'''''                "original receiver."
''
''    iRecips = GetRecips(iStatus)
''    If iRecips = -1 Then GoTo GetOut
''
'''''    Call GetApprovers(sGList, iStatus)
''
''    If Not bCitrix Then
''        ''APP IS RUNNING LOCAL OR THIN-CLIENT - LOTUS NOTES''
''        Dim myNotes As Object '' LOTUS.NotesSession '' NotesSession
''        Dim myDB As Object '' LOTUS.NotesDatabase
''
''
''        On Error Resume Next
''        Set myNotes = GetObject(, "Notes.NotesSession")
''
''        If Err Then
''            Err.Clear
''            Set myNotes = CreateObject("Notes.NotesSession")
''            If Err Then
''                MsgBox "Lotus Notes must exist locally to execute E-mail.", vbCritical, "Uh,oh..."
''                GoTo GetOut
''            End If
''        End If
''        On Error GoTo 0
''        Set myDB = myNotes.GetDatabase("", "")
''        myDB.OPENMAIL
''        Set myDoc = myDB.CreateDocument
''
''    Else
''        ''APP IS RUNNING ON CITRIX - USE DOMINO OBJECT''
''        Dim myDom As New Domino.NotesSession '''myNotes As Object ' NOTESSESSION
''        Dim myDomDB As New Domino.NotesDatabase '''myDB As Object ' NOTESDATABASE
''
''        myDom.Initialize (sGAnnoPW)
''        Set myDomDB = myDom.GetDatabase("Global_Links/IBM/GPJNotes", "mail\gannotat.nsf")
''        Set myDoc = myDomDB.CreateDocument
''
''        Call myDoc.ReplaceItemValue("Principal", LogName)
''        Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
''    End If
''
''    For i = LBound(GFXMandRecip) To UBound(GFXMandRecip)
''        sMess = ""
''        sMess = CreateMessage(sGList, sCC, GFXMandRecip(i), sComm)
''        If sMess <> "" Then
''            sMess = sMess '''& sNote
''
''            ''///// NOW, SEND OUT ALERT \\\\\''
''            sHDR = tFBCN & " -- GPJ Annotator Graphics Posting Alert"
''
'''            myNotes.Initialize
''
''            If Not bCitrix Then
''                Set myDoc = myDB.CreateDocument
''            Else
''                Set myDoc = myDomDB.CreateDocument
''                Call myDoc.ReplaceItemValue("Principal", LogName)
''                Set myReply = myDoc.AppendItemValue("ReplyTo", LogAddress)
''            End If
''
''            Set myItem = myDoc.AppendItemValue("Subject", sHDR)
''            Set myRichText = myDoc.CreateRichTextItem("Body")
''            myRichText.AppendText sMess & vbNewLine & vbNewLine ''& sLink_Disclaimer
''            myDoc.AppendItemValue "SENDTO", GFXAddress(i)
'''''            myDoc.SaveMessageOnSend = True
''
''            On Error Resume Next
''            Call myDoc.Send(False, GFXAddress(i))
''            If Err Then
''                MsgBox "ERROR: " & Err.Description & vbCr & vbCr & "Function Cancelled", _
''                            vbExclamation, "Error Encountered"
''                Err = 0
''                GoTo GetOut
''            End If
''
''            Set myRichText = Nothing
''            Set myItem = Nothing
''            Set myDoc = Nothing
''            Set myReply = Nothing
''        End If
''    Next i
''
''GetOut:
''    Set myReply = Nothing
''    Set myRichText = Nothing
''    Set myItem = Nothing
''    Set myDoc = Nothing
''
''    On Error Resume Next
''    If bCitrix Then
''        If Not myDomDB Is Nothing Then Set myDomDB = Nothing
''        If Not myDom Is Nothing Then Set myDom = Nothing
''    Else
''        If Not myDB Is Nothing Then Set myDB = Nothing
''        If Not myNotes Is Nothing Then Set myNotes = Nothing
''    End If
''
''    Screen.MousePointer = 0
''
''End Sub

Public Sub GetApprovers(sList As String, iStatus As Integer)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim i As Integer, iArray As Integer
    Dim bFound As Boolean
    
    iArray = UBound(GFXMandRecipFull)
    Select Case iStatus
        Case 10 ''GPJ INTERNAL ONLY''
            strSelect = "SELECT DISTINCT GM.GAPPROVER_ID, " & _
                        "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) APPROVER, " & _
                        "U.NAME_LOGON, U.EMAIL_ADDRESS AS EMAIL " & _
                        "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                        "WHERE GM.GID IN (" & sList & ") " & _
                        "AND GM.GAPPROVER_ID > 0 " & _
                        "AND GM.GAPPROVER_ID IS NOT NULL " & _
                        "AND GM.GAPPROVER_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0 " & _
                        "AND UPPER(SUBSTR(U.EMPLOYER, 1, 3)) = 'GPJ'"
        Case Else ''ALL USERS''
            strSelect = "SELECT DISTINCT GM.GAPPROVER_ID, " & _
                        "(TRIM(U.NAME_FIRST)||' '||TRIM(U.NAME_LAST)) APPROVER, " & _
                        "U.NAME_LOGON, U.EMAIL_ADDRESS AS EMAIL " & _
                        "FROM " & GFXMas & " GM, " & IGLUser & " U " & _
                        "WHERE GM.GID IN (" & sList & ") " & _
                        "AND GM.GAPPROVER_ID > 0 " & _
                        "AND GM.GAPPROVER_ID IS NOT NULL " & _
                        "AND GM.GAPPROVER_ID = U.USER_SEQ_ID " & _
                        "AND U.USER_STATUS > 0"
    End Select
    
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        bFound = False
        For i = LBound(GFXMandRecipFull) To UBound(GFXMandRecipFull)
            If UCase(GFXMandRecipFull(i)) = UCase(Trim(rst.Fields("APPROVER"))) Then
                bFound = True
                Exit For
            End If
        Next i
        If bFound = False Then
            If UCase(Trim(rst.Fields("APPROVER"))) <> UCase(LogName) Then
                sCC = sCC & vbTab & Trim(rst.Fields("APPROVER")) & vbNewLine ''' " " & _
                            Trim(rst.Fields("NAME_LAST")) & vbNewLine
                iArray = iArray + 1
                ReDim Preserve GFXMandRecipFull(iArray)
                ReDim Preserve GFXMandRecip(iArray)
                ReDim Preserve GFXAddress(iArray)
                GFXMandRecipFull(iArray) = Trim(rst.Fields("APPROVER"))
                GFXMandRecip(iArray) = Trim(rst.Fields("NAME_LOGON"))
                GFXAddress(iArray) = Trim(rst.Fields("EMAIL"))
            End If
        End If
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
End Sub

Public Function InsertGfxComment(pGID As Long, pNote As String) As Integer
    Dim strInsert As String
    Dim rstL As ADODB.Recordset
    Dim lCOMMID As Long
    Dim sComm As String
    
    On Error Resume Next
    
    '///// GET NEW COMMID \\\\\
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    lCOMMID = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing

    strInsert = "INSERT INTO " & ANOComment & " " & _
            "(COMMID, REFID, REFSOURCE, ANO_COMMENT, " & _
            "COMMUSER, COMMDATE, COMMSTATUS, " & _
            "ADDUSER, ADDDTTM, UPDUSER, UPDDTTM, UPDCNT) " & _
            "VALUES " & _
            "(" & lCOMMID & ", " & pGID & ", 'GFX_MASTER', '" & DeGlitch(pNote) & "', " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1, " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, " & _
            "'" & DeGlitch(Left(LogName, 24)) & "', SYSDATE, 1)"
    Conn.Execute (strInsert)
    
    InsertGfxComment = Err.Number
End Function


Public Sub GetGFXData(strSelect As String, sDisplay As String)
    Dim sMess As String, sSize As String
    Dim rst As ADODB.Recordset
    Dim sGStatus(0 To 30) As String
    Dim lSize As Long
    
    '///// FILE STATUS VARIABLES \\\\\
    sGStatus(0) = "DE-ACTIVED"
    sGStatus(10) = "INTERNAL"
    sGStatus(20) = "CLIENT DRAFT"
    sGStatus(27) = "RETURNED FOR CHANGES"
    sGStatus(30) = "APPROVED"
    
    
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        Select Case FileLen(Trim(rst.Fields("GPATH")))
            Case Is < 1000: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))), "#,##0") & " bytes"
            Case Is < 2000000: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))) / 1000, "#,##0") & " k"
            Case Else: sSize = Format(FileLen(Trim(rst.Fields("GPATH"))) / 1000000, "#,##0.00") & " mb"
        End Select
            
        sMess = "Graphic Description:" & vbTab & Trim(rst.Fields("GDESC")) & vbNewLine & _
                    "Database I.D.:          " & vbTab & rst.Fields("GID") & vbNewLine & _
                    "File Format:              " & vbTab & Trim(rst.Fields("GFORMAT")) & vbNewLine & _
                    "Graphic Type:           " & vbTab & GfxType(rst.Fields("GTYPE")) & vbNewLine & _
                    "Graphic Status:         " & vbTab & sGStatus(rst.Fields("GSTATUS")) & vbNewLine & _
                    "File Size:                   " & vbTab & sSize & vbNewLine & vbNewLine
        sMess = sMess & "File Added by " & Trim(rst.Fields("ADDUSER")) & " on " & _
                    Format(rst.Fields("ADDDTTM"), "mmmm d, yyyy") & "." & vbNewLine
        sMess = sMess & "File Last Edited by " & Trim(rst.Fields("UPDUSER")) & " on " & _
                    Format(rst.Fields("UPDDTTM"), "mmmm d, yyyy") & "."
        rst.Close
        Select Case sDisplay
            Case "msgbox"
                MsgBox sMess, vbInformation, "Graphic Data..."
'            Case "control"
'                txtXData(1).Text = sMess
        End Select
'        picXData.Visible = True
    Else
        rst.Close
'''        picXData.Visible = False
        MsgBox "No Data Available.", vbInformation, "Graphic Data..."
    End If
    Set rst = Nothing

End Sub

Public Function GetRecips(tStatus As Integer)
    Dim iAdd As Integer, i As Integer
    
    sCC = ""
    iAdd = -1
    For i = 0 To frmGfxApprove.lstTeam.ListCount - 1
        If frmGfxApprove.lstTeam.Selected(i) = True Then
            iAdd = iAdd + 1
            ReDim Preserve GFXAddress(iAdd)
            ReDim Preserve GFXMandRecip(iAdd)
            ReDim Preserve GFXMandRecipFull(iAdd)
            GFXAddress(iAdd) = frmGfxApprove.lstTeamEmail.List(i)
            GFXMandRecip(iAdd) = frmGfxApprove.lstTeamShort.List(i)
            GFXMandRecipFull(iAdd) = frmGfxApprove.lstTeam.List(i)
            sCC = sCC & vbTab & frmGfxApprove.lstTeam.List(i) & vbNewLine
        End If
    Next i
    
    If tStatus > 10 Then
        For i = 0 To frmGfxApprove.lstTeamClients.ListCount - 1
            If frmGfxApprove.lstTeamClients.Selected(i) = True Then
                iAdd = iAdd + 1
                ReDim Preserve GFXAddress(iAdd)
                ReDim Preserve GFXMandRecip(iAdd)
                ReDim Preserve GFXMandRecipFull(iAdd)
                GFXAddress(iAdd) = frmGfxApprove.lstClientEmail.List(i)
                GFXMandRecip(iAdd) = frmGfxApprove.lstClientShort.List(i)
                GFXMandRecipFull(iAdd) = frmGfxApprove.lstTeamClients.List(i)
                sCC = sCC & vbTab & frmGfxApprove.lstTeamClients.List(i) & vbNewLine
            End If
        Next i
    End If
    
    If frmGfxApprove.sstEmail.TabVisible(1) Then
        For i = 0 To frmGfxApprove.lstGPJ.ListCount - 1
            If frmGfxApprove.lstGPJ.Selected(i) = True Then
                iAdd = iAdd + 1
                ReDim Preserve GFXAddress(iAdd)
                ReDim Preserve GFXMandRecip(iAdd)
                ReDim Preserve GFXMandRecipFull(iAdd)
                GFXAddress(iAdd) = frmGfxApprove.lstGPJEmail.List(i)
                GFXMandRecip(iAdd) = frmGfxApprove.lstGPJShort.List(i)
                GFXMandRecipFull(iAdd) = frmGfxApprove.lstGPJ.List(i)
                sCC = sCC & vbTab & frmGfxApprove.lstGPJ.List(i) & vbNewLine
            End If
        Next i
    End If
    
    GetRecips = iAdd
End Function

Public Sub UnselectSelf(pName As String, pList As ListBox)
    Dim i As Integer
    For i = 0 To pList.ListCount - 1
        If UCase(Left(pList.List(i), Len(pName))) = UCase(pName) Then
            pList.ItemData(i) = 0
            pList.Selected(i) = False
        End If
    Next i
End Sub

Public Function PopShowInfo(tmpBCC As Long, tmpSHYR As Integer, tmpSHCD As Long, _
            tmpSHNM As String, tmpFBCN As String) As String
    Dim rst As ADODB.Recordset, rstX As ADODB.Recordset
    Dim strSelect As String, sHTML As String, sDate1 As String, sDate2 As String, tFile1 As String
    Dim i As Integer
    Dim htmO As String, htmC As String
    Dim hdO As String, hdC As String
    Dim tiO As String, tiC As String
    Dim bodO As String, bodC As String
    Dim f1O As String, f2O As String, f3O As String, fC As String
    Dim bolO As String, bolC As String
    Dim tblO As String, tblC As String
    Dim trO As String, trC As String
    Dim tdc2O As String, tdc3O As String, tdc4O As String, tdcC As String, tdOa As String, tdOb As String, tdC As String
    Dim tdNO As String, tdNC As String
    Dim hr As String, br As String
    
    
    
    htmO = "<HTML>": htmC = "</HTML>"
    hdO = "<HEAD>": hdC = "</HEAD>"
    tiO = "<TITLE>": tiC = "</TITLE>"
    bodO = "<BODY>": bodC = "</BODY>"
    f2O = "<FONT SIZE=2 FACE=""Arial"">"
    f3O = "<FONT SIZE=3 FACE=""Arial"">"
    fC = "</FONT>"
    bolO = "<B>": bolC = "</B>"
    tblO = "<TABLE WIDTH=""100%"" BORDER=0 CELLSPACING=0 CELLPADDING=0 VALIGN=""TOP"">": tblC = "</TABLE>"
    trO = "<TR VALIGN=""top"">": trC = "</TR>"
    tdc2O = "<TD WIDTH=""100%"" colspan=2><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc3O = "<TD WIDTH=""100%"" colspan=3><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdc4O = "<TD WIDTH=""100%"" colspan=4><DIV ALIGN=center><FONT SIZE=2 COLOR=""339900"" FACE=""Arial""><B>"
    tdcC = "</B></FONT></DIV></TD>"
    tdNO = "<TD WIDTH=""100%"" colspan=2><DIV align=left><FONT SIZE=2 COLOR=""#FF0000 "" FACE=""Arial"">"
    tdNC = "</FONT></DIV></TD>"
    tdOa = "<TD WIDTH=""": tdOb = "%"" VALIGN=""TOP""><FONT SIZE=2 FACE=""Arial"">": tdC = "</FONT></TD>"
    hr = "<HR>": br = "<BR>"
    
    
    strSelect = "SELECT SM.SHY56SHTP, SM.SHY56TENDT, " & _
                "SM.SHY56BEGDT, IGL_JDEDATE_TOCHAR(SM.SHY56BEGDT, 'MM/DD/YYYY') AS BEGD, SM.SHY56BEGTT, " & _
                "SM.SHY56ENDDT, IGL_JDEDATE_TOCHAR(SM.SHY56ENDDT, 'MM/DD/YYYY') AS ENDD, SM.SHY56ENDTT, " & _
                "CS.CSY56FARDT, IGL_JDEDATE_TOCHAR(CS.CSY56FARDT, 'MM/DD/YYYY') AS FRAD, CS.CSY56FARTT, " & _
                "SM.SHY56SBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56SBEDT, 'MM/DD/YYYY') AS SBED, SM.SHY56SBETT, " & _
                "SM.SHY56SENDT, IGL_JDEDATE_TOCHAR(SM.SHY56SENDT, 'MM/DD/YYYY') AS SEND, SM.SHY56SENTT, " & _
                "CS.CSY56VMVDT, IGL_JDEDATE_TOCHAR(CS.CSY56VMVDT, 'MM/DD/YYYY') AS VMVD, CS.CSY56VMVTT, " & _
                "SM.SHY56PBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56PBEDT, 'MM/DD/YYYY') AS PBED, SM.SHY56PBETT, " & _
                "SM.SHY56PENDT, IGL_JDEDATE_TOCHAR(SM.SHY56PENDT, 'MM/DD/YYYY') AS PEND, SM.SHY56PENTT, " & _
                "SM.SHY56VBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56VBEDT, 'MM/DD/YYYY') AS VBED, SM.SHY56VBETT, " & _
                "SM.SHY56VENDT, IGL_JDEDATE_TOCHAR(SM.SHY56VENDT, 'MM/DD/YYYY') AS VEND, SM.SHY56VENTT, " & _
                "SM.SHY56TBEDT, IGL_JDEDATE_TOCHAR(SM.SHY56TBEDT, 'MM/DD/YYYY') AS TBED, SM.SHY56TBETT, " & _
                "SM.SHY56TEDDT, IGL_JDEDATE_TOCHAR(SM.SHY56TEDDT, 'MM/DD/YYYY') AS TEDD, SM.SHY56TENTT, " & _
                "SM.SHY56FCCDT , SM.SHY56SMGRT, SM.SHY56DRAIT, SM.SHY56CARIT, SM.SHY56VACIT " & _
                "FROM " & F5601 & " SM, " & F5611 & " CS " & _
                "WHERE SM.SHY56SHCD = " & tmpSHCD & " " & _
                "AND SM.SHY56SHYR = " & tmpSHYR & " " & _
                "AND SM.SHY56SHCD =CS.CSY56SHCD " & _
                "AND SM.SHY56SHYR = CS.CSY56SHYR " & _
                "AND CS.CSY56CUNO = " & tmpBCC
                
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sHTML = htmO & vbNewLine
        sHTML = sHTML & hdO & tiO & tmpFBCN & " - " & tmpSHYR & " " & tmpSHNM & tiC & hdC & vbNewLine
        sHTML = sHTML & bodO & vbNewLine
        sHTML = sHTML & f3O & bolO & tmpFBCN & " - " & tmpSHYR & " " & tmpSHNM & bolC & fC & vbNewLine
        sHTML = sHTML & hr & vbNewLine
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & tdc2O & "Show Dates" & tdcC & trC & vbNewLine
        If rst.Fields("SHY56TENDT") = 1 Then
            sHTML = sHTML & trO & tdNO & bolO & "Note:  " & bolC & "All Dates are TENTATIVE" & tdNC & trC & vbNewLine
        End If
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Dates:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56BEGDT") <> 0 And rst.Fields("SHY56ENDDT") <> 0 Then 'CHANGE
            sDate1 = Format(DateValue(rst.Fields("BEGD")), "dddd mmm d")
            sDate2 = Format(DateValue(rst.Fields("ENDD")), "dddd mmm d")
            sHTML = sHTML & tdOa & "63" & tdOb & sDate1 & " - " & sDate2 & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Freight Arrive Date:" & bolC & tdC & vbNewLine
        If rst.Fields("CSY56FARDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("FRAD")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("CSY56FARTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Setup Begin Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56SBEDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("SBED")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56SBETT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Setup End Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56SENDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("SEND")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56SENTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        If UCase(Trim(rst.Fields("SHY56SHTP"))) = "S1" Then
            sHTML = sHTML & tblO & vbNewLine
            sHTML = sHTML & trO & vbNewLine
            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Vehicle Move-In:" & bolC & tdC & vbNewLine
            If rst.Fields("CSY56VMVDT") <> 0 Then
                sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("VMVD")), "dddd mmm d, yyyy") & _
                            " @ " & ConvertTime(rst.Fields("CSY56VMVTT")) & tdC & vbNewLine
            End If
            sHTML = sHTML & trC & vbNewLine
            sHTML = sHTML & tblC & vbNewLine
            
            sHTML = sHTML & br & vbNewLine
        End If
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Press Begin Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56PBEDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("PBED")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56PBETT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Press End Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56PENDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("PEND")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56PENTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Preview Begin Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56VBEDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("VBED")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56VBETT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Preview End Date:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56VENDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("VEND")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56VENTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Open to Public:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56BEGDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("BEGD")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56BEGTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Close to Public:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56ENDDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("ENDD")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56ENDTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
        sHTML = sHTML & br & vbNewLine
        
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Dismantle Begin:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56TBEDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("TBED")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56TBETT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Dismantle End:" & bolC & tdC & vbNewLine
        If rst.Fields("SHY56TEDDT") <> 0 Then
            sHTML = sHTML & tdOa & "63" & tdOb & Format(DateValue(rst.Fields("TEDD")), "dddd mmm d, yyyy") & _
                        " @ " & ConvertTime(rst.Fields("SHY56TENTT")) & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        
'''        sHTML = sHTML & br & vbNewLine
        
        
        sHTML = sHTML & hr & vbNewLine
        
        If rst.Fields("SHY56FCCDT") <> 0 Then
            sHTML = sHTML & tblO & vbNewLine
            sHTML = sHTML & trO & tdc2O & "Facility" & tdcC & trC & vbNewLine
            strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
                        "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
                        "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
                        "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
                        "WHERE AB.ABAN8 = " & rst.Fields("SHY56FCCDT") & " " & _
                        "AND AB.ABAN8 = AL.ALAN8 " & _
                        "AND AL.ALAN8 = WP.WPAN8"
            Set rstX = Conn.Execute(strSelect)
            If Not rstX.EOF Then
                sHTML = sHTML & trO & vbNewLine
                sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility:" & bolC & tdC & vbNewLine
                sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
                sHTML = sHTML & UCase(Trim(rstX.Fields("ABALPH"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD1")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD1"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD2")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD2"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD3")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD3"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD4")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD4"))) & br & vbNewLine
                If Trim(rstX.Fields("ALCTY1")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALCTY1"))) & ", " & _
                            UCase(Trim(rstX.Fields("ALADDS"))) & "  " & _
                            Trim(rstX.Fields("ALADDZ")) & br & vbNewLine
                sHTML = sHTML & tdC & vbNewLine
                sHTML = sHTML & trC & vbNewLine
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & tblO & vbNewLine
                
                Do While Not rstX.EOF
                    Select Case Trim(rstX.Fields("WPPHTP"))
                        Case ""
                            sHTML = sHTML & trO & vbNewLine
                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility Phone:" & bolC & tdC & vbNewLine
                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
                                        " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
                            sHTML = sHTML & trC & vbNewLine
                        Case "FAX"
                            sHTML = sHTML & trO & vbNewLine
                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility Fax:" & bolC & tdC & vbNewLine
                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
                                    " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
                            sHTML = sHTML & trC & vbNewLine
                    End Select
                    rstX.MoveNext
                Loop
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & hr & vbNewLine
            End If
            rstX.Close: Set rstX = Nothing
        End If
        
        
        If rst.Fields("SHY56SMGRT") <> 0 Then
            sHTML = sHTML & tblO & vbNewLine
            sHTML = sHTML & trO & tdc2O & "Show Manager" & tdcC & trC & vbNewLine
            strSelect = "SELECT AB.ABALPH, AL.ALADD1, AL.ALADD2, AL.ALADD3, AL.ALADD4, " & _
                        "AL.ALCTY1, AL.ALADDS, AL.ALADDZ, " & _
                        "WP.WPPHTP , WP.WPAR1, WP.WPPH1 " & _
                        "FROM " & F0101 & " AB, " & F0116 & " AL, " & F0115 & " WP " & _
                        "WHERE AB.ABAN8 = " & rst.Fields("SHY56SMGRT") & " " & _
                        "AND AB.ABAN8 = AL.ALAN8 " & _
                        "AND AL.ALAN8 = WP.WPAN8"
            Set rstX = Conn.Execute(strSelect)
            If Not rstX.EOF Then
                sHTML = sHTML & trO & vbNewLine
                sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Manager:" & bolC & tdC & vbNewLine
                sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
                sHTML = sHTML & UCase(Trim(rstX.Fields("ABALPH"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD1")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD1"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD2")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD2"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD3")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD3"))) & br & vbNewLine
                If Trim(rstX.Fields("ALADD4")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALADD4"))) & br & vbNewLine
                If Trim(rstX.Fields("ALCTY1")) <> "" Then _
                            sHTML = sHTML & UCase(Trim(rstX.Fields("ALCTY1"))) & ", " & _
                            UCase(Trim(rstX.Fields("ALADDS"))) & "  " & _
                            Trim(rstX.Fields("ALADDZ")) & br & vbNewLine
                sHTML = sHTML & tdC & vbNewLine
                sHTML = sHTML & trC & vbNewLine
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & tblO & vbNewLine
                
                Do While Not rstX.EOF
                    Select Case Trim(rstX.Fields("WPPHTP"))
                        Case ""
                            sHTML = sHTML & trO & vbNewLine
                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Mgr Phone:" & bolC & tdC & vbNewLine
                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
                                        " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
                            sHTML = sHTML & trC & vbNewLine
                        Case "FAX"
                            sHTML = sHTML & trO & vbNewLine
                            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Mgr Fax:" & bolC & tdC & vbNewLine
                            sHTML = sHTML & tdOa & "63" & tdOb & Trim(rstX.Fields("WPAR1")) & _
                                    " " & Trim(rstX.Fields("WPPH1")) & tdC & vbNewLine
                            sHTML = sHTML & trC & vbNewLine
                    End Select
                    rstX.MoveNext
                Loop
                sHTML = sHTML & tblC & vbNewLine
                sHTML = sHTML & hr & vbNewLine
            End If
            rstX.Close: Set rstX = Nothing
        End If
    End If
    rst.Close: Set rst = Nothing
        
    
    '///// GET SHOW REG ABSTRACT DATA \\\\\
    strSelect = "SELECT CH.HALLID, HM.AN8_FCCD, CH.AN8_SHCD, SU.ABALPH, " & _
                "HM.HALLDESC, HM.CLGHGT, HM.CLGUNIT, HM.CLGNOTE, HM.HALLNOTE, " & _
                "SHR.HGTRES, SHR.RESUNIT, SHR.RESNOTE, SHR.SHOWNOTE, " & _
                "EA.EASENAME , EA.EASEVAL, EA.EASEUNIT, EA.EASEDESC " & _
                "FROM IGLPROD.SRA_CLIENTHALL CH, IGLPROD.SRA_HALLMASTER HM, " & _
                "IGLPROD.SRA_EASEMENT EA, " & _
                "" & F0101 & " SU, IGLPROD.SRA_SHOWHALLRESTRICTION SHR " & _
                "WHERE CH.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND CH.SHYR = " & tmpSHYR & " " & _
                "AND CH.AN8_SHCD = " & tmpSHCD & " " & _
                "AND CH.HALLID = HM.HALLID " & _
                "AND HM.AN8_FCCD = SU.ABAN8 " & _
                "AND CH.HALLID =SHR.HALLID " & _
                "AND CH.AN8_SHCD = SHR.AN8_SHCD " & _
                "AND CH.HALLID = EA.HALLID " & _
                "AND CH.AN8_SHCD = EA.AN8_SHCD"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & tdc2O & "Hall Information from Show Regulation Abstract" & tdcC & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Facility:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & UCase(Trim(rst.Fields("ABALPH"))) & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Hall:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & UCase(Trim(rst.Fields("HALLDESC"))) & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Hall Ceiling Hgt:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & ConvertDims(CDbl(rst.Fields("CLGHGT")), rst.Fields("CLGUNIT")) & _
                    tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Show Restriction:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & ConvertDims(CDbl(rst.Fields("HGTRES")), rst.Fields("RESUNIT")) & _
                    tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        If Trim(rst.Fields("EASENAME")) <> "" Then
            sHTML = sHTML & trO & vbNewLine
            sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Easements:" & bolC & tdC & vbNewLine
            sHTML = sHTML & tdOa & "63" & tdOb & vbNewLine
            Do While Not rst.EOF
                sHTML = sHTML & UCase(Trim(rst.Fields("EASENAME"))) & "  (" & _
                            ConvertDims(CDbl(rst.Fields("EASEVAL")), rst.Fields("EASEUNIT")) & ")" & _
                            br & vbNewLine
                rst.MoveNext
            Loop
            sHTML = sHTML & tdC & vbNewLine
            sHTML = sHTML & trC & vbNewLine
        End If
        sHTML = sHTML & tblC & vbNewLine
        sHTML = sHTML & hr & vbNewLine
    End If
    rst.Close: Set rst = Nothing
    
    
    
    '///// GET FLOORPLAN STATUS \\\\\
    Dim sSQFT As String
    strSelect = "SELECT FP.FPSTATUS, FP.FPSTATBY, " & _
                "TO_CHAR(FP.FPSTATDT, 'MM/DD/YYYY') AS FPSTATDT, R.VAL_DESC " & _
                "FROM AQUA.AQUA_FLORRPLAN_STATUS FP, IGLPROD.IGL_REF R " & _
                "WHERE FP.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND FP.AN8_SHCD = " & tmpSHCD & " " & _
                "AND FP.SHYR = " & tmpSHYR & " " & _
                "AND FP.FPSTATUS = R.REF_ID " & _
                "AND R.TYPE_CD = 12"
    strSelect = "SELECT DM.DSTATUS, DM.UPDUSER, " & _
                "TO_CHAR(DM.UPDDTTM, 'MON DD, YYYY') AS FPSTATDT, R.VAL_DESC, " & _
                "CS.CSY56BOOTT BONO, CS.CSY56SQFTT AS SQFT, CS.CSY56BOTPT AS BOPH " & _
                "FROM ANNOTATOR.DWG_MASTER DM, ANNOTATOR.DWG_SHOW DS, IGLPROD.IGL_REF R, " & _
                "" & F5611 & " CS " & _
                "WHERE DS.AN8_CUNO = " & CLng(tmpBCC) & " " & _
                "AND DS.AN8_SHCD = " & tmpSHCD & " " & _
                "AND DS.SHYR = " & tmpSHYR & " " & _
                "AND DS.DWGID = DM.DWGID " & _
                "AND DM.DWGTYPE = 0 " & _
                "AND DM.DSTATUS = R.REF_ID " & _
                "AND R.TYPE_CD = 12 " & _
                "AND DS.AN8_CUNO = CS.CSY56CUNO " & _
                "AND DS.AN8_SHCD = CS.CSY56SHCD " & _
                "AND DS.SHYR = CS.CSY56SHYR"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        sHTML = sHTML & tblO & vbNewLine
        sHTML = sHTML & trO & tdc2O & "Space Information" & tdcC & trC & vbNewLine
        
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Space Number:" & bolC & tdC & vbNewLine
        If Not IsNull(rst.Fields("BONO")) Then _
            sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("BONO") & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Floor Space Area:" & bolC & tdC & vbNewLine
        If Not IsNull(rst.Fields("SQFT")) Then
            sSQFT = Format(rst.Fields("SQFT") / 1000, "#,##0")
            sHTML = sHTML & tdOa & "63" & tdOb & sSQFT & tdC & vbNewLine
        End If
        sHTML = sHTML & trC & vbNewLine
        
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Booth Phone No:" & bolC & tdC & vbNewLine
        If Not IsNull(rst.Fields("BOPH")) Then _
            sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("BOPH") & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Is Floorplan Req'd?:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & "YES" & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Floorplan Status:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & rst.Fields("DSTATUS") & _
                    " - " & UCase(Trim(rst.Fields("VAL_DESC"))) & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & trO & vbNewLine
        sHTML = sHTML & tdOa & "37" & tdOb & bolO & "Status By:" & bolC & tdC & vbNewLine
        sHTML = sHTML & tdOa & "63" & tdOb & "LAST EDIT ON " & UCase(Trim(rst.Fields("FPSTATDT"))) & _
                    " BY " & UCase(Trim(rst.Fields("UPDUSER"))) & tdC & vbNewLine
        sHTML = sHTML & trC & vbNewLine
        sHTML = sHTML & tblC & vbNewLine
        sHTML = sHTML & hr & vbNewLine
    End If
    rst.Close
    Set rst = Nothing
        
    
    sHTML = sHTML & bodC & vbNewLine
    sHTML = sHTML & htmC
    
    tFile1 = strHTMLPath & sIFile
    Open tFile1 For Output As #1
    Print #1, sHTML
    Close #1
    
    PopShowInfo = tFile1
End Function

Public Function ConvertTime(iTime As Long) As String
    Dim iHour As Long, iMin As Long
    Dim sAMPM As String
    Dim sTime As String, sHour As String, sMin As String
    
    
    Select Case iTime
        Case 0
            ConvertTime = "12:00 AM"
        Case Is < 120000
            sTime = Right("000000" & CStr(iTime), 6)
            sHour = Left(sTime, 2)
            sMin = Mid(sTime, 3, 2)
            If CInt(sHour) = 0 Then
                ConvertTime = "12:" & sMin & " AM"
            Else
                ConvertTime = CStr(CInt(sHour)) & ":" & sMin & " AM"
            End If
        Case 120000
            ConvertTime = "12:00 NOON"
        Case 240000
            ConvertTime = "12:00 MID"
        Case Else
            iTime = iTime - 120000
            sTime = Right("000000" & CStr(iTime), 6)
            sHour = Left(sTime, 2)
            sMin = Mid(sTime, 3, 2)
            If CInt(sHour) = 0 Then
                ConvertTime = "12:" & sMin & " PM"
            Else
                ConvertTime = CStr(CInt(sHour)) & ":" & sMin & " PM"
            End If
    End Select
    
'''    iMin = iTime Mod (iHour * 10000)
'''    If iHour = 0 Then iHour = 12
'''    Select Case iMin
'''        Case 0
'''            ConvertTime = iHour & ":00" & sAMPM
'''        Case Else
'''            ConvertTime = iHour & ":" & Right("00" & CStr(iMin / 100), 2) & sAMPM
'''    End Select
'''GotIt:
End Function

Public Function ConvertDims(Num As Double, iUnit As Integer) As String
    Dim Feet As Integer, Inch As Integer, Numer As Integer
    Dim Frac As Currency
    Dim strFrac As String
    Select Case iUnit
        Case 1
            Feet = Int(Num / 12)
            Inch = Int(Num - (Feet * 12))
            Frac = CCur((((Num / 12) - Feet) _
                    * 12) - Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case 2
            Feet = Int(Num)
            Inch = (Num - Feet) * 12
            Frac = Inch - Int(Inch)
            If Frac > 0 Then
                Numer = CInt(Frac * 8)
                Select Case Numer
                    Case 1
                        strFrac = " 1/8"""
                    Case 2
                        strFrac = " 1/4"""
                    Case 3
                        strFrac = " 3/8"""
                    Case 4
                        strFrac = " 1/2"""
                    Case 5
                        strFrac = " 5/8"""
                    Case 6
                        strFrac = " 3/4"""
                    Case 7
                        strFrac = " 7/8"""
                    Case Else
                        strFrac = Chr(34)
                End Select
        
            Else
                strFrac = Chr(34)
            End If
            ConvertDims = Feet & "'-" & Inch & strFrac
        Case Else
            ConvertDims = "Soon!"
    End Select
End Function

'''''Public Sub SetLotusVars(tUser As String, tForm As Form)
'''''    Dim myNotes As New Domino.NotesSession '' Object '' As NotesSession
'''''    Dim pdb As New Domino.NotesDatabase '' Object ''  NotesDatabase
'''''
'''''    Dim uidoc As Object ''  NotesUIDocument
'''''    Dim pview As Object ''  NotesView
'''''    Dim pdoc As Object ''  NotesDocument
'''''    Dim sdoc As Object ''  NotesDocument
'''''    Dim retGet As Variant
'''''    Dim strHomeServer As Variant, strFile As Variant  '' As String
'''''    Dim intLen As Integer, intSpot As Integer
'''''
'''''
'''''    On Error Resume Next
'''''    If sNOTESID = "GANNOTAT" Then
'''''        strMailSrvr = "detsrv1/det/GPJNotes"
'''''        strMailFile = "mail\gannotat.nsf"
'''''    Else
'''''        If sNOTESPASSWORD = "" Then
'''''            ''GET PASSWORD''
'''''TryPWAgain:
'''''            frmGetPassword.Show 1 '', tForm
'''''            Select Case sNOTESPASSWORD
'''''                Case "_CANCEL"
'''''                    sNOTESPASSWORD = ""
'''''                    MsgBox "No email will be sent", vbExclamation, "User Canceled..."
'''''                    Set myNotes = Nothing
'''''                    Set pdb = Nothing
'''''                Case Else
'''''                    Err.Clear
'''''                    myNotes.Initialize (sNOTESPASSWORD)
'''''                    If Err Then
'''''                        Err.Clear
'''''                        GoTo TryPWAgain
'''''                    End If
'''''            End Select
'''''        Else
'''''            myNotes.Initialize (sNOTESPASSWORD)
'''''        End If
'''''    End If
'''''
'''''    On Error GoTo 0
'''''    Set pdb = myNotes.GetDatabase(LKUP_SRVR, NAB_NAME)
'''''
'''''    Set pview = pdb.GetView("($Users)")
'''''    Set pdoc = pview.GetDocumentByKey(tUser)
'''''
'''''    If (pdoc Is Nothing) Then
'''''        MsgBox "Update unsuccessful, update document not found.  ", vbCritical, "Lookup Failed"
'''''        Exit Sub
'''''    Else
'''''        strHomeServer = pdoc.GetItemValue("MailServer")
'''''        intLen = Len(strHomeServer(0))
'''''        intSpot = InStr(3, strHomeServer(0), "/")
'''''        strMailSrvr = Mid(strHomeServer(0), 4, intSpot - 3)
'''''        intSpot = InStr(1, strHomeServer(0), "OU=") + 3
'''''        strMailSrvr = strMailSrvr & Mid(strHomeServer(0), intSpot, 3)
'''''        strMailSrvr = strMailSrvr & "/GPJNotes"
'''''
'''''        strFile = pdoc.GetItemValue("MailFile")
'''''        strMailFile = strFile(0) & ".nsf"
'''''
''''''        MsgBox "MailServer: " & vbTab & strMailSrvr & vbNewLine & _
''''''                    "MailFile:     " & vbTab & strMailFile, vbInformation, "Lookup Complete"
'''''
'''''    End If
'''''End Sub



Public Function GetAnoSeq() As Long
    Dim rstL As ADODB.Recordset
    
    Set rstL = Conn.Execute("SELECT " & ANOSeq & ".NEXTVAL FROM DUAL")
    GetAnoSeq = rstL.Fields("nextval")
    rstL.Close: Set rstL = Nothing
    
End Function

Public Function CheckForThumbPath(pGID As Long, pFormat As String, pVer As Integer) As String
    Dim sGPath As String, sVPath As String, sFile As String
    
    
    sGPath = "\\DETMSFS01\GPJAnnotator\Graphics\"
    sVPath = "\\DETMSFS01\GPJAnnotator\Graphics\Versions\"
    
    If UCase(pFormat) = "PDF" Then
        ''LOOK FOR BMP HERE''
        sFile = sGPath & "pdf_" & pGID & ".bmp"
        If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.BMP''
            ''PDF.BMP NOT FOUND''
            sFile = sGPath & "pdf_" & pGID & ".jpg"
            If Dir(sFile, vbNormal) = "" Then ''CHECK FOR PDF.JPG''
                ''NO THUMBNAIL AT ALL''
                sFile = sGPath & "pdf.bmp"
            End If
        End If
    Else
        sFile = sGPath & "Thumbs\thb_" & pGID & ".jpg"
        If Dir(sFile, vbNormal) = "" Then ''OPEN FULL FILE''
            sFile = sGPath & pGID & "." & pFormat
        End If
        
    End If
    
    CheckForThumbPath = sFile
End Function

Public Sub CleanUpAnnoLog()
    Dim strDelete As String
    
    On Error Resume Next
    
    Conn.BeginTrans
    strDelete = "DELETE FROM ANNOTATOR.ANO_LOCKLOG WHERE ADDDTTM < SYSDATE - 365"
    Conn.Execute (strDelete)
    If Err = 0 Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        MsgBox "Error while attempting to clean up ANO_LOCKLOG table", _
                    vbExclamation, "Just letting you know..."
    End If
End Sub

Public Function Legalize(sName As String) As String
    sName = Replace(sName, "\", "-")
    sName = Replace(sName, "/", "-")
    sName = Replace(sName, ":", "-")
    sName = Replace(sName, ";", "-")
    sName = Replace(sName, "*", "")
    sName = Replace(sName, """", "'")
    sName = Replace(sName, "|", "-")
    sName = Replace(sName, "<", "")
    sName = Replace(sName, ">", "")
    sName = Replace(sName, "?", "")
    sName = Replace(sName, ".", "")
    
    Legalize = sName
End Function
