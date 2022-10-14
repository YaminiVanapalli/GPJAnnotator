Attribute VB_Name = "Module1"
Option Explicit
'''' Copyright © 1997 by Desaware Inc. All Rights Reserved
'''
'''#If Win32 Then
'''Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'''Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Byte) As Long
'''Declare Function VerLanguageName Lib "version.dll" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
'''Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Byte, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
'''Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
'''Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'''Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'''Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpname As String, ByVal cbName As Long) As Long
'''Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpname As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
'''
'''Declare Function RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal _
'''hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName _
'''As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long)
'''
'''#Else
'''Declare Function GetFileVersionInfo% Lib "ver.dll" (ByVal lpszFileName$, ByVal handle&, ByVal cbBuf&, lpvData As Byte)
'''Declare Function GetFileVersionInfoSize% Lib "ver.dll" (ByVal lpszFileName$, lpdwHandle&)
'''Declare Function VerLanguageName% Lib "ver.dll" (ByVal Lang%, ByVal lpszLang$, ByVal cbLang%)
'''Declare Function VerQueryValue% Lib "ver.dll" (lpvBlock As Byte, ByVal SubBlock$, lpBuffer&, lpcb%)
'''Declare Function GetProfileString% Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName As String, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%)
'''
'''
'''
'''#End If
'''
''''----------------------------------------------
''''
''''   Public Constants
'''
'''Public Const VFT_UNKNOWN = &H0&
'''Public Const VFT_APP = &H1&
'''Public Const VFT_DLL = &H2&
'''Public Const VFT_DRV = &H3&
'''Public Const VFT_FONT = &H4&
'''Public Const VFT_VXD = &H5&
'''Public Const VFT_STATIC_LIB = &H7&
'''
'''Public Type VS_FIXEDFILEINFO
'''        dwSignature As Long
'''        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
'''        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
'''        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
'''        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
'''        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
'''        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
'''        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
'''        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
'''        dwFileType As Long             '  e.g. VFT_DRIVER
'''        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
'''        dwFileDateMS As Long           '  e.g. 0
'''        dwFileDateLS As Long           '  e.g. 0
'''End Type
'''
'''Public Type FILETIME
'''        dwLowDateTime As Long
'''        dwHighDateTime As Long
'''End Type
'''
'''Public Const HKEY_CLASSES_ROOT = &H80000000
'''Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
'''Public Const HKEY_USERS = &H80000003
'''Public Const HKEY_PERFORMANCE_DATA = &H80000004
'''
Public Const SYNCHRONIZE = &H100000
'''Public Const STANDARD_RIGHTS_READ = &H20000
'''Public Const STANDARD_RIGHTS_WRITE = &H20000
'''Public Const STANDARD_RIGHTS_EXECUTE = &H20000
'''Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
'''Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
'''Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'''Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'''Public Const ERROR_SUCCESS = 0&
'''
''''-------------------------------------------------
''''
''''   Public Variables
''''
'''' We changed this to Byte to prevent the string
'''' mangling of the buffer
'''Public verbuf() As Byte      ' Version buffer
'''Public FileName$    ' Current file to examine
'''
'''
