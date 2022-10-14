Attribute VB_Name = "basDL"
Option Explicit

Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
