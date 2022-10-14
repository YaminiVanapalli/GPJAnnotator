VERSION 5.00
Begin VB.Form frmRegTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmRegTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim hKey As Long, e As Long
    Dim v As Variant
    Dim sPrinter As String, sDriver As String
    
    sPrinter = Printer.DeviceName
    sPrinter = Replace(sPrinter, "\", ",")
    e = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                "SYSTEM\CurrentControlSet\Control\Print\Printers\" & sPrinter, _
                0&, KEY_ALL_ACCESS, hKey)
    If e Then Exit Sub
    e = GetRegValue(hKey, "Printer Driver", v)
    If e Then Exit Sub
    sDriver = v
    e = RegCloseKey(hKey)
    MsgBox "Printer:" & vbTab & Printer.DeviceName & _
                vbNewLine & _
                "Driver:" & vbTab & sDriver, vbInformation, "Printer Driver..."
End Sub

