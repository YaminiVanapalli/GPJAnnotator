VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'''    Dim Conn As Object '''ADODB.Connection
'''    Dim ConnStr As String, strSelect As String
'''    Dim rst As ADODB.Recordset
'''
'''
'''    MsgBox "About to create ADO Object"
'''    Set Conn = New ADODB.Connection
'''    MsgBox "ADO Object created"
'''
'''    ConnStr = "DSN=JDE;UID=ANNOTATOR;PWD=ANNOTATOR"
'''
'''    MsgBox "About to try connecting (" & ConnStr & ")"
'''
'''    Conn.Open (ConnStr)
'''
'''    MsgBox "Connection is open"
'''
'''    strSelect = "SELECT * FROM IGL_REF"
'''    Set rst = Conn.Execute(strSelect)
'''    If Not rst.EOF Then
'''        MsgBox "Successful Recordset"
'''    Else
'''        MsgBox "UNSuccessful Recordset"
'''    End If
'''    rst.Close: Set rst = Nothing
    
    
    
    Dim X As Printer
    Dim sDef As String
    Dim i As Integer
    
    sDef = Printer.DeviceName
    MsgBox "Default = " & sDef
    
    i = 0
    For Each X In Printers
        MsgBox X.DeviceName
        i = i + 1
'''        If X.DeviceName = sDef Then
'''            Printer.DeviceName = sDef
'''        End If
    Next
    
    MsgBox i & " Printers Found"
End Sub
