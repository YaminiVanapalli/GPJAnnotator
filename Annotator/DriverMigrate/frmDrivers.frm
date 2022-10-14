VERSION 5.00
Begin VB.Form frmDrivers 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   5715
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdGetDrivers 
      Caption         =   "Get Drivers"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   300
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "frmDrivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conn As New ADODB.Connection


Private Sub cmdGetDrivers_Click()
    Dim sFile As String, sRec As String, sDrv As String, sSub As String
    Dim i1 As Integer, i2 As Integer
    Dim strSelect As String, strInsert As String
    Dim rst As ADODB.Recordset
    
    Conn.BeginTrans
    sFile = App.Path & "\Wtsuprn.txt"
    Open sFile For Input As #1
    Do While Not EOF(1)
        Input #1, sRec
        If Left(sRec, 1) <> ";" Then
'''            i1 = 2
'''            i2 = InStr(i1, sRec, """ = ")
            sDrv = sRec
            List1.AddItem sDrv
'''            MsgBox "Original Driver:  " & sDrv

            '///// LOOK FOR ORIGINAL DRIVER \\\\\
            strSelect = "SELECT * FROM ANO_PRINTER_DRIVER " & _
                        "WHERE DRIVER_NAME = '" & UCase(sDrv) & "'"
            Set rst = Conn.Execute(strSelect)
            If rst.EOF Then
                strInsert = "INSERT INTO ANO_PRINTER_DRIVER " & _
                            "(DRIVER_NAME, DRIVER_STATUS, ADDDTTM, " & _
                            "UPDUSER, UPDDTTM, UPDCNT) " & _
                            "VALUES " & _
                            "('" & UCase(sDrv) & "', 2, SYSDATE, " & _
                            "'SAW - MIGRATION', SYSDATE, 1)"
                Conn.Execute (strInsert)
            End If
            rst.Close: Set rst = Nothing
            
            Input #1, sRec
            sDrv = Mid(sRec, 4, Len(sRec) - 4)
            List2.AddItem sDrv
'''            MsgBox "Substitute:  " & sDrv
            
            '///// LOOK FOR SUBSTITUTE DRIVER \\\\\
            strSelect = "SELECT * FROM ANO_PRINTER_DRIVER " & _
                        "WHERE DRIVER_NAME = '" & UCase(sDrv) & "'"
            Set rst = Conn.Execute(strSelect)
            If rst.EOF Then
                strInsert = "INSERT INTO ANO_PRINTER_DRIVER " & _
                            "(DRIVER_NAME, DRIVER_STATUS, ADDDTTM, " & _
                            "UPDUSER, UPDDTTM, UPDCNT) " & _
                            "VALUES " & _
                            "('" & UCase(sDrv) & "', 2, SYSDATE, " & _
                            "'SAW - MIGRATION', SYSDATE, 1)"
                Conn.Execute (strInsert)
            End If
            rst.Close: Set rst = Nothing
        End If
    Loop
    Close #1
    Conn.CommitTrans
End Sub

Private Sub Form_Load()
    Dim ConnStr As String
    
'''    Set Conn = New ADODB.Connection
    ConnStr = "DSN=JDE;UID=ANNOTATOR;PWD=ANNOTATOR"
    Conn.Open (ConnStr)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub
