VERSION 5.00
Object = "{51C0A9CA-F7B7-4F5A-96F4-43927C6FA50F}#1.0#0"; "MapPointControl.ocx"
Begin VB.Form frmMap 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
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
   ScaleHeight     =   8985
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   435
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   2235
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMap.frx":0000
      Left            =   240
      List            =   "frmMap.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   4575
   End
   Begin MapPointCtl.MappointControl map1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      BorderStyle     =   0
      MousePointer    =   0
      Object.TabStop         =   0   'False
      Appearance      =   1
      PaneState       =   3
      UnitsOfMeasure  =   0
   End
   Begin VB.Label lblFacility 
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
      Height          =   195
      Left            =   4980
      TabIndex        =   3
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   45
   End
   Begin VB.Label lblShow 
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
      Height          =   240
      Left            =   4980
      TabIndex        =   2
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   60
   End
   Begin VB.Menu mnuMap 
      Caption         =   "mnuMap"
      Visible         =   0   'False
      Begin VB.Menu mnuMapPrint 
         Caption         =   "Print Map"
      End
      Begin VB.Menu mnuMapDirs 
         Caption         =   "Print Directions"
      End
      Begin VB.Menu mnuMapType 
         Caption         =   "Set Map Style"
         Begin VB.Menu mnuMapStyle 
            Caption         =   "Road Map"
            Index           =   0
         End
         Begin VB.Menu mnuMapStyle 
            Caption         =   "Road and data map"
            Index           =   1
         End
         Begin VB.Menu mnuMapStyle 
            Caption         =   "Data map"
            Index           =   2
         End
         Begin VB.Menu mnuMapStyle 
            Caption         =   "Terrain map"
            Index           =   3
         End
         Begin VB.Menu mnuMapStyle 
            Caption         =   "Political map"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeather 
         Caption         =   "Weather..."
      End
      Begin VB.Menu mnuCancel02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New ADODB.Connection
Dim lFCCD As Long
Dim tState As String, sLoc As String


Private Sub cmdOptions_Click()
    Me.PopupMenu mnuMap, 0, cmdOptions.Left, cmdOptions.Top + cmdOptions.Height
End Sub

Private Sub Combo1_Click()
    Dim strAddress As String
    Dim strSelect As String, sMess As String
    Dim rst As ADODB.Recordset
    Dim i As Integer
    
    If Combo1.Text <> "" Then
        lFCCD = Combo1.ItemData(Combo1.ListIndex)
        sMess = ""
        strSelect = "SELECT ALADD1, ALADD2, ALADD3, ALADD4, " & _
                    "ALCTY1, ALADDS, ALADDZ " & _
                    "FROM proddta.F0116 " & _
                    "WHERE ALAN8 = " & lFCCD & " " & _
                    "AND ALEFTB IN " & _
                    "(SELECT MAX(ALEFTB) " & _
                    "FROM PRODDTA.F0116 " & _
                    "WHERE ALAN8 = " & lFCCD & ")"
        Set rst = Conn.Execute(strSelect)
        If Not rst.EOF Then
            For i = 0 To 3
                If Not IsNull(rst.Fields(i)) And Trim(rst.Fields(i)) <> "" Then
                    sMess = sMess & UCase(Trim(rst.Fields(i))) & ", "
                End If
            Next i
            strAddress = sMess & UCase(Trim(rst.Fields("ALCTY1"))) & ", " & _
                        UCase(Trim(rst.Fields("ALADDS"))) & ", " & _
                        Trim(rst.Fields("ALADDZ"))
            tState = LCase(Trim(rst.Fields("ALADDS")))
            sLoc = UCase(Trim(rst.Fields("ALCTY1"))) & ", " & UCase(Trim(rst.Fields("ALADDS")))
        End If
        rst.Close: Set rst = Nothing
'''        MsgBox strAddress
    End If
            
            
'''    Select Case Combo1.Text
'''        Case "Alan"
'''            strAddress = "15015 Dunn Dr., Traverse City, MI"
'''        Case "Sue"
'''            strAddress = "1325 Crestview, Alpena, MI, 48707"
'''        Case "Mom"
'''            strAddress = "2421 Arlington Rd., Lansing, MI, 48906"
'''        Case "Home"
'''            strAddress = "37059 Bradford, Sterling Heights, MI, 48312"
'''        Case "Donna in Lansing"
'''            strAddress = "7107 Willow Wood Circle, Lansing, MI, 48917"
'''        Case "Donna in TC"
'''            strAddress = "11172 Peninsula Dr., Traverse City, MI, 49686"
'''    End Select
    On Error GoTo ErrorTrap
    map1.Visible = False
    map1.PaneState = geoPaneNone
    map1.ActiveMap.ShowFindDialog strAddress, geoFindAddress
    map1.Visible = True
    map1.ActiveMap.AllowEdgePan = True
    lblShow = Combo1.Text
    strSelect = "SELECT ABALPH FROM PRODDTA.F0101 " & _
                "WHERE ABAN8 = " & lFCCD
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        lblFacility = sLoc & "  -  " & UCase(Trim(rst.Fields("ABALPH")))
    Else
        lblFacility = ""
    End If
    rst.Close: Set rst = Nothing
Exit Sub
ErrorTrap:
    map1.Visible = False
End Sub

Private Sub Combo2_Click()
    map1.ActiveMap.MapStyle = Combo2.ItemData(Combo2.ListIndex)
End Sub

'''''Private Sub Command1_Click()
'''''    FindAddressSearch
'''''End Sub

Private Sub Form_Load()
    Dim ConnStr As String
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    
    ConnStr = "DSN=JDE;UID=ANNOTATOR;PWD=ANNOTATOR"
    Conn.Open (ConnStr)
    
    Combo1.Clear
    strSelect = "SELECT SHY56NAMA, SHY56FCCDT " & _
                "FROM PRODDTA.F5601 " & _
                "WHERE SHY56SHYR = 2001 " & _
                "AND SHY56FCCDT > 0 " & _
                "ORDER BY UPPER(SHY56NAMA)"
    Set rst = Conn.Execute(strSelect)
    Do While Not rst.EOF
        Combo1.AddItem UCase(Trim(rst.Fields("SHY56NAMA")))
        Combo1.ItemData(Combo1.NewIndex) = rst.Fields("SHY56FCCDT")
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    
'''    map1.OpenMap "D:\Program Files\Microsoft MapPoint 2002 Trial\Samples\SampleTemplate.ptt"
    map1.NewMap geoMapNorthAmerica
'''    map1.ActiveMap.FindAddressResults "37059 Bradford", "Sterling Heights, MI",
    
'''    Combo2.AddItem "Road map"
'''    Combo2.ItemData(Combo2.NewIndex) = 0
'''    Combo2.AddItem "Road and data map"
'''    Combo2.ItemData(Combo2.NewIndex) = 1
'''    Combo2.AddItem "Data map"
'''    Combo2.ItemData(Combo2.NewIndex) = 2
'''    Combo2.AddItem "Terrain map"
'''    Combo2.ItemData(Combo2.NewIndex) = 3
'''    Combo2.AddItem "Political map"
'''    Combo2.ItemData(Combo2.NewIndex) = 4
    
End Sub


'''''Sub FindAddressSearch()
''''''''    Dim objApp As New MapPoint.Application
''''''''    Dim objFindResults As MapPoint.FindResults
''''''''
''''''''    'Set up application
''''''''    objApp.Visible = True
''''''''    objApp.UserControl = True
'''''
'''''    'Output first result of find search
'''''    map1.ActiveMap.ShowFindDialog "37059 Bradford, Sterling Heights, MI", geoFindAddress
'''''    map1.Visible = True
'''''    map1.ActiveMap.AllowEdgePan = True
''''''''    map1.ActiveMap.AddPushpin map1.ActiveMap.FindPlaceResults("37059 Bradford, Sterling Heights, MI")
''''''''    MsgBox "The first item in the find list is: " _
''''''''      + objFindResults.Item(1).Name
'''''
'''''End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
        map1.Width = Me.ScaleWidth - 480
        map1.Height = Me.ScaleHeight - map1.Top - 240
        cmdOptions.Left = map1.Left + map1.Width - cmdOptions.Width
        
'''        picCon.Top = map1.Top + 120
'''        picCon.Left = map1.Left + map1.Width - 120 - picCon.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    map1.CloseMap
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub mnuMapStyle_Click(Index As Integer)
    Dim i As Integer
    map1.ActiveMap.MapStyle = Index
    For i = 0 To 4
        If i = Index Then mnuMapStyle(i).Checked = True Else mnuMapStyle(i).Checked = False
    Next i
End Sub

Private Sub mnuWeather_Click()
    With frmWeather
        .PassSTATE = tState
        .Show 1
    End With
End Sub
