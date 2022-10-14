VERSION 5.00
Begin VB.Form frmElemInfo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmElemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   5820
      TabIndex        =   74
      Top             =   5640
      Width           =   1995
   End
   Begin VB.PictureBox picPhoto 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6000
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   67
      ToolTipText     =   "Double-Click to Enlarge"
      Top             =   300
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   2820
      Picture         =   "frmElemInfo.frx":08CA
      ScaleHeight     =   2115
      ScaleWidth      =   4875
      TabIndex        =   66
      Top             =   6540
      Width           =   4935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Changes"
      Height          =   435
      Left            =   3000
      TabIndex        =   62
      Top             =   5640
      Width           =   2715
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear && Reset to Current Save"
      Height          =   435
      Left            =   180
      TabIndex        =   61
      Top             =   5640
      Width           =   2715
   End
   Begin VB.Frame fraElemInfo 
      Enabled         =   0   'False
      Height          =   5355
      Left            =   180
      TabIndex        =   32
      Top             =   60
      Width           =   7635
      Begin VB.Frame Frame8 
         Caption         =   "Heights"
         Height          =   2055
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   3255
         Begin VB.TextBox txtPrimeHgt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtHgtA 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   780
            Width           =   1455
         End
         Begin VB.TextBox txtHgtB 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtHgtC 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   1380
            Width           =   1455
         End
         Begin VB.TextBox txtAssemHgt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton optHgtUnit 
            Caption         =   "Decimal Inches"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optHgtUnit 
            Caption         =   "Centimeters"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Height:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opt'l Height - 'A':"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opt'l Height - 'B':"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opt'l Height - 'C':"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Assembled Height:"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1695
         Left            =   60
         TabIndex        =   63
         Top             =   120
         Width           =   5955
         Begin VB.Frame Frame7 
            Caption         =   "Crated Weight"
            Height          =   795
            Left            =   3480
            TabIndex        =   69
            Top             =   840
            Width           =   2175
            Begin VB.TextBox txtEstWgt 
               Alignment       =   2  'Center
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   120
               TabIndex        =   72
               Text            =   "0"
               Top             =   240
               Width           =   915
            End
            Begin VB.OptionButton optWgtUnit 
               Caption         =   "lbs"
               Height          =   195
               Index           =   0
               Left            =   1080
               TabIndex        =   71
               Top             =   300
               Value           =   -1  'True
               Width           =   555
            End
            Begin VB.OptionButton optWgtUnit 
               Caption         =   "Kg"
               Height          =   195
               Index           =   1
               Left            =   1620
               TabIndex        =   70
               Top             =   300
               Width           =   495
            End
            Begin VB.Label lblDisc 
               BackStyle       =   0  'Transparent
               Caption         =   "* Unweighed Parts found"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   540
               Visible         =   0   'False
               Width           =   1875
               WordWrap        =   -1  'True
            End
         End
         Begin VB.CheckBox chkContainer 
            Caption         =   "Container Element Only"
            Height          =   255
            Left            =   3480
            TabIndex        =   65
            Top             =   540
            Width           =   2235
         End
         Begin VB.TextBox txtDesc 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1020
            MaxLength       =   30
            TabIndex        =   0
            Top             =   120
            Width           =   3495
         End
         Begin VB.TextBox txtCPRJ 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   135
            Width           =   855
         End
      End
      Begin VB.ComboBox cboSealReq 
         Height          =   315
         ItemData        =   "frmElemInfo.frx":99A48
         Left            =   2160
         List            =   "frmElemInfo.frx":99A4A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4500
         Width           =   1215
      End
      Begin VB.ComboBox cboCertLA 
         Height          =   315
         ItemData        =   "frmElemInfo.frx":99A4C
         Left            =   2160
         List            =   "frmElemInfo.frx":99A4E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ComboBox cboClgHung 
         Height          =   315
         ItemData        =   "frmElemInfo.frx":99A50
         Left            =   2160
         List            =   "frmElemInfo.frx":99A52
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtCanopyArea 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Text            =   "0"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtDwgNo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   14
         Top             =   4920
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox txtQtyPerElem 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2820
         TabIndex        =   9
         Text            =   "1"
         Top             =   2820
         Width           =   555
      End
      Begin VB.Frame Frame9 
         Caption         =   "Electrical Requirements"
         Height          =   2805
         Left            =   3540
         TabIndex        =   33
         Top             =   2400
         Width           =   3975
         Begin VB.TextBox txtECInit 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   17
            Top             =   210
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt500w 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Text            =   "0"
            Top             =   615
            Width           =   495
         End
         Begin VB.TextBox txt1000w 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   20
            Text            =   "0"
            Top             =   615
            Width           =   495
         End
         Begin VB.TextBox txt1500w 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            Text            =   "0"
            Top             =   615
            Width           =   495
         End
         Begin VB.TextBox txt2000w 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Text            =   "0"
            Top             =   615
            Width           =   495
         End
         Begin VB.TextBox txtPhones 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   23
            Text            =   "0"
            Top             =   615
            Width           =   615
         End
         Begin VB.TextBox txtWater 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   24
            Text            =   "0"
            Top             =   615
            Width           =   615
         End
         Begin VB.CheckBox cbxElecConf 
            Caption         =   "Elec Specs Confirmed"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   225
            Width           =   2115
         End
         Begin VB.TextBox txtECDate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   210
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame10 
            Caption         =   "220v Electrical"
            Height          =   1395
            Left            =   120
            TabIndex        =   34
            Top             =   1260
            Width           =   3735
            Begin VB.TextBox txt220v_3Qty 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   30
               Text            =   "0"
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox txt220v_2Qty 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   28
               Text            =   "0"
               Top             =   600
               Width           =   375
            End
            Begin VB.TextBox txt220v_1Qty 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   26
               Text            =   "0"
               Top             =   240
               Width           =   375
            End
            Begin VB.ComboBox cbo220v_1 
               Height          =   315
               ItemData        =   "frmElemInfo.frx":99A54
               Left            =   960
               List            =   "frmElemInfo.frx":99A56
               TabIndex        =   27
               Top             =   240
               Width           =   2655
            End
            Begin VB.ComboBox cbo220v_2 
               Height          =   315
               Left            =   960
               TabIndex        =   29
               Top             =   600
               Width           =   2655
            End
            Begin VB.ComboBox cbo220v_3 
               Height          =   315
               Left            =   960
               TabIndex        =   31
               Top             =   960
               Width           =   2655
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               Caption         =   "C:"
               Height          =   315
               Left            =   120
               TabIndex        =   40
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               Caption         =   "B:"
               Height          =   315
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               Caption         =   "A:"
               Height          =   315
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               Caption         =   "@"
               Height          =   255
               Left            =   720
               TabIndex        =   37
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               Caption         =   "@"
               Height          =   255
               Left            =   720
               TabIndex        =   36
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               Caption         =   "@"
               Height          =   255
               Left            =   720
               TabIndex        =   35
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.TextBox txtAir 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   25
            Text            =   "0"
            Top             =   615
            Width           =   615
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "500"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   990
            Width           =   480
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1000"
            Height          =   255
            Left            =   600
            TabIndex        =   46
            Top             =   990
            Width           =   480
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1500"
            Height          =   255
            Left            =   1080
            TabIndex        =   45
            Top             =   990
            Width           =   480
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2000"
            Height          =   255
            Left            =   1560
            TabIndex        =   44
            Top             =   990
            Width           =   480
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Phones"
            Height          =   255
            Left            =   2040
            TabIndex        =   43
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Water"
            Height          =   255
            Left            =   2640
            TabIndex        =   42
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CmpAir"
            Height          =   255
            Left            =   3240
            TabIndex        =   41
            Top             =   990
            Width           =   600
         End
      End
      Begin VB.ComboBox cboBase 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmElemInfo.frx":99A58
         Left            =   3540
         List            =   "frmElemInfo.frx":99A68
         TabIndex        =   15
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label lblPhoto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Photo Available"
         Height          =   495
         Left            =   6240
         TabIndex        =   68
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Sealed Prints Req'd for Major Shows?"
         Height          =   435
         Left            =   120
         TabIndex        =   60
         Top             =   4425
         Width           =   1830
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eng'd to L.A. Bldg Code?"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   4110
         Width           =   1770
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Storage Location:"
         Height          =   195
         Left            =   3540
         TabIndex        =   58
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label lblCanopyArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   3690
         Width           =   45
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Element is Ceiling Hung?"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   3270
         Width           =   1725
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eng Drawing No."
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   4920
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity of Units Crated as Single Element Code?"
         Height          =   405
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.Image imgSize 
      Height          =   735
      Left            =   8340
      ToolTipText     =   "Double-Click to Close"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmElemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lELTID As Long, lKITID As Long
Dim sHDR As String, sInit As String
Dim maxX As Double, maxY As Double, dTop As Double, dLeft As Double

Public Property Get PassELTID() As Long
    PassELTID = lELTID
End Property
Public Property Let PassELTID(ByVal vNewValue As Long)
    lELTID = vNewValue
End Property

Public Property Get PassKITID() As Long
    PassKITID = lKITID
End Property
Public Property Let PassKITID(ByVal vNewValue As Long)
    lKITID = vNewValue
End Property

Public Property Get PassHDR() As String
    PassHDR = sHDR
End Property
Public Property Let PassHDR(ByVal vNewValue As String)
    sHDR = vNewValue
End Property


Private Sub cbo220v_1_Change()
    If Len(cbo220v_1.Text) > 30 Then
        cbo220v_1.Text = Left(cbo220v_1.Text, 30)
        cbo220v_1.SelStart = Len(cbo220v_1.Text)
    End If
    ChangeCase cbo220v_1
End Sub

Private Sub cbo220v_2_Change()
    If Len(cbo220v_2.Text) > 30 Then
        cbo220v_2.Text = Left(cbo220v_2.Text, 30)
        cbo220v_2.SelStart = Len(cbo220v_2.Text)
    End If
    ChangeCase cbo220v_2
End Sub

Private Sub cbo220v_3_Change()
    If Len(cbo220v_3.Text) > 30 Then
        cbo220v_3.Text = Left(cbo220v_3.Text, 30)
        cbo220v_3.SelStart = Len(cbo220v_3.Text)
    End If
    ChangeCase cbo220v_3
End Sub

Private Sub cbxElecConf_Click()
    If cbxElecConf.value = 1 Then
        txtECInit.Text = sInit
        txtECInit.Visible = True
        txtECDate.Text = Format(Now, "m/d/yy")
        txtECDate.Visible = True
    Else
        txtECInit.Visible = False
        txtECDate.Visible = False
        txtECInit.Text = " "
        txtECDate.Text = Format(DateValue("1/1/100"), "m/d/yy")
    End If
End Sub

Private Sub cmdClear_Click()
    Screen.MousePointer = 11
    ClearEmAll
    Call FillElemInfo(lELTID, lKITID)
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strUpdate As String, strDelete As String
    Dim OHgt As Single, AssHgt As Single, AHgt As Single, BHgt As Single, CHgt As Single
    Dim iQty As Integer, iCeil As Integer, iCert As Integer, iSeal As Integer
    Dim lCano As Long
    Dim sDwg As String
    Dim ECDate As Date
    Dim ECInit As String
    
    Screen.MousePointer = 11
    
    If txtPrimeHgt.Text <> "" Then OHgt = CSng(txtPrimeHgt.Text) Else OHgt = 0
    If txtAssemHgt.Text <> "" Then AssHgt = CSng(txtAssemHgt.Text) Else AssHgt = 0
    If txtHgtA.Text <> "" Then AHgt = CSng(txtHgtA.Text) Else AHgt = 0
    If txtHgtB.Text <> "" Then BHgt = CSng(txtHgtB.Text) Else BHgt = 0
    If txtHgtC.Text <> "" Then CHgt = CSng(txtHgtC.Text) Else CHgt = 0
    
    If txtDwgNo.Text <> "" Then sDwg = txtDwgNo.Text Else sDwg = " "
    If cbxElecConf.value = 1 Then
        ECDate = CDate(txtECDate.Text)
        ECInit = txtECInit.Text
    Else
        ECDate = CDate("01/01/100")
        ECInit = " "
    End If
    
    On Error GoTo ErrorTrap
    Conn.BeginTrans
    strUpdate = "UPDATE " & IGLElt & " " & _
                "SET PRIHT = " & OHgt & ", " & _
                "ASMHT = " & AssHgt & ", " & _
                "UNITCNT = " & txtQtyPerElem.Text & ", " & _
                "CEILHANG = " & cboClgHung.ItemData(cboClgHung.ListIndex) & ", " & _
                "CANAREA = " & CLng(txtCanopyArea.Text) & ", " & _
                "SEISMICF = " & cboCertLA.ItemData(cboCertLA.ListIndex) & ", " & _
                "SEALREQD = " & cboSealReq.ItemData(cboSealReq.ListIndex) & ", " & _
                "DRWGNBR = '" & sDwg & "', " & _
                "ELCONFDT = TO_DATE('" & ECDate & "', 'MM/DD/YYYY'), " & _
                "ELCONFUR = '" & ECInit & "', " & _
                "UPDUSER = '" & Left(DeGlitch(LogName), 16) & "', " & _
                "UPDDTTM = SYSDATE, UPDCNT = UPDCNT + 1 " & _
                "WHERE KITID = " & lKITID & " " & _
                "AND ELTID = " & lELTID
    Conn.Execute (strUpdate)
    
    '///// DELETE ALL FROM ELTEXT, EXCEPT OPTELTS \\\\\
    strDelete = "DELETE FROM " & IGLEltX & " " & _
                "WHERE ELTID = " & lELTID & " " & _
                "AND EEXTYPE <> 9"
    Conn.Execute (strDelete)
    '///// NOW WRITE NEW \\\\\
    If txt500w.Text > 0 Then Err = WriteToEltExt(2, "500W", " ", CInt(txt500w.Text))
    If txt1000w.Text > 0 Then Err = WriteToEltExt(3, "1000W", " ", CInt(txt1000w.Text))
    If txt1500w.Text > 0 Then Err = WriteToEltExt(4, "1500W", " ", CInt(txt1500w.Text))
    If txt2000w.Text > 0 Then Err = WriteToEltExt(5, "2000W", " ", CInt(txt2000w.Text))
    If txtPhones.Text > 0 Then Err = WriteToEltExt(7, "PHONE", " ", CInt(txtPhones.Text))
    If txtWater.Text > 0 Then Err = WriteToEltExt(8, "WATER", " ", CInt(txtWater.Text))
    If txtAir.Text > 0 Then Err = WriteToEltExt(10, "AIR", " ", CInt(txtAir.Text))
    If txt220v_1Qty.Text > 0 Then
        Err = WriteToEltExt(6, "220V", cbo220v_1.Text, CInt(txt220v_1Qty.Text))
    End If
    If txt220v_2Qty.Text > 0 Then
        Err = WriteToEltExt(6, "220V", cbo220v_2.Text, CInt(txt220v_2Qty.Text))
    End If
    If txt220v_3Qty.Text > 0 Then
        Err = WriteToEltExt(6, "220V", cbo220v_3.Text, CInt(txt220v_3Qty.Text))
    End If
    If AHgt > 0 Then Err = WriteToEltExt(1, "OPTIONAL HEIGHT", "A", AHgt)
    If BHgt > 0 Then Err = WriteToEltExt(1, "OPTIONAL HEIGHT", "B", BHgt)
    If CHgt > 0 Then Err = WriteToEltExt(1, "OPTIONAL HEIGHT", "C", CHgt)
        
    Conn.CommitTrans
    
    Screen.MousePointer = 0
Exit Sub
ErrorTrap:
    Conn.RollbackTrans
    MsgBox "Error:  " & Err.Description, vbExclamation, "Error Encountered..."
    Err.Clear
End Sub

Public Function WriteToEltExt(extype As Integer, extdesc As String, extvalue As String, extqty As Single) As Integer
    Dim rstN As ADODB.Recordset
    Dim strXInsert As String
    Dim EExtID As Long
    
    On Error Resume Next
    Set rstN = Conn.Execute("SELECT " & IGLSeq & ".NEXTVAL FROM DUAL")
    EExtID = rstN.Fields("NEXTVAL")
    rstN.Close: Set rstN = Nothing
    strXInsert = "INSERT INTO " & IGLEltX & " " & _
            "(KITID, ELTID, " & _
            "EEXTID, EEXTYPE, " & _
            "EEXTDESC, VALUE, " & _
            "QTY, UPDUSER, " & _
            "UPDDTTM, UPDCNT) " & _
            "VALUES " & _
            "(" & lKITID & ", " & lELTID & ", " & _
            EExtID & ", " & extype & ", " & _
            "'" & extdesc & "', '" & extvalue & "', " & _
            extqty & ", '" & Left(DeGlitch(LogName), 16) & "', " & _
            "SYSDATE, 1)"
    On Error Resume Next
    Conn.Execute (strXInsert)
    WriteToEltExt = Err
End Function

Private Sub Form_Load()
    Dim iSpc As Integer
    
    Screen.MousePointer = 11
    
    Me.Caption = "Element:  " & sHDR
    
    sInit = Left(LogName, 1)
    iSpc = InStr(1, LogName, " ")
    Do While iSpc <> 0
        sInit = sInit & Mid(LogName, iSpc + 1, 1)
        iSpc = InStr(iSpc + 1, LogName, " ")
    Loop
    
    txtECInit.Text = " "
    txtECDate.Text = " "
        
    With cboClgHung
        .AddItem "Yes": .ItemData(.NewIndex) = 1
        .AddItem "No": .ItemData(.NewIndex) = 2
        .AddItem "TBD": .ItemData(.NewIndex) = 0
        .AddItem "CA Only": .ItemData(.NewIndex) = 3
    End With
    With cboCertLA
        .AddItem "Yes": .ItemData(.NewIndex) = 1
        .AddItem "No": .ItemData(.NewIndex) = 2
        .AddItem "TBD": .ItemData(.NewIndex) = 0
    End With
    With cboSealReq
        .AddItem "Yes": .ItemData(.NewIndex) = 1
        .AddItem "No": .ItemData(.NewIndex) = 2
        .AddItem "TBD": .ItemData(.NewIndex) = 0
    End With
    
    maxX = picPhoto.Width
    maxY = picPhoto.Height
    dTop = picPhoto.Top
    dLeft = picPhoto.Left
    
    If bPerm(42) Then
        Call Fill220Lists
        fraElemInfo.Enabled = True
        Me.Height = 6600
    Else
        fraElemInfo.Enabled = False
        Me.Caption = Me.Caption & "   [Read Only]"
        Me.Height = 6000
    End If
    
    Call FillElemInfo(lELTID, lKITID)
    Call CheckForPhoto(lELTID)
        
    Screen.MousePointer = 0
End Sub

Public Function FillElemInfo(EID As Long, KID As Long)
    Dim rstE As ADODB.Recordset
    Dim rstX As ADODB.Recordset
    Dim rstW As ADODB.Recordset
    Dim strESelect As String
    Dim strXSelect As String
    Dim strWSelect As String
    Dim i As Integer, i220v As Integer, iOpt As Integer
    Dim dWgt As Double
    Dim sDim As String, sDisc As String
    
    strESelect = "SELECT * " & _
                "FROM " & IGLElt & " " & _
                "WHERE ELTID = " & EID & " " & _
                "AND KITID = " & KID
    Set rstE = Conn.Execute(strESelect)
    If Not rstE.EOF Then
        With rstE
            txtDesc.Text = Trim(.Fields("ELTDESC"))
            If Not IsNull(.Fields("MCU")) Then txtCPRJ.Text = .Fields("MCU") Else txtCPRJ.Text = ""
            If .Fields("HTUNIT") = 1 Then optHgtUnit(0).value = True
            If .Fields("HTUNIT") = 5 Then optHgtUnit(1).value = True
            If bPerm(42) Then
                txtPrimeHgt.Text = Format(CSng(.Fields("PRIHT")), "0.00")
            Else
                If .Fields("HTUNIT") = 1 Then
                    txtPrimeHgt.Text = CalcDim(CSng(.Fields("PRIHT")))
                Else
                    txtPrimeHgt.Text = Format(CSng(.Fields("PRIHT")), "0.00") & " cm"
                End If
            End If
            If .Fields("ASMHT") > 0 Then
                If bPerm(42) Then
                    txtAssemHgt.Text = Format(CSng(.Fields("ASMHT")), "0.00")
                Else
                    txtAssemHgt.Text = CalcDim(CSng(.Fields("ASMHT")))
                End If
            Else
                txtAssemHgt.Text = ""
            End If
            
            txtQtyPerElem.Text = .Fields("UNITCNT")
'''            CurrUnitCnt = .FIELDS("T$UNITCNT")
            If .Fields("CEILHANG") = 0 Then cboClgHung.Text = "TBD"
            If .Fields("CEILHANG") = 1 Then cboClgHung.Text = "Yes"
            If .Fields("CEILHANG") = 2 Then cboClgHung.Text = "No"
            If .Fields("CEILHANG") = 3 Then cboClgHung.Text = "CA Only"
            txtCanopyArea.Text = .Fields("CANAREA")
            If .Fields("SEISMICF") = 0 Then cboCertLA.Text = "TBD"
            If .Fields("SEISMICF") = 1 Then cboCertLA.Text = "Yes"
            If .Fields("SEISMICF") = 2 Then cboCertLA.Text = "No"
            If .Fields("SEALREQD") = 0 Then cboSealReq.Text = "TBD"
            If .Fields("SEALREQD") = 1 Then cboSealReq.Text = "Yes"
            If .Fields("SEALREQD") = 2 Then cboSealReq.Text = "No"
            If Trim(.Fields("DRWGNBR")) = "0" Then _
                        txtDwgNo.Text = "" _
                        Else txtDwgNo.Text = Trim(.Fields("DRWGNBR"))
            chkContainer.value = .Fields("CONTAINF")
            
            dWgt = 0: sDisc = ""
            strWSelect = "SELECT WTUNIT, WEIGHT " & _
                        "FROM " & IGLPart & " " & _
                        "WHERE KITID = " & KID & " " & _
                        "AND ELTID = " & EID & " " & _
                        "AND TSTATUS > 0"
            Set rstW = Conn.Execute(strWSelect)
            If Not rstW.EOF Then
                If rstW.Fields("WTUNIT") = 1 Then optWgtUnit(0).value = True
                If rstW.Fields("WTUNIT") = 2 Then optWgtUnit(1).value = True
                Do While Not rstW.EOF
                    dWgt = dWgt + rstW.Fields("WEIGHT")
                    If rstW.Fields("WEIGHT") = 0 Then
                        lblDisc.Visible = True
                        sDisc = "*"
                    End If
                    rstW.MoveNext
                Loop
                rstW.Close
                Set rstW = Nothing
            End If
            txtEstWgt.Text = Format(dWgt, "#,##0") & sDisc
            
            
            '***** BASELOC STILL TO FOLLOW *****
            If Len(Trim(.Fields("ELCONFUR"))) > 0 Then
                cbxElecConf.value = 1
                txtECDate.Text = Format(.Fields("ELCONFDT"), "m/d/yy")
                txtECInit.Text = Trim(.Fields("ELCONFUR"))
                txtECInit.Visible = True
                txtECDate.Visible = True
            End If
        End With
        strXSelect = "SELECT * " & _
                    "FROM " & IGLEltX & " " & _
                    "WHERE ELTID = " & EID & " " & _
                    "ORDER BY EEXTYPE, VALUE"
        Set rstX = Conn.Execute(strXSelect)
        i220v = 1
        Do While Not rstX.EOF
            With rstX
                Select Case Trim(rstX.Fields("EEXTDESC"))
                Case "500W"
                    txt500w.Text = .Fields("QTY")
                Case "1000W"
                    txt1000w.Text = .Fields("QTY")
                Case "1500W"
                    txt1500w.Text = .Fields("QTY")
                Case "2000W"
                    txt2000w.Text = .Fields("QTY")
                Case "PHONE"
                    txtPhones.Text = .Fields("QTY")
                Case "WATER"
                    txtWater.Text = .Fields("QTY")
                Case "AIR"
                    txtAir.Text = .Fields("QTY")
                Case "220V"
                    Select Case i220v
                    Case 1
                        cbo220v_1.Text = Trim(.Fields("VALUE"))
                        txt220v_1Qty.Text = .Fields("QTY")
                        i220v = 2
                    Case 2
                        cbo220v_2.Text = Trim(.Fields("VALUE"))
                        txt220v_2Qty.Text = .Fields("QTY")
                        i220v = 3
                    Case 3
                        cbo220v_3.Text = Trim(.Fields("VALUE"))
                        txt220v_3Qty.Text = .Fields("QTY")
                        i220v = 4
                    End Select
                Case "OPTIONAL HEIGHT"
                    Select Case UCase(Trim(.Fields("VALUE")))
                    Case "A"
                        If bPerm(42) Then sDim = Format(CSng(.Fields("QTY")), "0.00") _
                                    Else sDim = CalcDim(CSng(.Fields("QTY")))
                        txtHgtA.Text = sDim
                    Case "B"
                        If bPerm(42) Then sDim = Format(CSng(.Fields("QTY")), "0.00") _
                                    Else sDim = CalcDim(CSng(.Fields("QTY")))
                        txtHgtB.Text = sDim
                    Case "C"
                        If bPerm(42) Then sDim = Format(CSng(.Fields("QTY")), "0.00") _
                                    Else sDim = CalcDim(CSng(.Fields("QTY")))
                        txtHgtC.Text = sDim
                    End Select
'''''                Case "OPTIONAL ELEMENT"
'''''                    lblOpt(iOpt).Caption = Trim(.FIELDS("VALUE"))
'''''                    iOpt = iOpt + 1
                End Select
                .MoveNext
            End With
        Loop
        rstX.Close
        Set rstX = Nothing
    End If
    rstE.Close
    Set rstE = Nothing
End Function

Private Sub imgSize_DblClick()
    imgSize.Visible = False
    picPhoto.Visible = True
    fraElemInfo.Visible = True
    cmdClear.Visible = True: cmdSave.Visible = True
End Sub

Private Sub optHgtUnit_Click(Index As Integer)
    Select Case Index
    Case 0
        lblCanopyArea.Caption = "Sqft of Canopy Area? <0 if none>"
    Case 1
        lblCanopyArea.Caption = "SqM of Canopy Area? <0 if none>"
    End Select
End Sub

Public Function CalcDim(Num As Single) As String
    Dim Feet As Integer, Inch As Integer, Numer As Integer
    Dim Frac As Currency
    Dim strFrac As String
    
    Feet = Int(Num / 12)
    Inch = Int(Num - (Feet * 12))
    Frac = CCur((((Num / 12) - Feet) * 12) - Inch)
    If Frac > 0 Then
        Numer = CInt(Frac * 8)
        Select Case Numer
            Case 1: strFrac = " 1/8"""
            Case 2: strFrac = " 1/4"""
            Case 3: strFrac = " 3/8"""
            Case 4: strFrac = " 1/2"""
            Case 5: strFrac = " 5/8"""
            Case 6: strFrac = " 3/4"""
            Case 7: strFrac = " 7/8"""
            Case 8
                Inch = Inch + 1
                If Inch = 12 Then
                    Feet = Feet + 1
                    Inch = 0
                End If
                strFrac = Chr(34)
        End Select
    Else
        strFrac = Chr(34)
    End If
    CalcDim = Feet & "'-" & Inch & strFrac

End Function

Public Sub CheckInteger(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = vbKeyBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub CheckNumeric(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = vbKeyBack And Not KeyAscii = 46 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub picPhoto_DblClick()
    picPhoto.Visible = False
    fraElemInfo.Visible = False
    imgSize.Visible = True
    cmdClear.Visible = False: cmdSave.Visible = False
End Sub

Private Sub txt1000w_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt1000w_LostFocus()
    Call SetToZero(txt1000w)
End Sub

Private Sub txt1500w_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt1500w_LostFocus()
    Call SetToZero(txt1500w)
End Sub

Private Sub txt2000w_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt2000w_LostFocus()
    Call SetToZero(txt2000w)
End Sub

Private Sub txt220v_1Qty_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt220v_1Qty_LostFocus()
    Call SetToZero(txt220v_1Qty)
End Sub

Private Sub txt220v_2Qty_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt220v_2Qty_LostFocus()
    Call SetToZero(txt220v_2Qty)
End Sub

Private Sub txt220v_3Qty_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt220v_3Qty_LostFocus()
    Call SetToZero(txt220v_3Qty)
End Sub

Private Sub txt500w_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txt500w_LostFocus()
    Call SetToZero(txt500w)
End Sub

Private Sub txtAir_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txtAir_LostFocus()
    Call SetToZero(txtAir)
End Sub

Private Sub txtAssemHgt_KeyPress(KeyAscii As Integer)
    Call CheckNumeric(KeyAscii)
End Sub

Private Sub txtCanopyArea_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txtCanopyArea_LostFocus()
    Call SetToZero(txtCanopyArea)
End Sub

Private Sub txtDwgNo_Change()
    ChangeCase txtDwgNo
End Sub

Private Sub txtECInit_Change()
    ChangeCase txtECInit
End Sub

Private Sub txtHgtA_KeyPress(KeyAscii As Integer)
    Call CheckNumeric(KeyAscii)
End Sub

Private Sub txtHgtB_KeyPress(KeyAscii As Integer)
    Call CheckNumeric(KeyAscii)
End Sub

Private Sub txtHgtC_KeyPress(KeyAscii As Integer)
    Call CheckNumeric(KeyAscii)
End Sub

Private Sub txtPhones_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txtPhones_LostFocus()
    Call SetToZero(txtPhones)
End Sub

Private Sub txtPrimeHgt_KeyPress(KeyAscii As Integer)
    Call CheckNumeric(KeyAscii)
End Sub

Private Sub txtQtyPerElem_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Private Sub txtQtyPerElem_LostFocus()
    Call SetToOne(txtQtyPerElem)
End Sub

Private Sub txtWater_KeyPress(KeyAscii As Integer)
    Call CheckInteger(KeyAscii)
End Sub

Public Function ChangeCase(Con As Control)
    Dim Pos As Integer
    Pos = Con.SelStart
    Con.Text = UCase(Con.Text)
    Con.SelStart = Pos
End Function

Public Function Fill220Lists()
    Dim rstV As ADODB.Recordset
    Dim strVSelect As String
    '****CONN IS ALREADY OPEN****
    strVSelect = "SELECT DISTINCT VALUE " & _
                "FROM " & IGLEltX & " " & _
                "WHERE EEXTDESC = '220V' " & _
                "ORDER BY VALUE"
    Set rstV = Conn.Execute(strVSelect)
    Do While Not rstV.EOF
        cbo220v_1.AddItem Trim(rstV.Fields("VALUE"))
        cbo220v_2.AddItem Trim(rstV.Fields("VALUE"))
        cbo220v_3.AddItem Trim(rstV.Fields("VALUE"))
        rstV.MoveNext
    Loop
    rstV.Close
    Set rstV = Nothing
End Function

Public Function SetToZero(tbx As TextBox)
    If Len(Trim(tbx.Text)) = 0 Then tbx.Text = 0
End Function

Public Function SetToOne(tbx As TextBox)
    If Len(Trim(tbx.Text)) = 0 Then tbx.Text = 1
End Function

Private Sub txtWater_LostFocus()
    Call SetToZero(txtWater)
End Sub

Public Sub ClearEmAll()
    txtPrimeHgt.Text = ""
    txtHgtA.Text = ""
    txtHgtB.Text = ""
    txtHgtC.Text = ""
    txtAssemHgt.Text = ""
    txtQtyPerElem.Text = 1
    cboClgHung.Text = "No"
    txtCanopyArea.Text = 0
    cboCertLA.Text = "TBD"
    cboSealReq.Text = "TBD"
    txtDwgNo.Text = ""
    cbxElecConf.value = 0
    txtECInit.Visible = False
    txtECInit.Text = " "
    txtECDate.Visible = False
    txt500w.Text = 0
    txt1000w.Text = 0
    txt1500w.Text = 0
    txt2000w.Text = 0
    txtPhones.Text = 0
    txtWater.Text = 0
    txtAir.Text = 0
    txt220v_1Qty.Text = 0
    txt220v_2Qty.Text = 0
    txt220v_3Qty.Text = 0
    cbo220v_1.Text = ""
    cbo220v_2.Text = ""
    cbo220v_3.Text = ""
End Sub

Public Sub CheckForPhoto(EID As Long)
    Dim strSelect As String
    Dim rst As ADODB.Recordset
    Dim rAsp As Double, rFAsp As Double
    
    strSelect = "SELECT M.GPATH " & _
                "FROM " & GFXElt & " E, " & GFXMas & " M " & _
                "WHERE E.ELTID = " & EID & " " & _
                "AND E.GID = M.GID " & _
                "AND M.GTYPE = 1 " & _
                "AND M.GSTATUS > 0 " & _
                "ORDER BY M.ADDDTTM DESC"
    Set rst = Conn.Execute(strSelect)
    If Not rst.EOF Then
        picPhoto.Picture = LoadPicture()
        imgSize.Picture = LoadPicture(rst.Fields("GPATH"))
        Debug.Print "X = " & imgSize.Width & ", y = " & imgSize.Height
        rAsp = imgSize.Width / imgSize.Height
        rFAsp = maxX / maxY
        Select Case rAsp
            Case Is = rFAsp
                With picPhoto
                    .Width = maxX
                    .Height = maxY
                    .Top = dTop: .Left = dLeft
                End With
            Case Is > rFAsp
                With picPhoto
                    .Width = maxX
                    .Height = .Width / rAsp
                    .Top = dTop + ((maxY - .Height) / 2)
                    .Left = dLeft
                End With
            Case Is < rFAsp
                With picPhoto
                    .Height = maxY
                    .Width = .Height * rAsp ''''' / rFAsp)
                    .Top = dTop
                    .Left = dLeft + (maxX - .Width) / 2
                End With
        End Select
        picPhoto.PaintPicture imgSize.Picture, 0, 0, picPhoto.Width, picPhoto.Height
        picPhoto.Visible = True
        
        '///// RESIZE IMGSIZE TO FIT FORM \\\\\
        imgSize.Stretch = True
        rFAsp = (Me.ScaleWidth - 120) / (Me.ScaleHeight - 120)
        Select Case rAsp
            Case Is = rFAsp
                With imgSize
                    .Width = Me.ScaleWidth - 120
                    .Height = Me.ScaleHeight - 120
                    .Top = 60: .Left = 60
                End With
            Case Is > rFAsp
                With imgSize
                    .Width = Me.ScaleWidth - 120
                    .Height = .Width / rAsp
                    .Top = 60 + (((Me.ScaleHeight - 120) - .Height) / 2)
                    .Left = 60
                End With
            Case Is < rFAsp
                With imgSize
                    .Height = Me.ScaleHeight - 120
                    .Width = .Height * rAsp ''''' / rFAsp)
                    .Top = 60
                    .Left = ((Me.ScaleWidth - 120) - .Width) / 2
                End With
        End Select
    Else
        picPhoto.Visible = False
    End If
    rst.Close
    Set rst = Nothing
End Sub
