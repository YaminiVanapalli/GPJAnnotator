VERSION 5.00
Object = "{8718C64B-8956-11D2-BD21-0060B0A12A50}#1.0#0"; "avviewx.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   15630
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   Begin VOLOVIEWXLibCtl.AvViewX AvViewX1 
      Height          =   1095
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      _cx             =   2566
      _cx             =   1931
      src             =   ""
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      UserMode        =   "Select"
      HighlightLinks  =   0   'False
      LayersOff       =   ""
      LayersOn        =   ""
      SrcTemp         =   ""
      FontPath        =   "C:\Program Files\Volo View Express\fonts"
      NamedView       =   ""
      BackgroundColor =   "DefaultColors"
      GeometryColor   =   "DefaultColors"
      PrintBackgroundColor=   "DefaultColors"
      PrintGeometryColor=   "DefaultColors"
      ShadingMode     =   "WireFrame"
      ProjectionMode  =   "Parallel"
      EnableUIMode    =   "DefaultUI"
      Layout          =   ""
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
