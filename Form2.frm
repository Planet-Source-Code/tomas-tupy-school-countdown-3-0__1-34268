VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3390
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   6747
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Date"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DTPicker1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Skins"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   675
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   37375
         MaxDate         =   401768
         MinDate         =   37257
      End
      Begin VB.Label Label2 
         Caption         =   "Do this yourself"
         Height          =   240
         Left            =   -74865
         TabIndex        =   3
         Top             =   615
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Date School Ends:"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   435
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dateVar As Date
Private Sub DTPicker1_Change()
Call WriteToINI("Main", "date", DTPicker1.Value, App.Path & "\data.ini")
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Form_Load()
dateVar = GetFromINI("Main", "date", App.Path & "\data.ini")
DTPicker1.Value = dateVar
End Sub
