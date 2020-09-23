VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "School Countdown 3.0"
   ClientHeight    =   1260
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   1260
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMinimize 
      Interval        =   10000
      Left            =   2145
      Top             =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Settings"
      CausesValidation=   0   'False
      Height          =   285
      Left            =   1065
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   915
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   315
      Top             =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Left Till the End Of The School Year"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   435
      Width           =   2700
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2970
   End
   Begin VB.Image close1 
      Height          =   150
      Left            =   2670
      Picture         =   "Form1.frx":190A
      Top             =   900
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image close2 
      Height          =   150
      Left            =   2850
      Picture         =   "Form1.frx":1BA9
      Top             =   900
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image closei 
      Height          =   150
      Left            =   2955
      Picture         =   "Form1.frx":1E44
      Top             =   60
      Width           =   150
   End
   Begin VB.Label lblMove 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   15
      TabIndex        =   2
      Top             =   0
      Width           =   3150
   End
   Begin VB.Menu mSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dateVar As Date
Dim ism As Boolean


'|||||||||||||||| GRAPHICS |||||||||||||||

Private Sub closei_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
closei.Picture = close1.Picture
End Sub

Private Sub closei_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
closei.Picture = close2.Picture
Call WriteToINI("Main", "Xpos", Form1.Left, App.Path & "\data.ini")
Call WriteToINI("Main", "Ypos", Form1.Top, App.Path & "\data.ini")
'----------- Minimize to taskbar ------------------
hideForm
'----------- /Minimize to taskbar ------------------
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

'|||||||||||||||| /GRAPHICS |||||||||||||||


'|||||||||||||||| FORM CONTROLS |||||||||||||||||||||||||||||
Private Sub Form_Load()
Xpos = GetFromINI("Main", "Xpos", App.Path & "\data.ini")
Ypos = GetFromINI("Main", "Ypos", App.Path & "\data.ini")
Form1.Top = Ypos
Form1.Left = Xpos
dateVar = GetFromINI("Main", "date", App.Path & "\data.ini")
End Sub






'|||||||||||||||| /FORM CONTROLS |||||||||||||||||||||||||||||

Private Sub Timer1_Timer()
 On Error Resume Next
    Dim i As Long, s As String, s2() As String
    Dim SecsRemaining As Long 'total secs to target date
    Dim dd As Long 'days to target date
    Dim hh As Long 'hrs to target date
    Dim mm As Long 'mins to target date
    Dim ss As Long 'secs to target date
    
    SecsRemaining = DateDiff("s", Now, dateVar)
    divideTime SecsRemaining, dd, hh, mm, ss
    Label1.Caption = dd & " days " & hh & " hours " & mm & " minutes " & ss & " seconds"
    
    
End Sub
Private Sub divideTime(ByVal i As Long, DaysLeft As Long, HoursLeft As Long, MinsLeft As Long, SecsLeft As Long)
    '# of days
    DaysLeft = i \ 86400
    i = i Mod 86400
    '# of hours
    HoursLeft = i \ 3600
    i = i Mod 3600
    '# of mins
    MinsLeft = i \ 60
    i = i Mod 60
    '# of secs
    SecsLeft = i
End Sub

'|||||||||||||| SYSTRAY CONTROLS |||||||||||||||||
'(You have to have a hidden menu)

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending
       'upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
       End If
       Select Case msg
        Case WM_LBUTTONUP        ' restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    ' restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        ' display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mSystray
       End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'remove the icon from the systray
Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub hideForm()
Me.Show
       Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "School Countdown 3.0" & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid
    'This is for a tray start up
    Me.Visible = False
End Sub

Private Sub mExit_Click()
Unload Me
Unload Form2
End Sub
Private Sub mRestore_Click()
       Dim Result As Long
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
End Sub

Private Sub tmrMinimize_Timer()
hideForm
End Sub
