Attribute VB_Name = "rwINITask"

Option Explicit
'------------ Minimize To Taskbar Declarations ---------------
 Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201
      Public Const WM_LBUTTONUP = &H202
      Public Const WM_LBUTTONDBLCLK = &H203
      Public Const WM_RBUTTONDOWN = &H204
      Public Const WM_RBUTTONUP = &H205
      Public Const WM_RBUTTONDBLCLK = &H206

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA
'------------ /Minimize To Taskbar Declarations ---------------
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
 Dim strBuffer As String
 strBuffer = String(750, Chr(0))
 Key$ = LCase$(Key$)
 GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
  Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub




