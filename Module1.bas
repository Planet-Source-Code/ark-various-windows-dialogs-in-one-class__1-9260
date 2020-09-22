Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_USER = &H400
Const BFFM_INITIALIZED = 1
Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
'
Public Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function
'
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   Select Case uMsg
      Case BFFM_INITIALIZED
           Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
      Case Else:
   End Select
End Function

