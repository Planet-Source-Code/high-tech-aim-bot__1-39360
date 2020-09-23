Attribute VB_Name = "Module1"
'created by high tech,
'claiming it as your own will result
'in you being raped
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Sub ReturnIM(SayWhat)
Dim aimimessage As Long, wndateclass As Long, ateclass As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
wndateclass = FindWindowEx(aimimessage, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimimessage, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass, WM_SETTEXT, 0&, SayWhat)
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)
End Sub

Function IM_GetLastLine()
Dim aimimessage As Long, wndateclass As Long, ateclass As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
wndateclass = FindWindowEx(aimimessage, 0&, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(ateclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(ateclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
IM_GetLastLine = TrimHTML(TheText)
For I = Len(IM_GetLastLine) To 1 Step -1
If Mid(IM_GetLastLine, I, 2) = vbCrLf Then Exit For
Next I
IM_GetLastLine = Right(IM_GetLastLine, Len(IM_GetLastLine) - I)
End Function

Function TrimHTML(HTML)
On Error Resume Next
Dim Switch As Boolean
Dim I As Integer
Dim Final As String
If HTML = "" Then Exit Function
HTML = Replace(HTML, "<BR>", vbCrLf)
For I = 1 To Len(HTML)
If Mid(HTML, I, 1) = "<" Then Switch = True
If Mid(HTML, I, 1) = ">" Then Switch = False
If Switch = False Then Final = Final & Mid(HTML, I, 1)
Next I
Final = Replace(Final, "<", "")
Final = Replace(Final, ">", "")
TrimHTML = Final
End Function

Function IM_GetLastSN()
On Error Resume Next
im_getlastsnx = Left(IM_GetLastLine, InStr(IM_GetLastLine, ":") - 1)
IM_GetLastSN = Right(im_getlastsnx, Len(im_getlastsnx) - 1)
End Function
Function IM_GetLastWords()
On Error Resume Next
im_getlastwordsx = Replace(IM_GetLastLine, IM_GetLastSN & ": ", "")
IM_GetLastWords = Right(im_getlastwordsx, Len(im_getlastwordsx) - 1)
End Function

Sub IM_Close()
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call SendMessageLong(aimimessage, WM_CLOSE, 0&, 0&)
End Sub
