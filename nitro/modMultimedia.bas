Attribute VB_Name = "modMultimedia"
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Type mnuCommands
 Captions As New Collection
 Commands As New Collection
End Type

Public Type filetype
 Commands As mnuCommands
 extension As String
 ProperName As String
 FullName As String
 ContentType As String
 IconPath As String
 IconIndex As Integer
End Type



Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const AliasName = "REDIBMP3"



Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_SETTEXT = &HC
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
Public Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
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
Public Sub CreateExtension(newfiletype As filetype)
Dim IconString As String
Dim Result As Long, Result2 As Long, ResultX As Long
Dim ReturnValue As Long, HKeyX As Long
Dim cmdloop As Integer
IconString = newfiletype.IconPath & "," & newfiletype.IconIndex
If Left$(newfiletype.extension, 1) <> "." Then newfiletype.extension = "." & newfiletype.extension
RegCreateKey HKEY_CLASSES_ROOT, newfiletype.extension, Result
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, ByVal newfiletype.ProperName, LenB(StrConv(newfiletype.ProperName, vbFromUnicode)))
If newfiletype.ContentType <> "" Then
ReturnValue = RegSetValueEx(Result, "Content Type", 0, REG_SZ, ByVal CStr(newfiletype.ContentType), LenB(StrConv(newfiletype.ContentType, vbFromUnicode)))
End If
RegCreateKey HKEY_CLASSES_ROOT, newfiletype.ProperName, Result
If Not IconString = ",0" Then
RegCreateKey Result, "DefaultIcon", Result2
ReturnValue = RegSetValueEx(Result2, "", 0, REG_SZ, ByVal IconString, LenB(StrConv(IconString, vbFromUnicode)))
End If
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, ByVal newfiletype.FullName, LenB(StrConv(newfiletype.FullName, vbFromUnicode)))
RegCreateKey Result, ByVal "Shell", ResultX
For cmdloop = 1 To newfiletype.Commands.Captions.Count
RegCreateKey ResultX, ByVal newfiletype.Commands.Captions(cmdloop), Result
RegCreateKey Result, ByVal "Command", Result2
Dim CurrentCommand$
CurrentCommand = newfiletype.Commands.Commands(cmdloop)
ReturnValue = RegSetValueEx(Result2, "", 0, REG_SZ, ByVal CurrentCommand$, LenB(StrConv(CurrentCommand$, vbFromUnicode)))
RegCloseKey Result
RegCloseKey Result2
Next
RegCloseKey Result2
End Sub

Public Function OpenMultimedia(hWnd As Long, AliasName As String, FileName As String, typeDevice As String) As String
Dim tmp As String * 255
lenShort = GetShortPathName(FileName, tmp, 255)
mciSendString "open " & Left$(tmp, lenShort) & " type " & typeDevice & " Alias " & AliasName & " parent " & hWnd & " Style " & &H40000000, 0&, 0&, 0&
End Function

Public Sub PlayMultimedia(AliasName As String, from_where As String, to_where As String)
to_where = GetTotalframes(AliasName)
mciSendString "play " & AliasName & " from " & from_where & " to " & to_where, 0&, 0&, 0&
End Sub

Public Function GetTotalframes(AliasName As String) As Long
Dim Total As String * 128
mciSendString "status " & AliasName & " length", Total, 128, 0&
GetTotalframes = Val(Total)
End Function

Public Sub CloseMultimedia(AliasName As String)
mciSendString "Close " & AliasName, 0&, 0&, 0&
End Sub

Public Sub PauseMultimedia(AliasName As String)
mciSendString "Pause " & AliasName, 0&, 0&, 0&
End Sub

Public Sub StopMultimedia(AliasName As String)
mciSendString "Stop " & AliasName, 0&, 0&, 0&
End Sub

Public Sub ResumeMultimedia(AliasName As String)
mciSendString "Resume " & AliasName, 0&, 0&, 0&
End Sub

Public Sub CloseAll()
mciSendString "Close All", 0&, 0&, 0&
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
GetDefaultDevice = Left$(tmp, Res)
End Function


Public Function AreMultimediaAtEnd(AliasName As String) As Boolean
lastFrame = GetTotalframes(AliasName)
currpos = Val(GetCurrentMultimediaPos(AliasName))
If lastFrame = currpos Or (lastFrame - 1) < currpos Then
AreMultimediaAtEnd = True
Else
AreMultimediaAtEnd = False
End If
End Function

Public Function GetCurrentMultimediaPos(AliasName As String) As Long
Dim pos As String * 128
mciSendString "status " & AliasName & " position", pos, 128, 0&
GetCurrentMultimediaPos = Val(pos)
End Function
