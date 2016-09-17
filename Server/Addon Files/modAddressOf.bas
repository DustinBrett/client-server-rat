Attribute VB_Name = "modAddressOf"
Option Explicit

'USER32.DLL (Function)
Private Declare Function GetWindowTextA Lib "USER32.DLL" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLengthA Lib "USER32.DLL" (ByVal hWnd As Long) As Long

Public Const vbParseData As String = ""

Public strWindows As String

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    'On Error Resume Next
    Dim strTitle As String, strLen As Integer
    
    'Get Window Text Length
    strLen = GetWindowTextLengthA(hWnd)
    strTitle = Space$(strLen)
    
    'Get Window Text
    GetWindowTextA hWnd, strTitle, strLen + 1
    
    'Set strWindows With Window Text and Window hWnd
    If LenB(strTitle) Then strWindows = strWindows & strTitle & vbParseData & hWnd & vbParseData
    
    'Window Found
    EnumWindowsProc = True
End Function
