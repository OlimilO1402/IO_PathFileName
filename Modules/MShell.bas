Attribute VB_Name = "MShell"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2'''''''''''''''''''''''''
'''''''''1'''''''''2'''''''''3'''''''''4'''''''''5'''''''''6'''''''''7'''''''''8'''''''''9'''''''''0'''''''''1'''''''''2'''''''''3'''''''''4'''''''''5'''''''''6'''''''''7'''''''''8'''''''''9'''''''''0'''''''''1'''''''''2'''''
''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0'''''
'        1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20        21        22
'23456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
'
Option Explicit 'OM 06.11.2015 13:45; Zeilen: 118
'And Win64
#If VBA7 = 0 Then
    'Private Enum LongPtr
    '    [_]
    'End Enum
#End If

Private Type STARTUPINFO       ' x86,   x64
    cb              As Long    '   4      4
    lpReserved      As LongPtr '   4      8
    lpDesktop       As LongPtr '   4      8
    lpTitle         As LongPtr '   4      8
    dwX             As Long    '   4      4
    dwY             As Long    '   4      4
    dwXSize         As Long    '   4      4
    dwYSize         As Long    '   4      4
    dwXCountChars   As Long    '   4      4
    dwYCountChars   As Long    '   4      4
    dwFillAttribute As Long    '   4      4
    dwFlags         As Long    '   4      4
    wShowWindow     As Integer '   2      2
    cbReserved2     As Integer '   2      2
    lpReserved2     As LongPtr '   4      8
    hStdInput       As LongPtr '   4      8
    hStdOutput      As LongPtr '   4      8
    hStdError       As LongPtr '   4      8
End Type                   ' Sum: 68     96

Private Type SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle       As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess    As LongPtr
    hThread     As LongPtr
    dwProcessID As Long
    dwThreadID  As Long
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessw

#If VBA7 Then
    
    Private Declare PtrSafe Function CreateProcessW Lib "kernel32" ( _
        ByVal lpAppName As LongPtr, _
        ByVal lpCmdLine As LongPtr, _
        lpProcAttr As Any, _
        lpThreadAttr As Any, _
        ByVal lpInheritedHandle As LongPtr, _
        ByVal lpCreationFlags As LongPtr, _
        ByVal lpEnv As Any, _
        ByVal lpCurDir As LongPtr, _
        lpStartupInfo As STARTUPINFO, _
        lpProcessInfo As PROCESS_INFORMATION _
        ) As LongPtr
    
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliSeconds As Long) As Long
        
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

#Else

    Private Declare Function CreateProcessW Lib "kernel32" ( _
        ByVal lpAppName As LongPtr, _
        ByVal lpCmdLine As LongPtr, _
        lpProcAttr As Any, _
        lpThreadAttr As Any, _
        ByVal lpInheritedHandle As LongPtr, _
        ByVal lpCreationFlags As LongPtr, _
        ByVal lpEnv As LongPtr, _
        ByVal lpCurDir As LongPtr, _
        lpStartupInfo As STARTUPINFO, _
        lpProcessInfo As PROCESS_INFORMATION _
        ) As LongPtr
    
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliSeconds As Long) As Long
        
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    
#End If

Private Const INFINITE              As Long = -1&
Private Const WAIT_TIMEOUT          As Long = 258&
Private Const STARTF_USESHOWWINDOW  As Long = &H1&
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const IDLE_PRIORITY_CLASS   As Long = &H40&
Private Const HIGH_PRIORITY_CLASS   As Long = &H80&

'#If VBA7 Then

Public Function ShellWait(ByVal PathName As String, _
                          Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, _
                          Optional ByVal WorkDir As String) As LongPtr
                          
    Dim AttrProc       As SECURITY_ATTRIBUTES: AttrProc.nLength = LenB(AttrProc) '
    Dim AttrThrd       As SECURITY_ATTRIBUTES: AttrThrd.nLength = LenB(AttrThrd)
    Dim CreationFlags  As Long:                   CreationFlags = NORMAL_PRIORITY_CLASS
    Dim bIsRunning     As Boolean
    
    Dim StartInfo      As STARTUPINFO
    With StartInfo
        
        .cb = LenB(StartInfo) 'x86: 68; Win64: 96
        'Debug.Print .cb
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = WindowStyle
        
    End With
    
    Dim ProcessInfo    As PROCESS_INFORMATION
    Dim lpEnvironment  As LongPtr
    Dim lpInheritHnd   As LongPtr
    Dim hr As LongPtr
    Dim nul As LongPtr
    
    'hr = CreateProcessW(nul, StrPtr(PathName), AttrProc, AttrThrd, lpInheritHnd, CreationFlags, lpEnvironment, nul, StartInfo, ProcessInfo)
    hr = CreateProcessW(nul, StrPtr(PathName), AttrProc, AttrThrd, nul, CreationFlags, nul, nul, StartInfo, ProcessInfo)
    
    If hr = 0 Then
        MErr.MessError "MShell", "ShellWait", PathName, hr
        Exit Function
    End If
    With ProcessInfo
        If .hProcess <> 0 Then
            
            Dim dwMilliSeconds As Long: dwMilliSeconds = 500
            
            Do While WaitForSingleObject(.hProcess, dwMilliSeconds) = WAIT_TIMEOUT
                DoEvents
            Loop
            
        End If
        Debug.Print "CloseHandle .hProcess"
        CloseHandle .hProcess
        
    End With
End Function
'
'#Else
'
'    Public Function ShellWait(ByVal PathName As String, _
'                              Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, _
'                              Optional ByVal WorkDir As String) As Long
'
'        Dim AttrProc As SECURITY_ATTRIBUTES
'        Dim AttrThrd As SECURITY_ATTRIBUTES
'        Dim CreationFlags As Long
'        Dim lpEnvironment As Long
'        Dim StartInfo   As STARTUPINFO
'        Dim ProcessInfo As PROCESS_INFORMATION
'        Dim bIsRunning  As Boolean
'        Dim dwMilliSeconds As Long
'
'        AttrProc.nLength = LenB(AttrProc) '
'        AttrThrd.nLength = LenB(AttrThrd)
'        CreationFlags = NORMAL_PRIORITY_CLASS
'
'        With StartInfo
'
'            .cb = LenB(StartInfo) 'x86: 68; Win64: 96
'            .dwFlags = STARTF_USESHOWWINDOW
'            .wShowWindow = WindowStyle
'
'        End With
'
'        If CreateProcess(vbNullString, PathName, AttrProc, AttrThrd, _
'                         False, CreationFlags, lpEnvironment, WorkDir, _
'                         StartInfo, ProcessInfo) <> False Then
'
'            With ProcessInfo
'                If .hProcess <> 0 Then
'                    'Debug.Print .hProcess
'                    dwMilliSeconds = 500
'
'                    Do While WaitForSingleObject(.hProcess, dwMilliSeconds) = WAIT_TIMEOUT
'                        DoEvents
'                    Loop
'                'Else
'                    'Debug.Print "hProcess = 0"
'                End If
'
'                CloseHandle .hProcess
'
'            End With
'        End If
'    End Function
'
'#End If
