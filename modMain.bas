Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&

Private Const STARTF_USESHOWWINDOW& = &H1

Private Const INFINITE = -1&

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Single
   cbReserved2 As Single
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Type CurScriptSettings
    ScriptStoredProcedures As Boolean
    ScriptViews As Boolean
    ScriptTables As Boolean
    ScriptColumns As Boolean
    ScriptOwnerDiff As Boolean
    ScriptAutoProcess As Boolean
End Type

Public ScriptSettings As CurScriptSettings

Public Function ComputerNaam() As String
    Dim pcName As String

    pcName = Space(250)

    GetComputerName pcName, Len(pcName)

    pcName = Left(pcName, InStr(pcName, vbNullChar) - 1)
    
    ComputerNaam = pcName

End Function

Public Function ExecCmd(ByVal Cmdline As String, ByRef CmdSucceeded As Boolean)
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    CmdSucceeded = False
    Dim proc As PROCESS_INFORMATION
    Dim Start As STARTUPINFO
    
    Start.cb = Len(Start)

    Start.dwFlags = STARTF_USESHOWWINDOW
    Start.wShowWindow = 1
        
    
    Dim Ret As Long
    
    Ret& = CreateProcessA(vbNullString, Cmdline$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, vbNullString, Start, proc)
    
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, Ret&)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = Ret&
    CmdSucceeded = True
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
Err_Handler:
    Screen.MousePointer = vbDefault
    CmdSucceeded = False
    Err.Clear
    On Error GoTo 0
End Function
