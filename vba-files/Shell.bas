'@Folder("MORFunctions.Modules")
Attribute VB_Name = "Shell"
Option Explicit

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
wShowWindow As Integer
cbReserved2 As Integer
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

Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
lpStartupInfo As STARTUPINFO, lpProcessInformation As _
PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&


Public Sub ExecCmd(cmdline As String, asynchronously As Boolean, Optional HasFocus As Boolean)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ReturnValue As Integer

' Initialize the STARTUPINFO structure:
start.cb = Len(start)
If HasFocus = False Then
    start.dwFlags = 1
    start.wShowWindow = 6
End If

' Start the shelled application:
ReturnValue = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

If asynchronously = True Then
    ' Wait for the shelled application to finish:
    Do
    ReturnValue = WaitForSingleObject(proc.hProcess, 0)
    DoEvents
    Loop Until ReturnValue <> 258
End If
ReturnValue = CloseHandle(proc.hProcess)
End Sub
Public Function OpenFile(Fullpath As String)
Dim result As Long
File2Open = ShellExecute(0&, vbNullString, Fullpath, vbNullString, vbNullString, vbNormalFocus)
End Function

Public Function ShellRun(sCmd As String) As String

''' <summary>
'''   Executes the given executable in a shell instance and returns the output produced
'''   by it. If iStdInput is given, it is passed to the executable during execution.
'''   Note: You must make sure to correctly enclose the executable path or any given
'''         arguments in quotes (") if they contain spaces.
''' </summary>
''' <param name="iExecutablePath">
'''   The full path to the executable (and its parameters). This string is passed to the
'''   shell unaltered, so be sure to enclose it in quotes if it contains spaces.
''' </param>
''' <param name="iStdInput">
'''   The (optional) input to pass to the executable. Default: Null
''' </param>''You will need to add a reference to Windows Script Host Object Model so the types WshShell and WshExec are known.
'''  To do this go to Extras -> References in the VBA IDE's menu bar.)
'''  Note: Note: You will need to add a reference to Windows Script Host Object Model so the types WshShell and WshExec are known.
'''  (To do this go to Extras -> References in the VBA IDE's menu bar.)


    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim S As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then S = S & sLine & vbCrLf
    Wend

    ShellRun = S

End Function
