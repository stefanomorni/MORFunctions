'@Folder("MORFunctions.Modules")
Attribute VB_Name = "ManageDLL"
Option Explicit

'------------------------------------------------------------------------------------------------
'This module contains macros that make the registration/unregistration of COM DLL files easier.
'The macros are based on regsvr32.exe utility, which is used with administrator elevated access.

'Written By:    Christos Samaras
'Date:          01/06/2013
'Last Update:   15/02/2017
'E-mail:        xristos.samaras@gmail.com
'Site:          http://www.myengineeringworld.net
'------------------------------------------------------------------------------------------------
    
'Declaring the necessary ShellExecute API function for running the regsvr32 utility.
#If VBA7 And Win64 Then
    
    'For 64 bit Excel.
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                    (ByVal hwnd As LongPtr, _
                                    ByVal lpOperation As String, _
                                    ByVal lpFile As String, _
                                    ByVal lpParameters As String, _
                                    ByVal lpDirectory As String, _
                                    ByVal nShowCmd As Long) As LongPtr
#Else

    'For 32 bit Excel.
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                            (ByVal hwnd As Long, _
                             ByVal lpOperation As String, _
                             ByVal lpFile As String, _
                             ByVal lpParameters As String, _
                             ByVal lpDirectory As String, _
                             ByVal nShowCmd As Long) As Long
#End If



Public Sub RegisterDLL(DLLPath)

    '------------------------------------
    'Registers a DLL file on the system.
    '------------------------------------
    
    RegisterOrUnegisterDLL shMain.Range("C4").value, True

End Sub

Public Sub UnregisterDLL(DLLPath)
 
    '-----------------------------------------
    'Unregisters a DLL file from the system.
    '-----------------------------------------
       
    RegisterOrUnegisterDLL shMain.Range("C4").value, False

End Sub

Sub RegisterOrUnegisterDLL(DLLPath As String, Register As Boolean)
    
    '-----------------------------------------------------
    'Registers or unregisters a DLL file from the system.
    'Helping macro that tests first the given path.
    '-----------------------------------------------------

    'Check if the given file path exists.
    If FileExists(DLLPath) = False Then
        MsgBox "The DLL file doesn't exist!" & vbNewLine & "The file path you entered is incorrect.", vbCritical, "File Path Error"
        shMain.Range("C4").Select
        Exit Sub
    End If
    
    'Check if the given file path is a DLL file.
    If UCase(Right(DLLPath, 3)) <> "DLL" Then
        MsgBox "The file you have selected is not a Dynamic-link library (.dll) file!", vbCritical, "File Type Error"
        Exit Sub
    End If
    
    'Register or unregister the DLL file based on the Register variable.
    If Register = True Then
        Call RegisterFile(DLLPath)
    Else
        Call UnregisterFile(DLLPath)
    End If
    
End Sub

Sub RegisterFile(filePath As String)
    
    '--------------------------------------------------------------------
    'Registers a file on the system using administrator elevated access.
    '--------------------------------------------------------------------
    
    ShellExecute 0, "RunAs", "cmd", "/c regsvr32 " & """" & filePath & """", "C:\", 0
    
End Sub

Sub UnregisterFile(filePath As String)

    '------------------------------------------------------------------------
    'Unregisters a file from the system using administrator elevated access.
    '------------------------------------------------------------------------
    
    ShellExecute 0, "RunAs", "cmd", "/c regsvr32 /u " & """" & filePath & """", "C:\", 0
    
End Sub

Function FileExists(filePath As String) As Boolean
    
    '-------------------------
    'Checks if a file exists.
    '-------------------------
    
    If filePath <> vbNullString Then
        If Not Dir(filePath, vbDirectory) = vbNullString Then FileExists = True
    End If
    
End Function


Public Sub RegisterXLL(filePath)
    On Error GoTo ErrorHandler
    Application.RegisterXLL filePath
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.Description
End Sub

Public Sub UnregisterXLL(xllName)
    On Error GoTo ErrorHandler
    Application.ExecuteExcel4Macro "UNREGISTER(""" & xllName & """)"
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.Description
End Sub
()
