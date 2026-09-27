Attribute VB_Name = "Registro"

Function RegKeyRead(i_RegKey As String) As String
'Reading from the Registry:
'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function

Function RegKeyExists(i_RegKey As String) As Boolean
'Checking if a Registry key exists
'returns True if the registry key i_RegKey was found
'and False if not
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  'key was not found
  RegKeyExists = False
End Function

Sub RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
Dim myWS As Object
'Saving a Registry key
'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created

  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type

End Sub


Function RegKeyDelete(i_RegKey As String) As Boolean
'Deleting a key from the Registry:
'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'delete registry key
  myWS.RegDelete i_RegKey
  'deletion was successful
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  'deletion wasn't successful
  RegKeyDelete = False
End Function

