'@Folder("MORFunctions.Modules")
Attribute VB_Name = "JsonHelper"
Option Explicit

Public Function SaveJsonFile(obj As Object, filePath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim jsonFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.CreateTextFile(filePath, True)

    jsonFile.Write ConvertToJson(obj)
    jsonFile.Close

    SaveJsonFile = True
    Exit Function

ErrorHandler:
    Debug.Print "Error in SaveJsonFile: " & Err.Number & " - " & Err.Description
    SaveJsonFile = False
End Function

Public Function ParseJsonFile(filePath As String) As Object
    Dim fso As Object
    Dim jsonFile As Object
    Dim jsonContent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fso.OpenTextFile(filePath, 1)
    jsonContent = jsonFile.ReadAll
    jsonFile.Close

    Set ParseJsonFile = ParseJson(jsonContent)
End Function

Private Function ConvertToJson(ByVal obj As Variant) As String
    Dim result As String
    Dim key As Variant
    Dim item As Variant

    Select Case VarType(obj)
        Case vbObject
            If TypeName(obj) = "Dictionary" Then
                result = "{"
                For Each key In obj.Keys
                    If Len(result) > 1 Then result = result & ","
                    result = result & """" & key & """:"
                    result = result & ConvertToJson(obj(key))
                Next key
                result = result & "}"
            ElseIf TypeName(obj) = "Collection" Then
                result = "["
                For Each item In obj
                    If Len(result) > 1 Then result = result & ","
                    result = result & ConvertToJson(item)
                Next item
                result = result & "]"
            Else
                ' Unsupported object type
                result = "null"
            End If
        Case vbString
            result = """" & Replace(obj, """", "\""") & """"
        Case vbBoolean
            result = IIf(obj, "true", "false")
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            result = obj
        Case vbNull
            result = "null"
        Case Else
            ' Unsupported type
            result = "null"
    End Select

    ConvertToJson = result
End Function
