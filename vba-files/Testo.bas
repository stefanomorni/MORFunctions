Attribute VB_Name = "Testo"
Public Function EstraiValoreDaStringa(testo As String, delimitatore As String, indice As Integer) As String
    Dim parti() As String
    parti = Split(testo, delimitatore)
    If indice >= LBound(parti) And indice <= UBound(parti) Then
        EstraiValoreDaStringa = parti(indice)
    Else
        EstraiValoreDaStringa = CVErr(xlErrNA)
    End If
End Function
Public Function NomeSenzaTipoFile(Oggetto As Object) As String

NomeSenzaTipoFile = Left(Oggetto.Name, Len(Oggetto.Name) - 4)

End Function


Public Function StringaSenzaSpazi(testo As String) As String

StringaSenzaSpazi = VBA.Strings.Replace(testo, " ", "")

End Function
Public Function RegExMatch(Text_To_Search, RegExPattern, Optional Index As Integer = 0)
'    Dim html As Object, objResult As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = RegExPattern
    RegEx.Global = True
    If Index < 0 Then Index = 0
    If RegEx.Test(Text_To_Search) Then
        Set matches = RegEx.Execute(Text_To_Search)
        RegExMatch = matches(Index).SubMatches(0)
    Else
        RegExMatch = ""
    End If
End Function

Public Function RegexCountMatches(str As String, reg As String) As String

'Returns the number of matches found for a given regex
'str - string to test the regex on
'reg - the regular expression

On Error GoTo ErrHandl
Set RegEx = CreateObject("VBScript.RegExp"): RegEx.Pattern = reg: RegEx.Global = True
If RegEx.Test(str) Then
    Set matches = RegEx.Execute(str)
    RegexCountMatches = matches.Count
    Exit Function
End If
ErrHandl:
    RegexCountMatches = CVErr(xlErrValue)
End Function
Public Function RegexExecute(str As String, reg As String, _
                             Optional matchIndex As Long, _
                             Optional subMatchIndex As Long) As String
                             
'Executes a Regular Expression on a provided string and returns a selected submatch
'str - string to execute the regex on
'reg - the regular expression with at least 1 capture '()'
'matchIndex - the index of the match you want to return (default: 0)
'subMatchIndex - the index of the submatch you want to return (default: 0)

On Error GoTo ErrHandl
Set RegEx = CreateObject("VBScript.RegExp"): RegEx.Pattern = reg
RegEx.Global = Not (matchIndex = 0 And subMatchIndex = 0) 'For efficiency
If RegEx.Test(str) Then
    Set matches = RegEx.Execute(str)
    RegexExecute = matches(matchIndex).SubMatches(subMatchIndex)
    Exit Function
End If
ErrHandl:
    RegexExecute = CVErr(xlErrValue)
End Function

Public Sub DumpText2File(TextData As String, FileFullPath As String)

Dim filename As String, textRow As String, fileNo As Integer

fileNo = FreeFile 'Get first free file number
    
Open FileFullPath For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
Print #fileNo, TextData
Close #fileNo

End Sub



Public Function Text_Before_Separator(ByVal txt As String, ByVal separator As String) As String
    If InStr(txt, separator) > 0 Then Extract_string = Split(txt, separator)(0)
End Function








