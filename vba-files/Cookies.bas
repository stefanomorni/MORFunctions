'@Folder("MORFunctions.Modules")
Attribute VB_Name = "Cookies"
Option Explicit

' API Declarations
Private Declare PtrSafe Function CryptUnprotectData Lib "Crypt32.dll" ( _
    ByRef pDataIn As DATA_BLOB, _
    ByVal ppszDataDescr As Long, _
    ByVal pOptionalEntropy As Long, _
    ByVal pvReserved As Long, _
    ByVal pPromptStruct As Long, _
    ByVal dwFlags As Long, _
    ByRef pDataOut As DATA_BLOB _
) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long _
)

Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

Private Type DATA_BLOB
    cbData As Long
    pbData As LongPtr
End Type

Private Const CRYPTPROTECT_UI_FORBIDDEN As Long = &H1


' Main function to extract Edge cookies
Public Function ExtractEdgeCookies(Domain As String, outputDir As String) As Boolean
    On Error GoTo ErrorHandler

    Dim conn As Object
    Dim rs As Object
    Dim cookiesPath As String
    Dim outputPath As String
    Dim cookieData As String

    ' Get Edge cookie database path
    cookiesPath = GetEdgeCookieDbPath()
    If cookiesPath = "" Then
        Debug.Print "Edge cookie database not found."
        ExtractEdgeCookies = False
        Exit Function
    End If

    ' Prepare output path
    outputPath = PrepareOutputPath(outputDir, Domain)

    ' Open SQLite database using ADO
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Driver={SQLite3 ODBC Driver};Database=" & cookiesPath & ";"
    conn.Open

    ' Execute query to get cookies
    Set rs = conn.Execute("SELECT host_key, path, is_secure, expires_utc, name, encrypted_value FROM cookies WHERE host_key LIKE '%" & Domain & "%'")

    ' Process results
    Do While Not rs.EOF
        On Error Resume Next
        Dim cookieLine As String
        cookieLine = FormatCookieLine(rs)
        If Err.Number = 0 Then
            cookieData = cookieData & cookieLine & vbNewLine
        Else
            Debug.Print "Error processing cookie: " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        rs.MoveNext
    Loop

    ' Close database connection
    rs.Close
    conn.Close

    ' Save cookies to file
    If SaveCookiesToFile(cookieData, outputPath) Then
        ExtractEdgeCookies = True
    Else
        ExtractEdgeCookies = False
    End If

    Exit Function

ErrorHandler:
    Debug.Print "Error in ExtractEdgeCookies: " & Err.Description
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    ExtractEdgeCookies = False
End Function

Private Function GetEdgeCookieDbPath() As String
    Dim fso As Object
    Dim possiblePaths(1) As String
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")

    possiblePaths(0) = Environ$("LOCALAPPDATA") & "\Microsoft\Edge\User Data\Default\Network\Cookies"
    possiblePaths(1) = Environ$("LOCALAPPDATA") & "\Microsoft\Edge\User Data\Default\Cookies"

    For i = 0 To UBound(possiblePaths)
        If fso.FileExists(possiblePaths(i)) Then
            GetEdgeCookieDbPath = possiblePaths(i)
            Exit Function
        End If
    Next i

    GetEdgeCookieDbPath = "" ' Return empty string if not found
End Function

Private Function PrepareOutputPath(outputDir As String, Domain As String) As String
    Dim fso As Object
    Dim formattedDomain As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Format domain name for filename
    formattedDomain = Replace(Domain, ".", "-")

    ' Ensure output directory exists
    If Not fso.FolderExists(outputDir) Then
        fso.CreateFolder outputDir
    End If

    ' Construct full output path
    PrepareOutputPath = fso.BuildPath(outputDir, formattedDomain & ".txt")
End Function

Private Function FormatCookieLine(rs As Object) As String
    Dim Domain As String
    Dim isHttpOnly As String
    Dim cookiePath As String
    Dim isSecure As Boolean
    Dim ExpiresUTC As String
    Dim cookieName As String
    Dim cookieValue As String

    Domain = rs.Fields("host_key").value
    isHttpOnly = IIf(Left(Domain, 1) = ".", "#HttpOnly_" & Domain, Domain)
    cookiePath = rs.Fields("path").value
    isSecure = CBool(rs.Fields("is_secure").value)

    ' Handle expires_utc as a string to avoid overflow
    On Error Resume Next
    ExpiresUTC = rs.Fields("expires_utc").value
    If Err.Number <> 0 Then
        ExpiresUTC = "9999999999" ' Set to a far future date if there's an error
        Err.Clear
    Else
        ' Convert microseconds to seconds and adjust for epoch
        ExpiresUTC = CLng(CDec(ExpiresUTC) / 1000000# - 11644473600#)
    End If
    On Error GoTo 0

    cookieName = rs.Fields("name").value
    cookieValue = DecryptCookieValue(rs.Fields("encrypted_value").value)

    FormatCookieLine = isHttpOnly & vbTab & _
                       "TRUE" & vbTab & _
                       cookiePath & vbTab & _
                       IIf(isSecure, "TRUE", "FALSE") & vbTab & _
                       ExpiresUTC & vbTab & _
                       cookieName & vbTab & _
                       cookieValue
End Function

Private Function DecryptCookieValue(encryptedValue As Variant) As String
    On Error GoTo ErrorHandler
    Dim bytesIn() As Byte
    Dim strValue As String
    
    ' Convert encrypted value to byte array
    bytesIn = encryptedValue
    Debug.Print "Encrypted value length: " & UBound(bytesIn) - LBound(bytesIn) + 1
    Debug.Print "First few bytes: " & BytesToHex(bytesIn, 10)

    ' Convert byte array to string
    strValue = BytesToString(bytesIn)
    
    ' Check if the value starts with "v10" or "v20"
    If Left$(strValue, 3) = "v10" Or Left$(strValue, 3) = "v20" Then
        ' Remove the version prefix and return the decrypted value
        DecryptCookieValue = Mid$(strValue, 4)
    Else
        ' If it doesn't start with v10 or v20, return the original string
        DecryptCookieValue = strValue
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "Error in DecryptCookieValue: " & Err.Description
    DecryptCookieValue = ""
End Function

Private Function BytesToString(bytes() As Byte) As String
    Dim result As String
    Dim i As Long
    
    For i = LBound(bytes) To UBound(bytes)
        result = result & Chr$(bytes(i))
    Next i
    
    BytesToString = result
End Function

Private Function BytesToHex(bytes() As Byte, Optional maxBytes As Long = -1) As String
    Dim result As String
    Dim i As Long
    Dim limit As Long
    
    If maxBytes = -1 Or maxBytes > UBound(bytes) - LBound(bytes) + 1 Then
        limit = UBound(bytes)
    Else
        limit = LBound(bytes) + maxBytes - 1
    End If
    
    For i = LBound(bytes) To limit
        result = result & Right$("0" & Hex$(bytes(i)), 2) & " "
    Next i
    
    BytesToHex = Trim$(result)
End Function

Private Function SaveCookiesToFile(cookieData As String, outputPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim txtFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(outputPath, True)
    
    ' Write Netscape cookie file header
    txtFile.WriteLine "# Netscape HTTP Cookie File"
    txtFile.WriteLine "# http://curl.haxx.se/rfc/cookie_spec.html"
    txtFile.WriteLine "# This file was generated by VBA"
    
    ' Write cookie data
    txtFile.Write cookieData
    
    txtFile.Close
    SaveCookiesToFile = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SaveCookiesToFile: " & Err.Description
    SaveCookiesToFile = False
End Function

Public Sub TestExtractEdgeCookies()
    Dim Domain As String
    Dim outputDir As String
    Dim result As Boolean
    
    Domain = "investing.com"  ' Replace with a domain you want to test
    outputDir = "D:\Cloud\OneDrive\Cookies"  ' Replace with your desired output directory
    
    result = ExtractEdgeCookies(Domain, outputDir)
    
    If result Then
        Debug.Print "Cookies extracted successfully for " & Domain
        Debug.Print "Output file: " & outputDir & "\" & Replace(Domain, ".", "-") & ".txt"
    Else
        Debug.Print "Failed to extract cookies for " & Domain
    End If
End Sub
