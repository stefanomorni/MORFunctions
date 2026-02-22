'@Folder("MORFunctions.Modules")
Attribute VB_Name = "SiteManagement"
Private Type Credential
    Flags As Long
    Type As Long
    TargetName As Long
    Comment As Long
    LastWritten As Currency
    CredentialBlobSize As Long
    CredentialBlob As Long
    Persist As Long
    AttributeCount As Long
    Attributes As Long
    TargetAlias As Long
    UserName As Long
End Type

Private Declare PtrSafe Function CredRead Lib "advapi32.dll" Alias "CredReadW" ( _
    ByVal TargetName As String, _
    ByVal CredType As Long, _
    ByVal Flags As Long, _
    ByRef Credential As Long) As Boolean

Private Declare PtrSafe Function CredFree Lib "advapi32.dll" ( _
    ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function lstrcpyW Lib "kernel32" ( _
    ByVal lpString1 As LongPtr, _
    ByVal lpString2 As LongPtr) As Long

Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As LongPtr) As Long

Private Const CONFIG_FILE_PATH As String = "D:\Cloud\OneDrive\Documenti\sites_config.json"

Public Function LoadSiteConfiguration(siteName As String) As SiteConfig
    Dim jsonData As Object
    Dim site As SiteConfig
    Set site = New SiteConfig
    
    Set jsonData = ParseJsonFile(CONFIG_FILE_PATH)
    If jsonData.Exists("websites") And jsonData("websites").Exists(siteName) Then
        Dim siteData As Object
        Set siteData = jsonData("websites")(siteName)
        Debug.Print "siteData type: " & TypeName(siteData)
        site.InitFromJson siteData
        Set LoadSiteConfiguration = site
    Else
        Set LoadSiteConfiguration = Nothing
    End If
End Function

Public Function UpdateSiteCookies(siteName As String, cookies As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    Dim jsonData As Object
    Set jsonData = ParseJsonFile(CONFIG_FILE_PATH)
    
    If Not jsonData("websites").Exists(siteName) Then
        jsonData("websites").Add siteName, CreateObject("Scripting.Dictionary")
    End If
    
    Dim siteData As Object
    Set siteData = jsonData("websites")(siteName)
    
    ' Update cookies
    If Not siteData.Exists("cookies") Then
        siteData.Add "cookies", CreateObject("Scripting.Dictionary")
    End If
    Set siteData("cookies") = cookies
    
    UpdateSiteCookies = SaveJsonFile(jsonData, CONFIG_FILE_PATH)
    Exit Function
    
ErrorHandler:
    UpdateSiteCookies = False
End Function

Public Function GetSiteCredentials(credentialKey As String) As Dictionary
    Dim cred As Long
    Dim credData As Credential
    Dim result As New Dictionary
    
    If CredRead(credentialKey, 1, 0, cred) <> 0 Then
        lstrcpyW VarPtr(credData), cred
        
        result.Add "username", PointerToString(credData.UserName)
        result.Add "password", PointerToString(credData.CredentialBlob)
        
        CredFree cred
    End If
    
    Set GetSiteCredentials = result
End Function

Private Function PointerToString(ByVal ptr As LongPtr) As String
    ' Convert pointer to string
    If ptr = 0 Then Exit Function
    PointerToString = Space$(lstrlenW(ptr))
    lstrcpyW StrPtr(PointerToString), ptr
End Function
