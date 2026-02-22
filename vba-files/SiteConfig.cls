'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SiteConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SiteCredentials
    WindowsCredentialKey As String
End Type

Private Type SiteHeaders
    Collection As Collection
End Type

Private Type SiteCookie
    Name As String
    value As String
    Domain As String
    Path As String
    Secure As Boolean
    HttpOnly As Boolean
    ExpiresUTC As Date
End Type

Private Type SiteRequest
    URL As String
    Method As String
    Headers As SiteHeaders
    Params As Collection
End Type

Private mName As String
Private mBaseUrl As String
Private mCredentials As SiteCredentials
Private mLogin As SiteRequest
Private mDownload As SiteRequest
Private mCookies As Collection

Public Property Let Name(value As String)
    mName = value
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let BaseUrl(value As String)
    mBaseUrl = value
End Property

Public Property Get BaseUrl() As String
    BaseUrl = mBaseUrl
End Property

Public Property Let credentialKey(value As String)
    mCredentials.WindowsCredentialKey = value
End Property

Public Property Get credentialKey() As String
    credentialKey = mCredentials.WindowsCredentialKey
End Property

Public Sub InitFromJson(ByVal jsonObj As Object)
    Dim site As Object
    Set site = jsonObj
    
    Debug.Print "JsonObj type: " & TypeName(jsonObj)
    Debug.Print "site type: " & TypeName(site)
    
    If site.Exists("name") Then
        mName = site("name")
        Debug.Print "mName set to: " & mName
    Else
        Debug.Print "Key 'name' does not exist in site"
    End If
    
    If site.Exists("base_url") Then
        mBaseUrl = site("base_url")
        Debug.Print "mBaseUrl set to: " & mBaseUrl
    Else
        Debug.Print "Key 'base_url' does not exist in site"
    End If
    
    If site.Exists("credentials") Then
        If site("credentials").Exists("windows_credential_key") Then
            mCredentials.WindowsCredentialKey = site("credentials")("windows_credential_key")
            Debug.Print "mCredentials.WindowsCredentialKey set to: " & mCredentials.WindowsCredentialKey
        Else
            Debug.Print "Key 'windows_credential_key' does not exist in credentials"
        End If
    Else
        Debug.Print "Key 'credentials' does not exist in site"
    End If
    
    ' Initialize collections
    Set mCookies = New Collection
    
    ' Load login request config
    If site.Exists("login") Then
        With mLogin
            If site("login").Exists("url") Then
                .URL = site("login")("url")
                Debug.Print "mLogin.URL set to: " & .URL
            Else
                Debug.Print "Key 'url' does not exist in login"
            End If
            
            If site("login").Exists("method") Then
                .Method = site("login")("method")
                Debug.Print "mLogin.Method set to: " & .Method
            Else
                Debug.Print "Key 'method' does not exist in login"
            End If
            
            Set .Headers.Collection = New Collection
            Set .Params = New Collection
        End With
    Else
        Debug.Print "Key 'login' does not exist in site"
    End If
    
    ' Load download request config
    If site.Exists("download") Then
        With mDownload
            If site("download").Exists("url") Then
                .URL = site("download")("url")
                Debug.Print "mDownload.URL set to: " & .URL
            Else
                Debug.Print "Key 'url' does not exist in download"
            End If
            
            If site("download").Exists("method") Then
                .Method = site("download")("method")
                Debug.Print "mDownload.Method set to: " & .Method
            Else
                Debug.Print "Key 'method' does not exist in download"
            End If
            
            Set .Headers.Collection = New Collection
            Set .Params = New Collection
        End With
    Else
        Debug.Print "Key 'download' does not exist in site"
    End If
End Sub

Public Function ToJson() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    With result
        .Add "name", mName
        .Add "base_url", mBaseUrl
        
        Dim creds As Object
        Set creds = CreateObject("Scripting.Dictionary")
        creds.Add "windows_credential_key", mCredentials.WindowsCredentialKey
        .Add "credentials", creds
        
        ' Add other properties as needed
    End With
    
    Set ToJson = result
End Function
