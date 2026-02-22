'@Folder("MORFunctions.Modules")
Attribute VB_Name = "TestSites"
Option Explicit


Public Sub TestEnvironment()
    #If Win64 Then
        Debug.Print "Running 64-bit"
    #Else
        Debug.Print "Running 32-bit"
    #End If
End Sub

Sub TestSiteManagement()
    ' Test 1: Load Configuration
    Dim site As SiteConfig
    Set site = LoadSiteConfiguration("lukb-ch")  ' Updated from "lukb"
    Debug.Print "Site loaded: " & (Not site Is Nothing)
    
    ' Test 2: Credentials
    Dim creds As Dictionary
    Set creds = GetSiteCredentials("lukb-ch")    ' Updated from "lukb_credentials"
    Debug.Print "Credentials found: " & (Not creds Is Nothing)
    If Not creds Is Nothing Then
        Debug.Print "Username: " & creds("username")
    End If
End Sub



Public Sub TestJsonHelper_Create()
    Dim testObj As Object
    Set testObj = CreateObject("Scripting.Dictionary")

    ' Create test site config
    With testObj
        .Add "websites", CreateObject("Scripting.Dictionary")
        With .item("websites")
            ' First site
            .Add "lukb-ch", CreateObject("Scripting.Dictionary")
            With .item("lukb-ch")
                .Add "name", "lukb-ch"
                .Add "base_url", "https://boersenundmaerkte.lukb.ch"
                .Add "credentials", CreateObject("Scripting.Dictionary")
                With .item("credentials")
                    .Add "windows_credential_key", "lukb-ch"
                End With
                .Add "login", CreateObject("Scripting.Dictionary")
                With .item("login")
                    .Add "url", "https://boersenundmaerkte.lukb.ch/login"
                    .Add "method", "POST"
                End With
                .Add "download", CreateObject("Scripting.Dictionary")
                With .item("download")
                    .Add "url", "https://boersenundmaerkte.lukb.ch/download"
                    .Add "method", "GET"
                End With
            End With

            ' Second site example
            .Add "finanzen-ch", CreateObject("Scripting.Dictionary")
            With .item("finanzen-ch")
                .Add "name", "finanzen-ch"
                .Add "base_url", "https://www.finanzen.ch"
                .Add "credentials", CreateObject("Scripting.Dictionary")
                With .item("credentials")
                    .Add "windows_credential_key", "finanzen-ch"
                End With
                .Add "login", CreateObject("Scripting.Dictionary")
                With .item("login")
                    .Add "url", "https://www.finanzen.ch/login"
                    .Add "method", "POST"
                End With
                .Add "download", CreateObject("Scripting.Dictionary")
                With .item("download")
                    .Add "url", "https://www.finanzen.ch/download"
                    .Add "method", "GET"
                End With
            End With
        End With
    End With

    ' Define the file path
    Dim filePath As String
    filePath = "D:\Cloud\OneDrive\Documenti\sites_config.json"

    Debug.Print "Attempting to save JSON file to: " & filePath
    ' Attempt to save the JSON file
    Dim result As Boolean
    result = SaveJsonFile(testObj, filePath)

    If result Then
        Debug.Print "SaveJsonFile returned True"
    Else
        Debug.Print "SaveJsonFile returned False"
    End If

    ' Verify if the file was created
    If Dir(filePath) <> "" Then
        Debug.Print "File exists at the specified location."

        ' Try to open and read the file
        Dim fileNum As Integer
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        Dim fileContent As String
        fileContent = Input$(LOF(fileNum), fileNum)
        Close #fileNum

        Debug.Print "File content (first 100 characters): " & Left$(fileContent, 100)
    Else
        Debug.Print "File does not exist at the specified location."
    End If
End Sub



Public Sub TestJsonHelper_Read()
    ' Test load only
    Dim loadedObj As Object
    Set loadedObj = ParseJsonFile("D:\Cloud\OneDrive\Documenti\sites_config.json")
    Debug.Print "Loaded sites:"
    Debug.Print "- " & loadedObj("websites")("lukb-ch")("base_url")
    Debug.Print "- " & loadedObj("websites")("finanzen-ch")("base_url")
End Sub


Public Sub TestSiteConfig()
    Dim site As New SiteConfig
    Dim jsonObj As Object
    Set jsonObj = CreateObject("Scripting.Dictionary")
    
    ' Create test site config
    With jsonObj
        .Add "name", "lukb-ch"
        .Add "base_url", "https://boersenundmaerkte.lukb.ch"
        
        Dim creds As Object
        Set creds = CreateObject("Scripting.Dictionary")
        creds.Add "windows_credential_key", "LUKB"
        .Add "credentials", creds
        
        Dim login As Object
        Set login = CreateObject("Scripting.Dictionary")
        login.Add "url", "https://boersenundmaerkte.lukb.ch/login"
        login.Add "method", "POST"
        .Add "login", login
        
        Dim download As Object
        Set download = CreateObject("Scripting.Dictionary")
        download.Add "url", "https://boersenundmaerkte.lukb.ch/download"
        download.Add "method", "GET"
        .Add "download", download
    End With
    
    ' Test initialization from JSON
    site.InitFromJson jsonObj
    Debug.Print "Site name: " & site.Name
    Debug.Print "Base URL: " & site.BaseUrl
    Debug.Print "Credential key: " & site.credentialKey
End Sub

Public Sub TestCredentials()
    Dim creds As Dictionary
    Set creds = GetSiteCredentials("lukb-ch")
    Debug.Print "Username: " & creds("username")
    Debug.Print "Password length: " & Len(creds("password"))
    
    Set creds = GetSiteCredentials("investing-com")
    Debug.Print "Username: " & creds("username")
    Debug.Print "Password length: " & Len(creds("password"))
End Sub
