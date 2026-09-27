Attribute VB_Name = "Investing"
Option Explicit

Private Function BuildInvestingFormData(UserName As String, password As String) As String
    BuildInvestingFormData = _
        "-----------------------------37643986137383348913537276092" & vbNewLine & _
        "Content-Disposition: form-data; name=""email""" & vbNewLine & vbNewLine & _
        UserName & vbNewLine & _
        "-----------------------------37643986137383348913537276092" & vbNewLine & _
        "Content-Disposition: form-data; name=""password""" & vbNewLine & vbNewLine & _
        password & vbNewLine & _
        "-----------------------------37643986137383348913537276092" & vbNewLine & _
        "Content-Disposition: form-data; name=""logintoken""" & vbNewLine & vbNewLine & _
        "12345" & vbNewLine & _
        "-----------------------------37643986137383348913537276092" & vbNewLine & _
        "Content-Disposition: form-data; name=""includeTranslations""" & vbNewLine & vbNewLine & _
        "true" & vbNewLine & _
        "-----------------------------37643986137383348913537276092" & vbNewLine & _
        "Content-Disposition: form-data; name=""onAuthCompleteAction""" & vbNewLine & vbNewLine & _
        "{""type"":""topBar"",""location"":{""path"":""/"""",""search"":"""",""hash"":""""},""mmID"":1}" & vbNewLine & _
        "-----------------------------37643986137383348913537276092--"
End Function

Public Function DownloadInvesting()
    Dim Headers As New Collection
    Dim auth As New WebAuth
    
    ' Setup Investing.com headers
    Headers.Add Array("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:133.0) Gecko/20100101 Firefox/133.0")
    Headers.Add Array("Accept", "application/json")
    Headers.Add Array("Accept-Language", "it-CH,en;q=0.5")
    Headers.Add Array("X-Requested-With", "XMLHttpRequest")
    Headers.Add Array("Sec-GPC", "1")
    Headers.Add Array("Alt-Used", "www.investing.com")
    
    ' Configure auth object
    auth.LoginEndpoint = "https://www.investing.com/members-admin/auth/signInByEmail"
    Set auth.RequestHeaders = Headers
    auth.PostData = BuildInvestingFormData("morniroli.stefano@gmail.com", "inver68filtro")
    auth.PostDataType = "multipart/form-data; boundary=---------------------------37643986137383348913537276092"
    
    ' Download file
    DownloadInvesting = DownloadSecureFile( _
        "https://www.investing.com/portfolio/service/Export?portfolio_id=3083029&exptype=csv", _
        auth, _
        "D:\Cloud\OneDrive\Documenti", _
        "D:\Cloud\OneDrive\Documenti", _
        "investing_portfolio", _
        "csv")
End Function
