'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLoginEndpoint As String
Private mRequestHeaders As Collection
Private mPostData As String
Private mPostDataType As String

Public Property Let LoginEndpoint(value As String)
    mLoginEndpoint = value
End Property

Public Property Get LoginEndpoint() As String
    LoginEndpoint = mLoginEndpoint
End Property

Public Property Set RequestHeaders(value As Collection)
    Set mRequestHeaders = value
End Property

Public Property Get RequestHeaders() As Collection
    Set RequestHeaders = mRequestHeaders
End Property

Public Property Let PostData(value As String)
    mPostData = value
End Property

Public Property Get PostData() As String
    PostData = mPostData
End Property

Public Property Let PostDataType(value As String)
    mPostDataType = value
End Property

Public Property Get PostDataType() As String
    PostDataType = mPostDataType
End Property
