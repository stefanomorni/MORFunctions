'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public GMTTimeStamp As Long
Public TimeKey As String
Public TimeKeyDateTime As Date
Public IndexNo As Long
Public Fields As Collection
Private Sub Class_Initialize()
    Set Fields = New Collection
End Sub
