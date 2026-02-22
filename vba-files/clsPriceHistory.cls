'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPriceHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Symbol As String
Public Fields As Collection
Public Quotes As Collection
Private Sub Class_Initialize()
    Set Fields = New Collection
    Set Quotes = New Collection
End Sub
