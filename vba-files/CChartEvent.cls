'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CChartEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
' Declare object of type "Chart" with events
Public WithEvents EventChart As Chart
Attribute EventChart.VB_VarHelpID = -1
 
Private Sub EventChart_Select(ByVal ElementID As Long, _
        ByVal Arg1 As Long, ByVal Arg2 As Long)
 
    MsgBox "Element: " & ElementID & vbCrLf & "  Arg 1: " & Arg1 _
        & vbCrLf & "  Arg 2: " & Arg2
 
End Sub
