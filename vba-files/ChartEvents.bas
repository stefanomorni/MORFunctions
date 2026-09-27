Attribute VB_Name = "ChartEvents"
'----------------------------------------------------------------------------------------
' to use this fuctionality, you have to add following code to the workbook module
'-----------------------------------------------------------------------------------------
'Option Explicit
'
'Private Sub Workbook_Open()
'    InitializeAppEvents
'End Sub
'
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    TerminateAppEvents
'End Sub
'-------------------------------------------------------------------------------------------
'End of Workoob Module Code
'-------------------------------------------------------------------------------------------


Option Explicit

Dim clsAppEvent As New CAppEvent
Dim clsChartEvent As New CChartEvent
Dim clsChartEvents() As New CChartEvent

Sub InitializeAppEvents()
  Set clsAppEvent.EventApp = Application
  Set_All_Charts
End Sub
 
Sub TerminateAppEvents()
  Set clsAppEvent.EventApp = Nothing
  Reset_All_Charts
End Sub

Sub Set_All_Charts()
    ' Enable events for active sheet if sheet is a chart sheet
    If TypeName(ActiveSheet) = "Chart" Then
        Set clsChartEvent.EventChart = ActiveSheet
    End If
 
    ' Enable events for all charts embedded on a sheet
    ' Works for embedded charts on a worksheet or chart sheet
    If ActiveSheet.ChartObjects.Count > 0 Then
        ReDim clsChartEvents(1 To ActiveSheet.ChartObjects.Count)
        Dim chtObj As ChartObject
        Dim chtnum As Integer
 
        chtnum = 1
        For Each chtObj In ActiveSheet.ChartObjects
            ' Debug.Print chtObj.Name, chtObj.Parent.Name
            Set clsChartEvents(chtnum).EventChart = chtObj.Chart
            chtnum = chtnum + 1
        Next ' chtObj
    End If
End Sub
 
Sub Reset_All_Charts()
    ' Disable events for all charts previously enabled together
    Dim chtnum As Integer
    On Error Resume Next
    Set clsChartEvent.EventChart = Nothing
    For chtnum = 1 To UBound(clsChartEvents)
        Set clsChartEvents(chtnum).EventChart = Nothing
    Next ' chtnum
End Sub
