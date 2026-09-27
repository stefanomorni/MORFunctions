Attribute VB_Name = "Ambiente"

Public Sub BloccaXSecondi(Seconds)

' Wait a while blocking excel activity

'Declarations

Dim TheTime As Date

'Procedure

Let TheTime = Timer
Do
Loop Until (Timer - TheTime) >= Seconds

End Sub

' ' This procedure was used to ensure that ExcelDNA Intellisense would be loaded AFTER xloil
' ' That was needed, because loading xlOIL after ExcelDNA would crash excel.
' ' The problem hase been solved with xlöOil v. 0.15.3, so thios i not needed anymore

' Function Register_ExcelDna_IntelliSense64_xll()
' FullPath = "D:\Cloud\OneDrive\Office Junctions\AddIns\ExcelDna.IntelliSense64.xll"
' Debug.Print "Sto registrando " & FullPath
' ManageDLL.RegisterXLL (FullPath)
' Debug.Print "Ho registrato " & FullPath
' End Function

Function Debug_clear()
Debug.Print Now
Application.SendKeys "^g ^a {DEL}"
End Function

Sub CancelAllQueriesAndWait()
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim co As WorkbookConnection
    
    ' --- Cancel classic QueryTable refreshes (all sheets) ---
    For Each ws In ThisWorkbook.Worksheets
        For Each qt In ws.QueryTables
            If qt.Refreshing Then
                qt.CancelRefresh
            End If
        Next qt
    Next ws

    ' --- Cancel connection-based queries (Power Query, OLEDB, etc.) ---
    For Each co In ThisWorkbook.Connections
        On Error Resume Next ' Some connection types may not support .Refreshing
        If co.Refreshing Then
            co.CancelRefresh
        End If
        On Error GoTo 0
    Next co

    ' --- Wait until ALL are no longer refreshing ---
    Do
        Dim stillRefreshing As Boolean
        stillRefreshing = False

        ' Check QueryTables
        For Each ws In ThisWorkbook.Worksheets
            For Each qt In ws.QueryTables
                If qt.Refreshing Then
                    stillRefreshing = True
                    Exit For
                End If
            Next qt
            If stillRefreshing Then Exit For
        Next ws
        
        ' Check Connections
        If Not stillRefreshing Then
            For Each co In ThisWorkbook.Connections
                On Error Resume Next
                If co.Refreshing Then
                    stillRefreshing = True
                End If
                On Error GoTo 0
                If stillRefreshing Then Exit For
            Next co
        End If

        DoEvents ' Let Excel process cancellation internally
    Loop While stillRefreshing
End Sub




