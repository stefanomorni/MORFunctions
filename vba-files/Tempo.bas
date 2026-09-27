Attribute VB_Name = "Tempo"

Public Function Date2Timestamp(dtDate, Optional GMTOffsetSeconds = 0) As Long
If Not IsNull(dtDate) Then
    Date2Timestamp = (CDate(dtDate) - #1/1/1970#) * 86400 - GMTOffsetSeconds
End If
End Function

Public Function Timestamp2Date(lgTimestamp, Optional GMTOffsetSeconds = 0) As Date
If Not IsNull(lgTimestamp) Then
    Timestamp2Date = (lgTimestamp + GMTOffsetSeconds) / 86400 + #1/1/1970#
End If
End Function

Public Function DayOfWeek(DateRange As Date, Optional DayInWeek = 1) As Date 'Day in Week 1 is onday, 2, tuesday and so on
    DayOfWeek = DateRange - Weekday(DateRange, vbUseSystem) + DayInWeek
End Function

Public Function LastDayOfMonth(ByVal myDate As Date) As Date
    LastDayOfMonth = DateSerial(Year(myDate), Month(myDate) + 1, 0)
End Function



