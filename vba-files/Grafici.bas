Attribute VB_Name = "Grafici"
Function MaxCht(ByVal Cht As Object)
Dim MaxSerie As Double
MaxCht = -1 * 10 ^ 20
For Each Serie In Cht.SeriesCollection
    MaxSerie = CDbl(Application.Aggregate(14, 6, rgValoriSerie(Serie), 1))
    If MaxSerie > MaxCht Then
        MaxCht = MaxSerie
    End If
Next Serie
End Function

Function MinCht(ByVal Cht As Object)
Dim MinSerie As Double
Dim Rg As Object
MinCht = 10 ^ 20
For Each Serie In Cht.SeriesCollection
    MinSerie = CDbl(Application.Aggregate(15, 6, rgValoriSerie(Serie), 1))
    If MinSerie < MinCht Then
        MinCht = MinSerie
    End If
Next Serie
End Function

Function rgSourceDataregion(ByVal Cht As Object) As Range
Set rgSourceDataregion = rgValoriSerie(Cht.SeriesCollection(1)).Cells(1, 1).CurrentRegion
End Function

Function rgValoriSerie(ByVal Serie As Object) As Range
Dim TestoFormula, Elementi
TestoFormula = Split(Serie.Formula, "(")
TestoFormula = Split(TestoFormula(1), ")")
Elementi = Split(TestoFormula(0), ",")
Set rgValoriSerie = Range(Elementi(2))
End Function

Function rgAsseXSerie(ByVal Serie As Object) As Range
Dim TestoFormula, Elementi
TestoFormula = Split(Serie.Formula, "(")
TestoFormula = Split(TestoFormula(1), ")")
Elementi = Split(TestoFormula(0), ",")
Set rgAsseXSerie = Range(Elementi(1))
End Function
Function rgNomeSerie(ByVal Serie As Object) As Range
Dim TestoFormula, Elementi
TestoFormula = Split(Serie.Formula, "(")
TestoFormula = Split(TestoFormula(1), ")")
Elementi = Split(TestoFormula(0), ",")
Set rgNomeSerie = Range(Elementi(0))
End Function

Function PlotOrderSerie(ByVal Serie As Object) As Integer
Dim TestoFormula, Elementi
TestoFormula = Split(Serie.Formula, "(")
TestoFormula = Split(TestoFormula(1), ")")
Elementi = Split(TestoFormula(0), ",")
PlotOrderSerie = Range(Elementi(3))
End Function

Public Function Sorted_SeriesCollection_On_Last_Point(Coll As SeriesCollection, Optional Ascending As Boolean = True) As SeriesCollection
    Dim vItm As Variant
    Dim i As Long, j As Long
    Dim sr_sorted As SeriesCollection
    Dim i_values, j_values As Variant
    Dim srTemp As Variant
    
    For i = 1 To Coll.Count - 1
        i_values = Coll(i).Values
        For j = i + 1 To Coll.Count
            j_values = Coll(j).Values
            'Ascending
            If Ascending Then
                If i_values(UBound(i_values)) > j_values(UBound(j_values)) Then
                    'store the lesser item
                    Set srTemp = Coll(i)
                    'remove the greater item
                    Coll(i).Delete
                    're-add the lesser item before the
                    'greater Item
                    Coll.Add srTemp
                End If
            Else 'Descending
                If i_values(UBound(i_values)) < j_values(UBound(j_values)) Then
                    'store the greater item
                    Set srTemp = Coll(j)
                    'remove the greater item
                    Coll(j).Delete
                    're-add the lesser item before the
                    'greater Item
                    Coll.Add srTemp
                End If
            End If
        Next j
    Next i
Set Sorted_SeriesCollection_On_Last_Point = Coll
End Function


