Attribute VB_Name = "Calcoli"

Option Base 0
Public Function ref_add_in_somma(A, b)
    ref_add_in_somma = xlwings.addin_somma(A, b)
End Function
Public Function ref_sec_info(ISIN)
    ref_sec_info = xlwings.fin_ch_sec_info(ISIN, "valoren")
End Function

Public Function SommaSeguentiColonna()
    Dim RigaCaller, UltimaRigaUtilizzata 'As Integer
    Application.Volatile
    UltimaRigaUtilizzata = Application.Caller.Parent.UsedRange.Cells.SpecialCells(xlCellTypeLastCell).Rows.Count
    RigaCaller = Application.Caller.Row
    ColonnaCaller = Application.Caller.Column
    On Error GoTo RifCirc
    SommaSeguentiColonna = Application.WorksheetFunction.Sum(Range(Cells(RigaCaller + 1, ColonnaCaller), Cells(UltimaRigaUtilizzata, ColonnaCaller)))
    Exit Function
RifCirc:
    SommaSeguentiColonna = "RifCirc"
End Function

Public Function SumSectionRows(SectionTitleColumn As Range, TopOutline As Boolean)
    Dim FormulaCell, FirstCell, FirstRow, LastRow, LastCell As Range
    Application.Volatile
    
On Error GoTo CircRef

    If TopOutline = True Then
        Set FirstCell = Application.Caller.Offset(1, 0)
        Debug.Print FirstCell.Address
        Set LastRow = Intersect(Application.Caller.EntireRow, SectionTitleColumn).End(xlDown).EntireRow.Offset(-1, 0)
        Set LastCell = Intersect(LastRow, FirstCell.EntireColumn)
        Debug.Print LastCell.Address
    Else
        Set LastCell = Application.Caller.Offset(-1, 0)
        Set FirstRow = Intersect(Application.Caller.EntireRow, SectionTitleColumn).End(xlUp).EntireRow.Offset(1, 0)
        Set FirstCell = Intersect(FirstRow, LastCell.EntireColumn)
    End If
    SumSectionRows = Application.Sum(Range(FirstCell, LastCell))
    Debug.Print SumSectionColumn
    Exit Function
CircRef:
    SumSectionRows = "CircRef"
End Function

Public Function SumOutlineRows(OutlineTitleColumn As Range, TopOutline As Boolean, Optional TableRange As Range) As Double
    Dim ValueToSum As Double
    Dim OutlineColumnCell, NextOutlineColumnCell As Range
    Application.Volatile
    
On Error GoTo CircRef

    If TableRange Is Nothing Then
        Set TableRange = Application.Caller.Parent.UsedRange
    End If

    Set ColumnToSum = Application.Caller.EntireColumn
    Set OutlineColumnCell = Intersect(Application.Caller.EntireRow, OutlineTitleColumn)
    SumOutlineRows = 0
    
    Do While Not Intersect(OutlineColumnCell, TableRange) Is Nothing
        If TopOutline = True Then
            Set NextOutlineColumnCell = OutlineColumnCell.End(xlDown)
        Else
            Set NextOutlineColumnCell = OutlineColumnCell.End(xlUp)
        End If
        ValueToSum = Intersect(NextOutlineColumnCell.EntireRow, ColumnToSum).Value2
        SumOutlineRows = SumOutlineRows + ValueToSum
        Set OutlineColumnCell = NextOutlineColumnCell
    Loop
    Exit Function
CircRef:
    SumSectionColumn = "CircRef"
End Function
Public Function SommaSeguentiRiga()
    Dim ColonnaCaller, UltimaColonnaUtilizzata 'As Integer
    'Application.Volatile
    UltimaColonnaUtilizzata = Application.Caller.Parent.UsedRange.Cells.SpecialCells(xlCellTypeLastCell).Columns.Count
    RigaCaller = Application.Caller.Row
    ColonnaCaller = Application.Caller.Column
    On Error GoTo RifCirc
    SommaSeguentiRiga = Application.WorksheetFunction.Sum(Range(Cells(RigaCaller, ColonnaCaller + 1), Cells(RigaCaller, UltimaColonnaUtilizzata)))
    Exit Function
RifCirc:
    SommaSeguentiRiga = "RifCirc"
End Function


Public Function Commissione(ImportoCHF As Currency, ValoriSoglie As Variant, ValoriCommissioni As Variant, Minimo As Variant, CumOrDir As String, PctOrAbs As String)
 Dim Soglia As Variant
 Dim Comm As Variant
 Dim Costo, Residuo, MaxPct, SogliaPrecedente As Currency
 Dim x As Integer
 
 ' se ValoriSoglia, Commissione e Minimo sono un Range, registrane i valori in una variabile
 If Information.IsObject(ValoriSoglie) Then
     ReDim Soglia(ValoriSoglie.Columns(1).Cells.Count - 1)
     x = 0
     For Each cella In ValoriSoglie.Columns(1).Cells
         Soglia(x) = cella.value
         x = x + 1
     Next cella
 Else
     ReDim Soglia(UBound(ValoriSoglie, 1) - 1)
     For x = 0 To UBound(ValoriSoglie, 1) - 1
         Soglia(x) = ValoriSoglie(x + 1, 1)
     Next x
 End If
 
 If Information.IsObject(ValoriCommissioni) Then
     ReDim Comm(ValoriCommissioni.Columns(1).Cells.Count - 1)
     x = 0
     For Each cella In ValoriCommissioni.Columns(1).Cells
         Comm(x) = cella.value
         x = x + 1
     Next cella
 Else
     ReDim Soglia(UBound(ValoriCommissioni, 1) - 1)
     For x = 0 To UBound(ValoriCommissioni, 1) - 1
         Comm(x) = ValoriCommissioni(x + 1, 1)
     Next x
 End If
 If IsObject(Minimo) Then
     Minimo = Minimo.value
 End If

'Calcolo della commissione
Costo = 0
SogliaPrecedente = 0
Residuo = ImportoCHF

For x = 0 To UBound(Soglia)
    If CumOrDir = "Cum" Then
        If ImportoCHF >= Soglia(x) Then
            Costo = Costo + (Soglia(x) - SogliaPrecedente) * Comm(x)
            SogliaPrecedente = Soglia(x)
            Residuo = ImportoCHF - SogliaPrecedente
        Else
            Costo = Costo + Residuo * Comm(x)
            x = UBound(Soglia)
        End If
    Else
        If Soglia(x) >= ImportoCHF Then
            Costo = ImportoCHF * Comm(x)
            x = UBound(Soglia)
        End If
    End If
Next x
'Applica il minomo se rilevante
If Costo < Minimo Then Costo = Minimo

'Calcola il percento se richiesto
If PctOrAbs = "Pct" Then Commissione = Costo / ImportoCHF Else Commissione = Costo
    
End Function
Sub TestCosti()
Costo "NP", 7000000, "DCust", "INT", True, False, "Gestione", "INT", "Classica"
End Sub

Public Function Costo(NomeFileTariffa As String, ImportoCHF As Currency, AbrTariffaOrForma As String, IntestazioneColonna As String, CumOrDir As String, PctOrAbs As String, Servizio As String, Domicilio As String, Optional TipoTariffa As String)

Application.Volatile

Dim Soglie As Variant
Dim Tassi As Variant
Dim Minima As Variant
Dim NomeFile, NomeFoglio, NomeRange As String
Dim Cartella, FileTariffe As Object
Dim AreaTariffe, FoglioTariffe As Object
Dim ColonnaTassi, ColonnaMinimo, RigaMinimo As Object


If Not IsNumeric(ImportoCHF) Then GoTo fine

NomeFile = "H:\Condizioni\" & NomeFileTariffa & ".xlsx"
NomeFoglio = Servizio

For Each Cartella In Workbooks   ' Esegue un'iterazione in ogni elemento.
    If Cartella.FullName = NomeFile Then
        Set FileTariffe = Cartella
        Set FoglioTariffe = Cartella.Worksheets(Servizio)
    End If
Next


With FoglioTariffe
'Nome Completo
'Se Forma a "" assume "Classica"
 If AbrTariffaOrForma = "" Then AbrTariffaOrForma = "Classica"
 
' Verifica in "decrescendo" l'esistenza delle varie tabelle di commissioni
 NomeRange = Servizio & "_" & AbrTariffaOrForma & "_" & TipoTariffa & "_" & Domicilio
 On Error Resume Next
 Set AreaTariffe = .Range(NomeRange)
 On Error GoTo 0
 If Not IsObject(AreaTariffe) Then
    'Nome Breve
    NomeRange = Servizio & "_" & AbrTariffaOrForma & "_" & Domicilio
    On Error Resume Next
    Set AreaTariffe = .Range(NomeRange)
    On Error GoTo 0
        If Not IsObject(AreaTariffe) Then
            'Range � 1 cella
            NomeRange = Domicilio & "_" & AbrTariffaOrForma
            On Error Resume Next
            Set AreaTariffe = .Range(NomeRange)
            Set Minima = .Range("Min_" & AbrTariffaOrForma)
            On Error GoTo 0
            If Not IsObject(AreaTariffe) Then
                GoTo NonTrovato
            End If
            
            'Verifica se la commissione � una percentuale o un valore assoluto (ovvero >1)
            If AreaTariffe.value < 1 Then
                If ImportoCHF * AreaTariffe > Minima Then
                    Costo = ImportoCHF * AreaTariffe
                Else
                    Costo = Minima
                End If
            Else
                Costo = AreaTariffe.value
            End If
            
            If PctOrAbs = "Pct" Then Costo = Costo / ImportoCHF
            GoTo fine
        End If
 End If
End With
Set Soglie = AreaTariffe.Resize(AreaTariffe.Rows.Count - 1, 1)
' Se non c� Profilo, assume il valore "Difensivo"
If IntestazioneColonna = "" Then IntestazioneColonna = "Difensivo"
Set ColonnaTassi = AreaTariffe.Columns(AreaTariffe.Rows(1).Offset(-1, 0).Find(What:=IntestazioneColonna, _
       LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
       MatchCase:=False, SearchFormat:=False).Column)
Set Tassi = ColonnaTassi.Resize(AreaTariffe.Rows.Count - 1, 1)
Set Minima = ColonnaTassi.Cells(ColonnaTassi.Rows.Count, 1)
Costo = Commissione(ImportoCHF, Soglie, Tassi, Minima, CumOrDir, PctOrAbs)
GoTo fine
NonTrovato:
   Costo = 0
fine:
End Function

Public Function ValPonderato(ValoreA As Variant, PesoA As Variant, ValoreB As Variant, PesoB As Variant, _
Optional Converti_Valori_Input_In_Pct_Loro_Totale As Boolean, Optional Converti_Pesi_Input_In_Pct_Loro_Totale As Boolean, Optional Converti_Risultati_In_Pct_Loro_Totale As Boolean, Optional SeErrore) As Variant

Dim x As Long
Dim Destinazione As Range
Dim Rige, Colonne
Dim ValorePonderato()
Dim ValP()
Dim ValA()
Dim ValB()
Dim WgtP()
Dim WgtA()
Dim WgtB()

If TypeName(Application.Caller) = "Range" Then
    With Application.Caller
        rCount = .Rows.Count
        cCount = .Columns.Count
        detinazione = .Address
    End With
End If

If IsObject(ValoreA) Then ValA() = WorksheetFunction.Transpose(ValoreA.Value2)
If IsObject(ValoreB) Then ValB() = WorksheetFunction.Transpose(ValoreB.Value2)
If IsObject(PesoA) Then WgtA() = WorksheetFunction.Transpose(PesoA.Value2)
If IsObject(PesoB) Then WgtB() = WorksheetFunction.Transpose(PesoB.Value2)

If (UBound(ValA) - LBound(ValA)) <> (UBound(WgtA) - LBound(WgtA)) Or (UBound(ValB) - LBound(ValB)) <> (UBound(WgtB) - LBound(WgtB)) _
    Or (UBound(ValA) - LBound(ValA)) <> (UBound(ValB) - LBound(ValB)) Then
    If SeErrore Then
        ValPonderato = SeErrore
        Exit Function
    End If
End If


If Converti_Valori_Input_In_Pct_Loro_Totale Then
   SumValoriA = Application.WorksheetFunction.Sum(ValA)
   SumValoriB = Application.WorksheetFunction.Sum(ValB)
   For x = LBound(ValA) To UBound(ValA)
       ValA(x) = ValA(x) / SumValoriA
   Next x
   For x = LBound(ValB) To UBound(ValB)
     ValB(x) = ValB(x) / SumValoriB
   Next x
End If

If Converti_Pesi_Input_In_Pct_Loro_Totale Then
    SumWgtA = Application.WorksheetFunction.Sum(WgtA)
    SumWgtB = Application.WorksheetFunction.Sum(WgtB)
    If SumWgtA <> 0 Then
        For x = LBound(WgtA) To UBound(WgtA)
            WgtA(x) = WgtA(x) / SumWgtA
        Next x
    End If
    If SumWgtB <> 0 Then
        For x = LBound(WgtB) To UBound(WgtB)
            WgtB(x) = WgtB(x) / SumWgtB
        Next x
    End If
End If

ValP = ValA
For x = LBound(ValP) To UBound(ValP)
    If (WgtA(x) + WgtB(x)) = 0 Then
        ValP(x) = 0
    Else
        ValP(x) = (ValA(x) * WgtA(x) + ValB(x) * WgtB(x)) / (WgtA(x) + WgtB(x))
    End If
Next x

If Converti_Risultati_In_Pct_Loro_Totale Then
    SumValP = Application.WorksheetFunction.Sum(ValP)
    If SumValP <> 0 Then
        For x = LBound(ValP) To UBound(ValP)
            ValP(x) = ValP(x) / SumValP
        Next x
    End If
End If

If IsObject(ValoreA) Then
'    Set
Else
    ValPonderato = ValP
End If
End Function



