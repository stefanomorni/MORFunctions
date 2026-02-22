'@Folder("MORFunctions.Modules")
Attribute VB_Name = "Informational"
Option Explicit
Public Function FileExistsInt(filePath As String) As Integer
        Application.Volatile
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If filePath <> "" And fso.FileExists(filePath) Then
            FileExistsInt = 1
        Else
            FileExistsInt = 0
        End If
        Set fso = Nothing
End Function
Public Function GetURL(pWorkRng As Range) As String
    GetURL = pWorkRng.Hyperlinks(1).Address
End Function

Public Function NPRU(ByRef CellaFoglio As Range) As Long 'Numero Primariga utilizzata
NPRU = CellaFoglio.Parent.UsedRange.Cells(1).Row
End Function
Public Function NURU(ByRef CellaFoglio As Range) As Long 'Numero Ultima riga utilizzata
NURU = CellaFoglio.Parent.UsedRange.SpecialCells(xlCellTypeLastCell).Rows.Count + NPRU(CellaFoglio) - 1
End Function
Public Function ACUR(ByRef cella As Range) As Range 'Area da cella a ultima riga
Dim UltimaRiga As Long
UltimaRiga = NURU(cella)
Set ACUR = Range(cella, cella.Offset(UltimaRiga - cella.Row, 0))
End Function
Public Function ICUR(ByRef cella As Range) As String ' Indirizzo da cella a Ultima riga
Dim UltimaRiga As Long
UltimaRiga = NURU(cella)
ICUR = Range(cella, cella.Offset(UltimaRiga - cella.Row, 0)).Address(external:=True)
End Function

Public Function RDUR(NumeroRighe As Long, Optional ByRef CellaFoglio As Range) As Range ' Riga da ultima riga
Dim NumeroRiga As Long
If CellaFoglio Is Nothing Then Set CellaFoglio = Cells(1)
NumeroRiga = NURU(CellaFoglio) + NumeroRighe
Set RDUR = CellaFoglio.Parent.Rows(NumeroRiga)
End Function
Public Function TestoUltimaCellaColonnaContigua(ByRef PrimaCella As Range)
If PrimaCella.End(xlDown).NumberFormat = "General" Then
    TestoUltimaCellaColonnaContigua = PrimaCella.End(xlDown)
Else
    TestoUltimaCellaColonnaContigua = Format(PrimaCella.End(xlDown), PrimaCella.End(xlDown).NumberFormat)
End If
End Function
'Public Function CellaQuestaRigaColonnaRicerca(ByVal RangeRicerca As Range, ByVal TestoRicerca As String) As Range
'Dim Colonna, Riga As Integer
'Dim Foglio As Worksheet
'Dim Trovata As Range
'On Error GoTo Fine
'Set Foglio = Application.Caller.Parent
'Riga = Application.Caller.Row
'Set Trovata = Foglio.Range(RangeRicerca.Address).Cells.Find(What:=TestoRicerca, After:=RangeRicerca.Cells(1, 1), LookIn:=xlValues, LookAt _
'        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'        False, SearchFormat:=False)
'Colonna = Trovata.Column
'Set CellaQuestaRigaColonnaRicerca = Foglio.Cells(Riga, Colonna)
'Fine:
'End Function


Public Function RangeColonnaArea(ByRef NomeArea As String) As Range
Dim ColonnaCaller As Range
Application.Volatile
Set RangeColonnaArea = Intersect(Range(Names(NomeArea).RefersTo), Columns(Application.Caller.Column))
End Function

Public Function V(ByRef Rif As Range) As Variant 'ritorna una matrice Bolean che dice se la cella ï¿½ visibile o meno
Application.Volatile
Dim r, c As Integer
Dim Riga As Object
Dim Colonna As Object
Dim Visibile()
ReDim Visibile(Rif.Rows.Count - 1, Rif.Columns.Count - 1)

r = 0
For Each Riga In Rif.Rows.EntireRow
    c = 0
    For Each Colonna In Rif.Columns.EntireColumn
        If Colonna.Hidden = False And Riga.Hidden = False Then 'visibile
            Visibile(r, c) = True
        Else
            Visibile(r, c) = False
        End If
        c = c + 1
    Next Colonna
    r = r + 1
Next Riga
V = Visibile
End Function


Public Function WorksheetDaNome(ByVal File As Object, NomeFoglio As String) As Object

Dim Foglio As Object

For Each Foglio In File.Worksheets
    If Foglio.Name = NomeFoglio Then
        Set WorksheetDaNome = Foglio
        Exit Function
    Else
        Set WorksheetDaNome = Nothing
    End If
Next Foglio

End Function

Public Function SheetExists(Wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = Wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function

Public Function WorkbookDaNome(NomeCartella As String) As Object

Dim Cartella As Object

For Each Cartella In Application.Workbooks
    If Cartella.Name = NomeCartella Then
        Set WorkbookDaNome = Cartella
        Exit Function
    Else
        Set WorkbookDaNome = Nothing
    End If
Next Cartella

End Function

Public Function FoglioDaCodename(ByVal WorksheetCodename As String, Optional Wb As Workbook) As Worksheet
    Dim iSheet As Long
    If Wb Is Nothing Then Set Wb = ThisWorkbook ' mimics the default behaviour
    For iSheet = 1 To Wb.Worksheets.Count
        If Wb.Worksheets(iSheet).CodeName = WorksheetCodename Then
            Set FoglioDaCodename = Wb.Worksheets(iSheet)
            Exit Function
        End If
    Next iSheet
End Function

Public Function LibroDaCodename(ByVal WorkbookCodename) As Workbook
Dim iBook As Long
For iBook = 1 To Workbooks.Count
    If Workbooks(iBook).CodeName = WorkbookCodename Then
        Set LibroDaCodename = Workbooks(iBook)
        Exit Function
    End If
Next iBook
End Function

Public Function LarghezzaAreaStampa(ByVal ws As Worksheet, ByVal Unit As String) As Double 'in pt o cm
Dim LargPg, AltPg, ptLargAS As Double
With ws.PageSetup
If .PaperSize = xlA5 Then
    If .Orientation = xlLandscape Then
        LargPg = 842
        AltPg = 595
    Else
        LargPg = 595
        AltPg = 842
    End If
ElseIf .PaperSize = xlA3 Then
    If .Orientation = xlLandscape Then
        LargPg = 1190
        AltPg = 842
    Else
        LargPg = 842
        AltPg = 1190
    End If
Else ' xlA4

    If .Orientation = xlLandscape Then
        LargPg = 842
        AltPg = 595
    Else
        LargPg = 595
        AltPg = 842
    End If
End If

ptLargAS = LargPg - .LeftMargin - .RightMargin

If Unit = "pt" Or Unit = "Pt" Or Unit = "PT" Then
    LarghezzaAreaStampa = ptLargAS
ElseIf Unit = "cm" Or Unit = "cm" Or Unit = "CM" Then
    LarghezzaAreaStampa = Points2Cm(ptLargAS)
Else
    LarghezzaAreaStampa = 0
End If
End With
End Function

Public Function URLFromHyperlink(Rg As Object) As Variant

Dim URL()
Dim x As Integer 'counter
Dim cl As Range

x = 0
ReDim URL((Rg.Columns(1).Cells.Count - 1))
For Each cl In Rg.Columns(1).Cells
    If IsError(cl.Hyperlinks(1).Address) Then
        URL(x) = ""
    Else
        URL(x) = cl.Hyperlinks(1).Address
    End If
    x = x + 1
Next cl
URLFromHyperlink = Application.Transpose(URL)
End Function

Public Function InStr(stringa, Testo_Ricercato) As Boolean
'Ritorna Vero o Falso a seconda se il la string ricercata ï¿½ prsente o meno nel testo
If VBA.InStr(stringa, Testo_Ricercato) = 0 Then
    InStr = False
Else
    InStr = True
End If
End Function

Function FindAll(What, Optional SearchWhat As Variant, _
Optional LookIn, _
Optional LookAt, _
Optional SearchOrder, _
Optional SearchDirection As XlSearchDirection = xlNext, _
Optional MatchCase As Boolean = False, _
Optional MatchByte, _
Optional SearchFormat) As Range
       
'LookIn can be xlValues or xlFormulas, _
 LookAt can be xlWhole or xlPart, _
 SearchOrder can be xlByRows or xlByColumns, _
 SearchDirection can be xlNext, xlPrevious, _
 MatchCase, MatchByte, and SearchFormat can be True or False. _
 Before using SearchFormat = True, specify the appropriate settings _
 for the Application.FindFormat object, e.g., _
 Application.FindFormat.NumberFormat = "General;-General;""-"""
 
Dim aRng As Range

If IsMissing(SearchWhat) Then
    On Error Resume Next
    Set aRng = ActiveSheet.UsedRange
    On Error GoTo 0
ElseIf TypeOf SearchWhat Is Range Then
    If SearchWhat.Cells.Count = 1 Then
        Set aRng = SearchWhat.Parent.UsedRange
    Else
        Set aRng = SearchWhat
    End If
ElseIf TypeOf SearchWhat Is Worksheet Then
    Set aRng = SearchWhat.UsedRange
Else
    Exit Function                       '*****
End If

If aRng Is Nothing Then Exit Function   '*****

Dim FirstCell As Range, CurrCell As Range

With aRng.Areas(aRng.Areas.Count)
    Set FirstCell = .Cells(.Cells.Count)
    'This little 'dance' ensures we get the first matching _
     cell in the range first
End With
Set FirstCell = aRng.Find(What:=What, After:=FirstCell, _
    LookIn:=LookIn, LookAt:=LookAt, _
    SearchDirection:=SearchDirection, MatchCase:=MatchCase, _
    MatchByte:=MatchByte, SearchFormat:=SearchFormat)

If FirstCell Is Nothing Then Exit Function          '*****

Set CurrCell = FirstCell
Set FindAll = CurrCell
Do
    Set FindAll = Application.Union(FindAll, CurrCell)
    'Setting FindAll at the top of the loop ensures _
     the result is arranged in the same sequence as _
     the  matching cells; the duplicate assignment of _
     the first matching cell to FindAll being a small _
     price to pay for the ordered result
    Set CurrCell = aRng.Find(What:=What, After:=CurrCell, _
        LookIn:=LookIn, LookAt:=LookAt, _
        SearchDirection:=SearchDirection, MatchCase:=MatchCase, _
        MatchByte:=MatchByte, SearchFormat:=SearchFormat)
    'FindNext is not reliable because it ignores the FindFormat settings
Loop Until CurrCell.Address = FirstCell.Address
End Function

Public Function CurrentPath() As String
CurrentPath = Application.Caller.Parent.Parent.Path
End Function

Function ActiveTable(Target As Object) As Object

Dim TableName As String


'Determine if ActiveCell is inside a Table
  On Error GoTo NoTableSelected
    TableName = Target.ListObject.Name
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
  On Error GoTo 0
Exit Function

'Error Handling
NoTableSelected:
Set ActiveTable = Nothing
End Function

Public Function GetFileNamesbyExt(ByVal FolderPath As String, FileExt As String) As Variant
Dim result As Variant
Dim i As Integer
Dim MyFile As Object
Dim MyFSO As Object
Dim MyFolder As Object
Dim MyFiles As Object
Set MyFSO = CreateObject("Scripting.FileSystemObject")
Set MyFolder = MyFSO.GetFolder(FolderPath)
Set MyFiles = MyFolder.Files
ReDim result(1 To MyFiles.Count)
i = 1
For Each MyFile In MyFiles
If InStr(1, MyFile.Name, FileExt) <> 0 Then
result(i) = MyFile.Name
i = i + 1
End If
Next MyFile
ReDim Preserve result(1 To i - 1)
GetFileNamesbyExt = result
End Function

Public Function adrPivotTable(ByVal pTbName As String) As String
Application.Volatile
Dim Foglio, Tabella As Object
Set Foglio = Application.Caller.Parent
Set Tabella = Foglio.PivotTables(pTbName)

adrPivotTable = Tabella.TableRange1.Address


End Function


Public Function ValTable(Table As Range, ValueToSearchFor As Variant, ColumnToSearchIn As Range, ColumnToReturnValueFrom As Range, Optional SeErrore As Variant) As Variant
Dim ColumnToSearchFor
Application.Volatile
On Error GoTo Errore
If IsObject(ValueToSearchFor) Then ValueToSearchFor = ValueToSearchFor.Cells(1, 1).text
If IsObject(ColumnToReturnValueFrom) Then ColumnToSearchFor = ColumnToReturnValueFrom.Offset(-1, 0).Cells(1, 1).value

Dim Riga, Colonna As Integer

With Application.WorksheetFunction
    Riga = .Match(ValueToSearchFor, ColumnToSearchIn, 0)
    Colonna = .Match(ColumnToSearchFor, Table.Rows(1).Offset(-1, 0), 0)
    ValTable = .Index(Table, Riga, Colonna)
End With
Exit Function
Errore:
If IsMissing(SeErrore) Then
    ValTable = CVErr(xlErrNA)
Else
    ValTable = SeErrore
End If
On Error GoTo 0
End Function

Public Function LastModified(Fullpath As String, FormatoEng As String)
LastModified = Format(FileDateTime(Fullpath), FormatoEng)
End Function

Public Function CountInString(strIn As Variant, SplitChr As String, Optional SeZero As Variant) As Variant
Dim arrOut() As String

If IsError(SeZero) Then SeZero = Empty
If IsObject(strIn) Then strIn = strIn.Cells.text
arrOut = Split(strIn, SplitChr)
CountInString = UBound(arrOut) - LBound(arrOut) + 1
If CountInString = 0 Then
    If Not IsEmpty(SeZero) Then CountInString = SeZero
End If
End Function
Private Function SumInString(strIn As Variant, SplitChr As String, Optional SeZero As Variant) As Variant
Dim arrOut() As String
Dim strItem As Variant

If IsObject(strIn) Then strIn = strIn.Cells.text
If IsError(SeZero) Then SeZero = Empty
arrOut = Split(strIn, SplitChr)
SumInString = 0
For Each strItem In arrOut
    SumInString = SumInString + Val(strItem)
Next
If SumInString = 0 Then
    If Not IsEmpty(SeZero) Then SumInString = SeZero
End If
End Function

Public Function SpecialFolderPath(SpecialFolder As String) As String
'reference: https://renenyffenegger.ch/notes/Microsoft/dot-net/namespaces-classes/System/Environment/SpecialFolder

Dim Number As Variant

Select Case SpecialFolder
    Case "Desktop"
        Number = 0
    Case "Programs"
        Number = 2
    Case "Documenti"
        Number = 5
    Case "Favoriti"
        Number = 6
    Case "ApplicationData"
        Number = 26
    Case "LocalApplicationData"
        Number = 28
    Case "InternetCache"
        Number = 32
    Case "Cookies"
        Number = 33
    Case "CommonApplicationData"
        Number = 35
End Select
SpecialFolderPath = CreateObject("Shell.Application").Namespace(Number).Self.Path
End Function

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Private Function Test()
Debug.Print SpecialFolderPath("Desktop")
End Function

Function file_first_line(file_path As String)
Dim fso As FileSystemObject
Dim txtStream As TextStream
Dim strFilename As String
Dim strText As String

On Error GoTo NotFound
    strFilename = file_path
    Set fso = New FileSystemObject
    Set txtStream = fso.OpenTextFile(strFilename, ForReading, False)
    file_first_line = txtStream.ReadLine
    txtStream.Close
    Exit Function
NotFound:
    file_first_line = ""
End Function

Private Function get_first_line()
Debug.Print file_first_line("B:\Posizioni\FinCh.csv")
End Function

Public Function get_file_last_update(file_path As String)
get_file_last_update = FileDateTime(file_path)
End Function
