'@Folder("MORFunctions.Modules")
Attribute VB_Name = "CSVConverter"
Public Function colonna_da_stringa(stringa As String, separatore)
colonna_da_stringa = Application.Transpose(VBA.Split(stringa, separatore))
End Function

Public Function riga_da_stringa(stringa As String, separatore)
riga_da_stringa = VBA.Split(stringa, separatore)
End Function

Public Function stringa_da_area(area As Object, separatore As String, Optional solo_celle_visibili As Boolean, Optional scarta_vuote As Boolean)
Application.Volatile

stringa_da_area = ""
vuota = False
prima = True

For Each cella In area.Cells
    nascosta = cella.EntireRow.Hidden Or cella.EntireColumn.Hidden
    If Trim(CStr(cella)) = "" Then vuota = True Else vuota = False
    If Not stringa_da_area = "" Then prima = False
    
    If vuota And scarta_vuote Then GoTo prossima
    If nascosta And solo_celle_visibili Then GoTo prossima
    If prima Then
        stringa_da_area = Trim(CStr(cella))
    Else
        stringa_da_area = stringa_da_area & "," & Trim(CStr(cella))
    End If
prossima:
Next cella
stringa_da_area = Trim(stringa_da_area)
End Function
