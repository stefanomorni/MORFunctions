'@Folder("MORFunctions.Modules")
Attribute VB_Name = "Aggiornamenti"
Sub Aggiorna_Links_Immagini_To_Edit()

For Each immagine In ActiveSheet.Shapes
    On Error Resume Next
    Debug.Print immagine.Hyperlink.Address
    PreviousAddress = immagine.Hyperlink.Address = ""
    immagine.Hyperlink.Address
Next immagine
End Sub
