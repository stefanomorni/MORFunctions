Attribute VB_Name = "Trasformazione"
Function HexToDec(hexStr$)
'VBA function to convert a hexadecimal string into a decimal. Uses
'the Decimal Variant subtype to allow for much larger values than the
'worksheet function HEX2DEC() or CLng("&h" & HexVal) can achieve. It
'returns an UNSIGNED Decimal integer...
'source:https://www.codegrepper.com/code-examples/vb/excel+vba+Convert+HEX+string+to+Unsigned+INT

    Dim c&, i&, bitVal
    Static hexVals&(), b() As Byte
    
    If (Not Not hexVals) = 0 Then
        ReDim hexVals(48 To 102)
        For i = 48 To 57
            hexVals(i) = c
            c = c + 1
        Next
        For i = 65 To 70
            hexVals(i) = c
            c = c + 1
        Next
        c = 10
        For i = 97 To 102
            hexVals(i) = c
            c = c + 1
        Next
    End If
    
    bitVal = CDec(1)
    HexToDec = CDec(0)
    b = StrConv(hexStr, vbFromUnicode)
    For i = UBound(b) To 0 Step -1
        HexToDec = HexToDec + hexVals(b(i)) * bitVal
        bitVal = bitVal * 16&
    Next
End Function

Public Function Symbol(Unicode_Hex_String As String)
Dim dec_code As Long
dec_code = Application.WorksheetFunction.Hex2Dec(Unicode_Hex_String)
Symbol = ChrW(dec_code)
End Function

Public Function CdblOrNull(Exp As Variant) As Variant
If IsNull(Exp) Then
    CdblOrNull = Null
Else
    CdblOrNull = CDbl(Exp)
End If
If CdblOrNull = 0 Then CdblOrNull = Null
End Function
Public Function Null2NA(Exp As Variant) As Variant
If Not IsNull(Exp) Then
    Null2NA = Exp
Else
    Null2NA = CVErr(xlErrNA)
End If
End Function

Public Function SortCollection(Coll As Collection, Optional Ascending As Boolean = True) As Collection
    Dim vItm As Variant
    Dim i As Long, j As Long
    Dim vTemp As Variant
    
    For i = 1 To Coll.Count - 1
        For j = i + 1 To Coll.Count
            'Ascending
            If Ascending Then
                If Coll(i) > Coll(j) Then
                    'store the lesser item
                    vTemp = Coll(j)
                    'remove the lesser item
                    Coll.Remove j
                    're-add the lesser item before the
                    'greater Item
                    Coll.Add vTemp, vTemp, i
                End If
            Else 'Descending
                If Coll(i) < Coll(j) Then
                    'store the greater item
                    vTemp = Coll(j)
                    'remove the greater item
                    Coll.Remove j
                    're-add the lesser item before the
                    'greater Item
                    Coll.Add vTemp, vTemp, i
                End If
            End If
        Next j
    Next i
Set SortCollection = Coll
End Function

Public Function ScrP(Numero As Long)
Dim strScribbled As String
Dim StrTemp As String
Dim Intero As Long
Dim x As Long


strSribbled = ""
StrTemp = Right(str(Numero), 7)
For x = 1 To Len(StrTemp)
    Intero = CSng(Mid(StrTemp, x, 1))
    strScribbled = strScribbled & Chr(77 - Intero)
Next x
ScrP = strScribbled
End Function

Public Function UnscrP(stringa As String)
Dim x As Integer
Dim strNumero As String
Dim strIntero As String
strNumero = str(10)
For x = 1 To Len(stringa)
    strIntero = 77 - CStr(Mid(stringa, x, 1))
    strNumero = strNumero & strIntero
Next x
ScrP = CSng(strNumero)
End Function

Public Function Num2Apsys(Oggetto As Range) As String
Num2Apsys = Left(Oggetto, 2) & "." & Mid(Oggetto, 3, 6) & "_" & Right(Oggetto, 1)
End Function

Public Function Apsys2Num(Oggetto As Range)
Apsys2Num = Int(Left(Oggetto, 2) & Mid(Oggetto, 4, 6) & Right(Oggetto, 1))
End Function

Public Function Txt2Num(testo)
Dim stringa As String
stringa = Trim(testo)
stringa = Replace(stringa, "'", "")
If Right(stringa, 1) = "%" Then
    stringa = Left(stringa, (Len(stringa) - 1))
    Txt2Num = CDbl(stringa) / 100
Else
    Txt2Num = CDbl(stringa)
End If
End Function
Public Function Points2Cm(ByVal pnt As Double) As Double
    Points2Cm = pnt * 0.03527778
End Function
Public Function Cm2Points(ByVal cm As Double) As Double
    Cm2Points = cm * 28.34646
End Function

Public Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Stream_StringToBinary(text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function



'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string

Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

Function Range2Array(Source_Range) As Variant
Dim Data_Array()
RangeToArray = WorksheetFunction.Transpose(Source_Range.Value2)
End Function




Public Function encodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument60
    
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    encodeBase64 = objNode.text
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function
 
      
 
Public Function DecodeBase64(ByVal strData As String) As Byte()
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function
 
Public Function bin2Byte(ByVal S As String) As Byte()
    Dim bitsIn As Long
    bitsIn = 8
 
    Dim i As Long
    'pad with zeros
    If Len(S) Mod bitsIn <> 0 Then
        For i = 1 To bitsIn - Len(S) Mod bitsIn
            S = "0" & S
        Next i
    End If
     
    i = Len(S)
    Dim bytes() As Byte
    Dim byteCount As Long
    byteCount = -1
    Dim sByte As String
    Do While LenB(S) > 0
        byteCount = byteCount + 1
        ReDim Preserve bytes(byteCount)
         
        sByte = Mid$(S, Len(S) - bitsIn + 1)
        'sByte = Mid$(s, 1, bitsIn)
        For i = 0 To 7 Step 1
            bytes(byteCount) = bytes(byteCount) + CLng(Mid$(sByte, 8 - i, 1)) * 2 ^ i
        Next i
        S = Mid$(S, 1, Len(S) - bitsIn)
        's = Mid$(s, bitsIn + 1)
    Loop
    bin2Byte = bytes
End Function
 
Public Function byte2Bin(ByRef bytes() As Byte) As String
    Dim i As Long, j As Long
    Dim bin As String
    For i = 0 To UBound(bytes)
        bin = Space$(8)
         
        For j = 0 To 7
            If bytes(i) And 2 ^ j Then
                Mid(bin, 8 - j, 1) = "1"
            Else
                'Mid(bin, 8 - j, 1) = "0"
            End If
        Next j
         
        byte2Bin = bin & byte2Bin
    Next i
    byte2Bin = LTrim$(byte2Bin)
    byte2Bin = Replace(byte2Bin, " ", "0", 1, -1, vbBinaryCompare)
End Function
 
Private Sub tester()
    'note we can't add any 0 padding to the test binary string
    Dim bin As String
    bin = "111101000001100010101"
    Dim binOut As String
    binOut = byte2Bin(DecodeBase64(encodeBase64(bin2Byte(bin))))
     
    MsgBox binOut = bin
End Sub

