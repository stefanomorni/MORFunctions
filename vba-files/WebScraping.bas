Attribute VB_Name = "WebScraping"
Option Explicit

Sub Test_Yahoo()
'YahooHistory Array("6502.T"), "01.06.2020", "A4", "05.06.2020", Array("idx_refcy_adjclose"), 1, "1d", "CHF"
YahooHistory Array("NOVN.SW", "^SSMI"), "01.01.2019", , , , , , , "Test", "N6"
'YahooHistory Array("NOVN.SW"), "01.06.2020", "A4", "05.06.2020", Array("pct_close_startclose"), 1, "1d", , "^SSMI"
'YahooHistory Array("AAPL"), "01.06.2020", "A4", "05.06.2020", Array("pct_close_startclose"), 1, "1d"
'YahooHistory Array("CHF=X"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("^GSPC"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("^SSMI"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("^N225"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("ES=F"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("NIY=F"), "01.06.2020", "A4", "05.06.2020", Array("close"), 1, "1d"
'YahooHistory Array("^SSMI"), "01.01.2020", "D4", "08.05.2020", Array("pct_close_startclose"), 1, , "CHF"
'YahooHistory Array("^STOXX50E"), "01.01.2020", "G4", "08.05.2020", Array("pct_close_startclose"), 1, , "CHF"
'YahooHistory Array("^FTSE"), "01.01.2020", "J4", "08.05.2020", Array("pct_close_startclose"), 1, , "CHF"
'YahooHistory Array("^N225"), "01.01.2020", "M4", "08.05.2020", Array("pct_close_startclose"), 1, , "CHF"
End Sub



Function YahooHistory(Tickers As Variant, StartDate, Optional EndDate, Optional Fields = "Close", Optional Granularity As String = "1d", _
                    Optional RefCurrency As String = "", Optional Benchmark As String = "", _
                    Optional OutputFieldsLables As Boolean = True, Optional RangeName As String = "", Optional OutputFirstCell As Variant = "")
'(Symbol, StartDate, EndDate, Optional Fields = Array("adjClose"), Optional Granularity = "d")

'Yahoo finance paramters guide
'Granularity can be:
' Code  Max Period
' 1m    7 days
' 2m    60 days
' 5m    60 days
' 15m   60 days
' 30m   60 days
' 60m   730 days
' 90m   60 days
' 1h    730 days
' 1d    All
' 5d    All
' 1wk   All
' 1mo   All
' 3mo   All

'Events can be (separated by �7
'split
'dividends


'Variable declaration
Dim TickersRequests As clsHistoryRequest 'Buld to store common request arguments
Dim FXRequests As clsHistoryRequest 'to temporary store FX rate for base conversion
Dim Bmk As clsPriceHistory 'to temporary store Benchmark for relative quotes calculation
Dim Ticker As clsPriceHistory 'Built to stroe all price request in an object and then aggregate them
Dim FXRate As clsPriceHistory
Dim FxQuote As clsQuote
Dim Quote As clsQuote 'Class to add undefined number with Quotes variables
Dim RequestHeader As New Collection 'To be adappted to the relevant site
Dim URI As String 'Active URL Request
Dim response ' HTTP request respnse collector
Dim S, i, j, r, c, Rows, Columns As Integer 'Counters
Dim RefCur, TkCur As Object
Dim Tk, Fx, Timestamp, qt, TkQuote, BmkQt, Fld, DT, Tm As Variant    'Iterators
Dim Multiplier, Divisor, AdjFactor, RVcloseAdjFactor, RVadjCloseAdjFactor As Long
Dim QuoteFields As New Collection
Dim DataFields As Dictionary
Dim DataField As clsDataField
Dim MissingFields As String
Dim SupportedFields As String
Dim Intraday As Boolean
Dim Currencies As Collection
Dim TruncTimestamp As Long
Dim AggregateResponseTimeKeys As New Collection ' Added to get all different dates (without doubles) for aggregate response
Dim AggregateResponse() ' Aggragates all quotes on a cummom date range as an array to paste in worksheet
Dim Test As Variant ' Variable for debugging
Dim TempValue As Variant
Dim ResponseCurncy As String
Dim UDF As Boolean
Dim t As Integer

'Ambient
'Application.Volatile


'Determine if the functionis called as UDF or from a procedure
If TypeName(Application.Caller) = "Range" Then UDF = True Else UDF = False
'<> String Then UDF = True else UDF = False
If UDF = True And (OutputFirstCell <> "" Or RangeName <> "") Then
    MsgBox "Output Cell selection and Rangename are only supported if the function is called from a subrutine, nota as a UDF." & Chr(13) & _
           "Please leave those fileds empty if you are callng the function form a worksheet."
End If

'STEP: Initialize missing parameters
If (IsMissing(EndDate) Or IsEmpty(EndDate)) Then EndDate = Now()
If VarType(Fields) = vbString Then Fields = Array(Fields)
If VarType(Tickers) = vbString Then Tickers = Array(Tickers)
If VarType(OutputFirstCell) = vbString And OutputFirstCell <> "" Then
    Set OutputFirstCell = Range(OutputFirstCell)
End If



'Step: Check if Currency Transformation or multiple securities requested for intraday data and stores intraday Variable for Timestamp transformation
If Right(Granularity, 1) = "m" Then Intraday = True Else Intraday = False

If Intraday And RefCurrency <> "" Then
    MsgBox "Currency Transformation is not supported for intraday data." & Chr(13) & _
            "Please review formula."
    GoTo CleanUp
End If

''STEP: Read requested fields to collection
'For i = LBound(Fields) To UBound(Fields)
'    QuoteFields.Add LCase(Fields(i)) ' ensures is lowercase
'Next i

For Each Fld In Fields
    QuoteFields.Add LCase(Fld) ' ensures is lowercase
Next Fld


'Step: Initiate DataFields end DataField object
Set DataFields = New Dictionary
DataFields.CompareMode = vbTextCompare

Set DataField = New clsDataField
DataField.Label = "Symbol"
DataField.Scope = "Ticker"
DataFields.Add "Symbol", DataField

Set DataField = New clsDataField
DataField.Label = "Currency"
DataField.Scope = "Ticker"
DataFields.Add "curncy", DataField

Set DataField = New clsDataField
DataField.Label = "Exch.Name"
DataField.Scope = "Ticker"
DataFields.Add "exchange_name", DataField

Set DataField = New clsDataField
DataField.Label = "Type"
DataField.Scope = "Ticker"
DataFields.Add "instrument_type", DataField

Set DataField = New clsDataField
DataField.Label = "GMO Offset(sec)"
DataField.Scope = "Ticker"
DataFields.Add "gmt_offset_seconds", DataField

Set DataField = New clsDataField
DataField.Label = "Timezone"
DataField.Scope = "Ticker"
DataFields.Add "timezone", DataField

Set DataField = New clsDataField
DataField.Label = "Exch. Timezone Name"
DataField.Scope = "Ticker"
DataFields.Add "exchange_timezone_name", DataField

Set DataField = New clsDataField
DataField.Label = "Last"
DataField.Scope = "Ticker"
DataFields.Add "last_price", DataField

Set DataField = New clsDataField
DataField.Label = "Time Last"
DataField.Scope = "Ticker"
DataFields.Add "last_time", DataField

Set DataField = New clsDataField
DataField.Label = "Close Before Period"
DataField.Scope = "Ticker"
DataFields.Add "close_before_start_date", DataField

Set DataField = New clsDataField
DataField.Label = "GMT Timestamp"
DataField.Scope = "Quote"
DataFields.Add "gmt_timestamp", DataField

Set DataField = New clsDataField
DataField.Label = "Market Time"
DataField.Scope = "Quote"
DataFields.Add "market_datetime", DataField

Set DataField = New clsDataField
DataField.Label = "Local Time"
DataField.Scope = "Quote"
DataFields.Add "local_datetime", DataField

Set DataField = New clsDataField
DataField.Label = "Open"
DataField.Scope = "Quote"
DataFields.Add "open", DataField

Set DataField = New clsDataField
DataField.Label = "High"
DataField.Scope = "Quote"
DataFields.Add "high", DataField

Set DataField = New clsDataField
DataField.Label = "Low"
DataField.Scope = "Quote"
DataFields.Add "low", DataField

Set DataField = New clsDataField
DataField.Label = "Close"
DataField.Scope = "Quote"
DataFields.Add "close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj. Close"
DataField.Scope = "Quote"
DataFields.Add "adjclose", DataField

Set DataField = New clsDataField
DataField.Label = "Volume"
DataField.Scope = "Quote"
DataFields.Add "volume", DataField

Set DataField = New clsDataField
DataField.Label = "Prev. Close"
DataField.Scope = "Quote"
DataFields.Add "previous_close", DataField

Set DataField = New clsDataField
DataField.Label = "Close Start Date"
DataField.Scope = "Ticker"
DataFields.Add "first_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close Start Date"
DataField.Scope = "Quote"
DataFields.Add "first_adjclose", DataField

Set DataField = New clsDataField
DataField.Label = "% Close/Prev. Close"
DataField.Scope = "Quote"
DataFields.Add "pct_close_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "% Open/Pr.Close"
DataField.Scope = "Quote"
DataFields.Add "pct_open_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "% High/Low"
DataField.Scope = "Quote"
DataFields.Add "pct_high_low", DataField

Set DataField = New clsDataField
DataField.Label = "% Close/Open"
DataField.Scope = "Quote"
DataFields.Add "pct_close_open", DataField

Set DataField = New clsDataField
DataField.Label = "Close Indexed"
DataField.Scope = "Quote"
DataFields.Add "idx_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close Indexed"
DataField.Scope = "Quote"
DataFields.Add "idx_adjclose", DataField

'Reference Currency Fields
Set DataField = New clsDataField
DataField.Label = "Ref.Curncy"
DataField.Scope = "Ticker"
DataFields.Add "refcy", DataField

Set DataField = New clsDataField
DataField.Label = "Last" & " " & RefCurrency
DataField.Scope = "Ticker"
DataFields.Add "refcy_last_price", DataField

Set DataField = New clsDataField
DataField.Label = "First Close" & " " & RefCurrency
DataField.Scope = "Ticker"
DataFields.Add "refcy_first_close", DataField

Set DataField = New clsDataField
DataField.Label = "First Adj.Close" & " " & RefCurrency
DataField.Scope = "Ticker"
DataFields.Add "refcy_first_adjclose", DataField

Set DataField = New clsDataField
DataField.Label = "Close Before Period" & " " & RefCurrency
DataField.Scope = "Ticker"
DataFields.Add "refcy_close_before_start_date", DataField

Set DataField = New clsDataField
DataField.Label = "Open" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_open", DataField

Set DataField = New clsDataField
DataField.Label = "High" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_high", DataField

Set DataField = New clsDataField
DataField.Label = "Low" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_low", DataField

Set DataField = New clsDataField
DataField.Label = "Close" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_adj_open", DataField

Set DataField = New clsDataField
DataField.Label = "Prev.Close" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_previous_close", DataField

Set DataField = New clsDataField
DataField.Label = "%Close-Prv.Close" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_pct_close_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "%Open-Prv.Close" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_pct_open_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "%High-Low" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_pct_high_low", DataField

Set DataField = New clsDataField
DataField.Label = "%Close-Open" & " " & RefCurrency
DataField.Scope = "Quote"
DataFields.Add "refcy_pct_close_open", DataField

Set DataField = New clsDataField
DataField.Label = "Close" & " " & RefCurrency & " - Indexed"
DataField.Scope = "Quote"
DataFields.Add "idx_refcy_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close" & " " & RefCurrency & " - Indexed"
DataField.Scope = "Quote"
DataFields.Add "idx_refcy_adjclose", DataField

'Benchmark Relative Fields
Set DataField = New clsDataField
DataField.Label = "Last(Relative)"
DataField.Scope = "Ticker"
DataFields.Add "rv_last_price", DataField

Set DataField = New clsDataField
DataField.Label = "Close Before Period(Relative)"
DataField.Scope = "Ticker"
DataFields.Add "rv_close_before_start_date", DataField

Set DataField = New clsDataField
DataField.Label = "Open(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_open", DataField

Set DataField = New clsDataField
DataField.Label = "High(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_high", DataField

Set DataField = New clsDataField
DataField.Label = "Low(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_low", DataField

Set DataField = New clsDataField
DataField.Label = "Close(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_adj_open", DataField

Set DataField = New clsDataField
DataField.Label = "Prev.Close(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_previous_close", DataField

Set DataField = New clsDataField
DataField.Label = "%Close-Prv.Close(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_pct_close_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "%Open-Prv.Close(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_pct_open_previousclose", DataField

Set DataField = New clsDataField
DataField.Label = "%High-Low(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_pct_high_low", DataField

Set DataField = New clsDataField
DataField.Label = "%Close-Open(Relative)"
DataField.Scope = "Quote"
DataFields.Add "rv_pct_close_open", DataField

Set DataField = New clsDataField
DataField.Label = "Close Indexed Relative"
DataField.Scope = "Quote"
DataFields.Add "idx_rv_close", DataField

Set DataField = New clsDataField
DataField.Label = "Adj.Close Indexed Relative"
DataField.Scope = "Quote"
DataFields.Add "idx_rv_adjclose", DataField


'Debug.Print Join(DataFields.Keys, Chr(13))

'Step: Check if requested field is in supported fields _
        Return error message with list off supported fields if not supported and exit function

MissingFields = ""
For Each Fld In QuoteFields
    If Not DataFields.Exists(Fld) Then
        If MissingFields = "" Then
            MissingFields = Fld
        Else
            MissingFields = MissingFields & Chr(10) & Fld
        End If
    End If
Next Fld
If MissingFields <> "" Then
    MsgBox "Following requested field(s) are not supported:" & Chr(13) & Chr(13) & _
            MissingFields & Chr(13) & Chr(13) & _
            "Supported Fields are: " & Chr(13) & Chr(13) & Join(DataFields.Keys, Chr(10)) & Chr(13) & Chr(13) & _
            "Please correct formuala accordingly."
    GoTo CleanUp
    
End If

'STEP: Transform input parmaters to yahoo finance URLencoded format
If Intraday Then 'intrady chart; starttime is assumed to be in GMT and the response will be in GMT
    StartDate = Date2Timestamp(StartDate, 7200)
    EndDate = Date2Timestamp(EndDate, 7200)
Else 'daily, weekly or monthly chart, starttime ist irrelevant
    Debug.Print "Startdate: " & StartDate
    Debug.Print "Enddate: " & EndDate
    StartDate = Date2Timestamp(StartDate)
    EndDate = Date2Timestamp(EndDate) '+ 86400
    Debug.Print "Timestamp Start: " & StartDate
    Debug.Print "Timestamp End: " & EndDate
End If

'Step: Set Common Request Header Collection
RequestHeader.Add Array("Content-Type", "text/json")
RequestHeader.Add Array("Accept", "application/json, text/javascript, */*; q=0.01")
RequestHeader.Add Array("Accept-Language", "it-CH,en;q=0.5")
RequestHeader.Add Array("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
RequestHeader.Add Array("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:75.0) Gecko/20100101 Firefox/75.0")

'Step: Create new Tickers Requests Collector
Set TickersRequests = New clsHistoryRequest

'Step: Define URL specific Request parameters Get Price History of every Ticker
For Each Tk In Tickers
    'Prepare and send the request to request function
    'If needed, you can add to the request URL underneath "&events=" & "split%7Cdividends" &
    'if Filename is not empty, the file will be saved and opened in the browser for debugging
    'Example:
    'https://query1.finance.yahoo.com/v8/finance/chart/NOVN.SW?symbol=NOVN.SW&period1=1545087600&period2=1560463200&interval=1d&includePrePost=true&events=div%7Csplit%7Cearn&lang=en-US&region=US&crumb=bhA6qXHt.gb&corsDomain=finance.yahoo.com
    'N.B. Return time stamps, in GMT
    'Currencies: D-1 23:00
    'Asia: D-0 00:00
    'Europe: D-0 Market Open Time - Market Close Time
    'US: D-0 Market Opentime - Market Opentime
    'Futures: D-0 04:00 - 04:00
    'Offset by one day if currencies, because Yahoo returns the previsous day close at 2300
    If (Not Intraday) And Right(Tk, 1) = "X" Then
            URI = "https://query1.finance.yahoo.com/v8/finance/chart/" & CStr(Tk) & "?" & "symbol=" & CStr(Tk) & _
            "&period1=" & StartDate & _
            "&period2=" & EndDate & _
            "&interval=" & Granularity & _
            "&corsDomain=finance.yahoo.com"
    'https://query1.finance.yahoo.com/v8/finance/chart/AUDCHF=X?&interval=1d&period1=1591142400&period2=1591606547&events=div%7Csplit&corsDomain=finance.yahoo.com
    Else
            URI = "https://query1.finance.yahoo.com/v8/finance/chart/" & CStr(Tk) & "?" & "symbol=" & CStr(Tk) & _
            "&period1=" & StartDate & _
            "&period2=" & EndDate & _
            "&interval=" & Granularity & _
            "&includePrePost=true" & _
            "&corsDomain=finance.yahoo.com"
    End If

    'To store Response Json substitute "" with Tk in Filename
    Set response = HTTPRequest( _
        Agent:="XMLHttp", _
        URL:=URI, _
        Format:="json", _
        RequestHeader:=RequestHeader, _
        RequestBody:="", _
        User:="", _
        password:="", _
        filename:="", _
        filePath:="")
    Debug.Print URI
    'Create Object to store response data for the current ticker and store it into Ticker collection
    Set Ticker = New clsPriceHistory
    With response.item("chart").item("result").item(1)
        'Reads the Ticker history meta data
        Ticker.Symbol = .item("meta").item("symbol")
        Ticker.Fields.Add .item("meta").item("currency"), "curncy"
        Ticker.Fields.Add .item("meta").item("exchangeName"), "exchange_name"
        Ticker.Fields.Add .item("meta").item("instrumentType"), "instrument_type"
        Ticker.Fields.Add .item("meta").item("gmtoffset"), "gmt_offset_seconds"
        Ticker.Fields.Add .item("meta").item("timezone"), "timezone"
        Ticker.Fields.Add .item("meta").item("exchangeTimezoneName"), "exchange_timezone_name"
        Ticker.Fields.Add Timestamp2Date(.item("meta").item("firstTradeDate"), .item("meta").item("gmtoffset")), "first_trade_date"
        Ticker.Fields.Add CdblOrNull(.item("meta").item("regularMarketPrice")), "last_price"
        Ticker.Fields.Add Timestamp2Date(.item("meta").item("regularMarketTime"), .item("meta").item("gmtoffset")), "last_time"
        Ticker.Fields.Add CdblOrNull(.item("meta").item("chartPreviousClose")), "close_before_start_date"
        i = 0
        
        'Read quotes data and store it in the Quotes Collection (Key = Timestamp truncated to the day)
        For Each Timestamp In .item("timestamp")
            i = i + 1
            'Create new element in Quotes collection to store the fields values
            Set TkQuote = New clsQuote
            If Intraday Then
                TkQuote.TimeKey = Timestamp
            ElseIf Right(Tk, 1) = "X" Then
                TkQuote.TimeKey = CStr(Fix((Timestamp + 3600) / 86400) * 86400)
            Else
                TkQuote.TimeKey = CStr(Fix(Timestamp / 86400) * 86400) 'Truncate Timestamp to Day and stores it a search key;
            End If
            TkQuote.TimeKeyDateTime = Timestamp2Date(TkQuote.TimeKey) 'added as parameter for debugging
            TkQuote.IndexNo = i      'Useful to retrieve relative positions from start / end of series for calculations; MA for instance
            On Error Resume Next 'in case the key already exists in the collection
               AggregateResponseTimeKeys.Add TkQuote.TimeKey, CStr(TkQuote.TimeKey)  'Sets the day date as search key
            On Error GoTo 0
            'Store additonal human readable quote Time Information in fields
            TkQuote.Fields.Add Timestamp, "gmt_timestamp"
            TkQuote.Fields.Add Timestamp2Date(.item("timestamp").item(i), Ticker.Fields("gmt_offset_seconds")), "market_datetime"
            TkQuote.Fields.Add Timestamp2Date(.item("timestamp").item(i), 7200), "local_datetime"
            With .item("indicators")
                TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("open").item(i)), "open"
                TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("high").item(i)), "high"
                TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("low").item(i)), "low"
                TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("close").item(i)), "close"
                On Error Resume Next ' in case there is no AdjClose (as for FX)
                    TkQuote.Fields.Add CdblOrNull(.item("adjclose").item(1).item("adjclose").item(i)), "adjclose"
                    On Error Resume Next
                    'Stores first value for percentage calculation
                    If i = 1 Then Ticker.Fields.Add CdblOrNull(TkQuote.Fields("adjclose")), "first_adjclose"
                On Error GoTo 0
                'Stores first value for percentage calculation
                If i = 1 Then
                    Ticker.Fields.Add CdblOrNull(TkQuote.Fields("close")), "first_close"
                    TkQuote.Fields.Add CdblOrNull(Ticker.Fields("close_before_start_date")), "previous_close"
                Else
                    TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("close").item(i - 1)), "previous_close"
                End If
                TkQuote.Fields.Add CdblOrNull(.item("quote").item(1).item("volume").item(i)), "volume"
                
                'Add some commonly requested fields '??
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("close") / TkQuote.Fields("previous_close") - 1), "pct_close_previousclose"
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("open") / TkQuote.Fields("previous_close") - 1), "pct_open_previousclose"
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("high") / TkQuote.Fields("low") - 1), "pct_high_low"
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("close") / TkQuote.Fields("open") - 1), "pct_close_open"
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("close") / Ticker.Fields("first_close")), "idx_close"
                TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("adjclose") / Ticker.Fields("first_adjclose")), "idx_adjclose"
            End With
            'Todo: Calculate and stores additional indicators
            
            'Store Quote object in Ticker object with Timestamp as key
            Ticker.Quotes.Add item:=TkQuote, key:=TkQuote.TimeKey
            'debug.print "BmkTkQuote TimeKey: " & Timestamp2Date(TkQuote.TimeKey) & "(" & TkQuote.TimeKey & ")" _
                        & " - " & "Timestamp: " & Timestamp2Date(Timestamp) & "(" & Timestamp & ")"
            Set TkQuote = Nothing
        Next Timestamp
    ' Add last quote to FX quotes, because Yahoo doesn't return it
'    If (Not Intraday) And Right(Ticker.Symbol, 1) = "X" And (Fix(EndDate / 86400) = Fix(Date2Timestamp(Now) / 86400)) Then
'    ' The current FX rate is requested
'        Set TkQuote = New clsQuote
'        Timestamp = .Item("meta").Item("currentTradingPeriod").Item("regular").Item("end") ' Time of last trading on end date
'        TkQuote.TimeKey = CStr(Fix(Timestamp / 86400) * 86400) 'Truncate Timestamp to Day and stores it a search key
'        TkQuote.TimeKeyDateTime = Timestamp2Date(TkQuote.TimeKey) 'added as parameter for debugging
'        TkQuote.IndexNo = i + 1    'Useful to retrieve relative positions from start / end of series for calculations; MA for instance
'        TkQuote.Fields.Add Timestamp, "gmt_timestamp"
'        TkQuote.Fields.Add Timestamp2Date(Timestamp, Ticker.Fields("gmt_offset_seconds")), "market_datetime"
'        TkQuote.Fields.Add Timestamp2Date(Timestamp, 7200), "local_datetime"
'        TkQuote.Fields.Add CdblOrNull(Ticker.Fields.Item("last_price")), "close"
'        TkQuote.Fields.Add CdblOrNull(TkQuote.Fields("close") / Ticker.Fields("first_close")), "idx_close"
'        Ticker.Quotes.Add Item:=TkQuote, Key:=TkQuote.TimeKey
'        Set TkQuote = Nothing
'    End If
    TickersRequests.Tickers.Add Ticker, Ticker.Symbol
    Set Ticker = Nothing
    End With
Next Tk


'Step: Get Reference Currency and Tickers currencies if required

If RefCurrency <> "" Then
    Set Currencies = New Collection
    'Add Benchmark currency to list of FXrates to download
    If RefCurrency <> "USD" Then
        Currencies.Add RefCurrency, RefCurrency
    End If
    'Determine other needed currencies for conversion
    For Each Tk In TickersRequests.Tickers
        On Error Resume Next
            If Tk.Fields("curncy") <> "USD" Then
                Currencies.Add item:=Tk.Fields("curncy"), key:=Tk.Fields("curncy")
            End If
        On Error GoTo 0
    Next Tk
    
    'Get releavant FX closing prices vs. USD
    Set FXRequests = New clsHistoryRequest
    For Each Fx In Currencies
        If (Not Intraday) Then
                URI = "https://query1.finance.yahoo.com/v8/finance/chart/" & CStr(Fx) & "=X" & "?" & "symbol=" & CStr(Fx) & "=X" & _
                "&period1=" & StartDate & _
                "&period2=" & EndDate & _
                "&interval=" & Granularity & _
                "&corsDomain=finance.yahoo.com"
        Else
                URI = "https://query1.finance.yahoo.com/v8/finance/chart/" & CStr(Fx) & "=X" & "?" & "symbol=" & CStr(Fx) & "=X" & _
                "&period1=" & StartDate & _
                "&period2=" & EndDate & _
                "&interval=" & Granularity & _
                "&includePrePost=true" & _
                "&corsDomain=finance.yahoo.com"
        End If
        Debug.Print Timestamp2Date(StartDate)
        Debug.Print Timestamp2Date(EndDate)
        Debug.Print URI
        'To store Response Json substitute "" with CStr(Fx) & "=X" in Filename
        Set response = HTTPRequest( _
                Agent:="XMLHttp", _
                URL:=URI, _
                Format:="json", _
                RequestHeader:=RequestHeader, _
                RequestBody:="", _
                User:="", _
                password:="", _
                filename:="", _
                filePath:="")

        Set FXRate = New clsPriceHistory
        With response.item("chart").item("result").item(1)
            'Reads the price hisotry meta data
            FXRate.Symbol = .item("meta").item("symbol")
            FXRate.Fields.Add .item("meta").item("currency"), "curncy"
            FXRate.Fields.Add .item("meta").item("exchangeName"), "exchange_name"
            FXRate.Fields.Add .item("meta").item("instrumentType"), "instrument_type"
            FXRate.Fields.Add .item("meta").item("gmtoffset"), "gmt_offset_seconds"
            FXRate.Fields.Add .item("meta").item("timezone"), "timezone"
            FXRate.Fields.Add .item("meta").item("exchangeTimezoneName"), "exchange_timezone_name"
            FXRate.Fields.Add Timestamp2Date(.item("meta").item("firstTradeDate"), .item("meta").item("gmtoffset")), "first_trade_date"
            FXRate.Fields.Add CdblOrNull(.item("meta").item("regularMarketPrice")), "last_price"
            FXRate.Fields.Add Timestamp2Date(.item("meta").item("regularMarketTime"), .item("meta").item("gmtoffset")), "last_time"
            FXRate.Fields.Add CdblOrNull(.item("meta").item("chartPreviousClose")), "close_before_start_date"

            'Read quotes into FXRequests tickers collection and store it in the Quotes Collection (Key = Timestamp truncated to the day)
            i = 0
            For Each Timestamp In .item("timestamp")
                i = i + 1
                Set FxQuote = New clsQuote
                If Intraday Then
                    FxQuote.TimeKey = Timestamp
                Else
                   FxQuote.TimeKey = CStr(Fix((Timestamp + 3600) / 86400) * 86400) 'Truncate Timestamp to Day and stores it a search key;
                End If
                FxQuote.TimeKeyDateTime = Timestamp2Date(FxQuote.TimeKey)
                FxQuote.IndexNo = i
                FxQuote.Fields.Add Timestamp, "gmt_timestamp"
                If IsNull(.item("indicators").item("quote").item(1).item("close").item(i)) Then
                    j = 1
                    Do While IsNull(.item("indicators").item("quote").item(1).item("close").item(i - j))
                        j = j + 1
                    Loop
                    FxQuote.Fields.Add CdblOrNull(.item("indicators").item("quote").item(1).item("close").item(i - j)), "close"
                Else
                    FxQuote.Fields.Add CdblOrNull(.item("indicators").item("quote").item(1).item("close").item(i)), "close"
                End If
                If i = 1 Then
                    FxQuote.Fields.Add CdblOrNull(FXRate.Fields("close_before_start_date")), "previous_close"
                    FXRate.Fields.Add CdblOrNull(FxQuote.Fields("close")), "first_close"
                Else
                    FxQuote.Fields.Add CdblOrNull(FXRate.Quotes(i - 1).Fields("close")), "previous_close"
                End If
                FXRate.Quotes.Add item:=FxQuote, key:=FxQuote.TimeKey
                'debug.print "FxQuote TimeKey: " & Timestamp2Date(FxQuote.TimeKey) & "(" & FxQuote.TimeKey & ")" _
                            & " - " & "Timestamp: " & Timestamp2Date(Timestamp) & "(" & Timestamp & ")"
                Set FxQuote = Nothing
            Next Timestamp
'            ' Add last quote to FX quotes, because Yahoo doesn't return it
'            ' The current FX rate is requested
'            Set FxQuote = New clsQuote
'            Timestamp = .Item("meta").Item("currentTradingPeriod").Item("regular").Item("end") ' Time of last trading on end date
'            FxQuote.TimeKey = CStr(Fix(Timestamp / 86400) * 86400) 'Truncate Timestamp to Day and stores it a search key
'            FxQuote.TimeKeyDateTime = Timestamp2Date(FxQuote.TimeKey) 'added as parameter for debugging
'            FxQuote.IndexNo = i + 1    'Useful to retrieve relative positions from start / end of series for calculations; MA for instance
'            FxQuote.Fields.Add Timestamp, "gmt_timestamp"
'            FxQuote.Fields.Add Timestamp2Date(Timestamp, FXRate.Fields("gmt_offset_seconds")), "market_datetime"
'            FxQuote.Fields.Add Timestamp2Date(Timestamp, 7200), "local_datetime"
'            FxQuote.Fields.Add CdblOrNull(FXRate.Fields.Item("last_price")), "close"
'            FxQuote.Fields.Add CdblOrNull(FxQuote.Fields("close") / FXRate.Fields("first_close")), "idx_close"
'            FXRate.Quotes.Add Item:=FxQuote, Key:=FxQuote.TimeKey
'            Set FxQuote = Nothing
            FXRequests.Tickers.Add FXRate, FXRate.Fields("curncy")
            Set FXRate = Nothing
        End With
    Next Fx
    
    'Add all relevant RefCurncy to Ticker fields
    For Each Tk In TickersRequests.Tickers
        If RefCurrency <> "USD" Then
            Set RefCur = FXRequests.Tickers(RefCurrency)
            Tk.Fields.Add RefCur.Fields("curncy"), "refcy"
            If Tk.Fields("curncy") <> "USD" Then
                Set TkCur = FXRequests.Tickers(Tk.Fields("curncy"))
                Tk.Fields.Add CdblOrNull(Tk.Fields("last_price") / TkCur.Fields("last_price") * RefCur.Fields("last_price")), "refcy_last_price"
                Tk.Fields.Add CdblOrNull(Tk.Fields("close_before_start_date") / TkCur.Fields("close_before_start_date") * RefCur.Fields("close_before_start_date")), "refcy_close_before_start_date"
            Else
                Tk.Fields.Add CdblOrNull(Tk.Fields("last_price") * RefCur.Fields("last_price")), "refcy_last_price"
                Tk.Fields.Add CdblOrNull(Tk.Fields("close_before_start_date") * RefCur.Fields("close_before_start_date")), "refcy_close_before_start_date"
            End If
        Else
            Tk.Fields.Add "USD", "refcy"
            If Tk.Fields("curncy") <> "USD" Then
                Tk.Fields.Add CdblOrNull(Tk.Fields("last_price")), "refcy_last_price"
                Tk.Fields.Add CdblOrNull(Tk.Fields("close_before_start_date")), "refcy_close_before_start_date"
            Else
            Tk.Fields.Add CdblOrNull(Tk.Fields("last_price") / TkCur.Fields("last_price")), "refcy_last_price"
            Tk.Fields.Add CdblOrNull(Tk.Fields("close_before_start_date") / TkCur.Fields("close_before_start_date")), "refcy_close_before_start_date"
            End If
        End If
        i = 0
        For Each qt In Tk.Quotes
            i = i + 1
            If i = 1 Then
                If RefCurrency <> "USD" Then
                    If Tk.Fields("curncy") <> "USD" Then
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_close") / TkCur.Fields("first_close") * RefCur.Fields("first_close")), key:="refcy_first_close"
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_adjclose") / TkCur.Fields("first_close") * RefCur.Fields("first_close")), key:="refcy_first_adjclose"
                    Else
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_close") * RefCur.Fields("first_close")), key:="refcy_first_close"
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_adjclose") * RefCur.Fields("first_close")), key:="refcy_first_adjclose"
                    End If
                Else
                    If Tk.Fields("curncy") <> "USD" Then
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_close") / TkCur.Fields("first_close")), key:="refcy_first_close"
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_adjclose") / TkCur.Fields("first_close")), key:="refcy_first_adjclose"
                    Else
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_close")), key:="refcy_first_close"
                        Tk.Fields.Add item:=CdblOrNull(Tk.Fields("first_adjclose")), key:="refcy_first_adjclose"
                    End If
                End If
            End If
            If RefCurrency <> "USD" Then
                Multiplier = CdblOrNull(RefCur.Quotes(qt.TimeKey).Fields("close"))
            Else
                Multiplier = 1
            End If
            If Tk.Fields("curncy") <> "USD" Then
                Divisor = CdblOrNull(TkCur.Quotes(CStr(qt.TimeKey)).Fields("close"))
            Else
                Divisor = 1
            End If
            AdjFactor = CdblOrNull(Multiplier / Divisor)
            
            qt.Fields.Add item:=CdblOrNull(qt.Fields("open") * AdjFactor), key:="refcy_open"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("high") * AdjFactor), key:="refcy_high"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("low") * AdjFactor), key:="refcy_low"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("close") * AdjFactor), key:="refcy_close"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("adjclose") * AdjFactor), key:="refcy_adjclose"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("previous_close") * AdjFactor), key:="refcy_previous_close"
            'Add some commonly requested fields
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_close") / qt.Fields("refcy_previous_close") - 1)), "refcy_pct_close_previousclose"
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_open") / qt.Fields("refcy_previous_close") - 1)), "refcy_pct_open_previousclose"
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_high") / qt.Fields("refcy_low") - 1)), "refcy_pct_high_low"
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_close") / qt.Fields("refcy_open") - 1)), "refcy_pct_close_open"
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_close") / Tk.Fields("refcy_first_close"))), "idx_refcy_close"
            qt.Fields.Add CdblOrNull((qt.Fields("refcy_adjclose") / Tk.Fields("refcy_first_adjclose"))), "idx_refcy_adjclose"
            On Error GoTo 0
        Next qt
    Next Tk
End If

'Step: Get Benchmark if required
If Benchmark <> "" Then
    'Get Price History of benchmark
    'To store Response Json substitute "" with Benchmark in Filename
    URI = "https://query1.finance.yahoo.com/v8/finance/chart/" & CStr(Benchmark) & "?" & "symbol=" & CStr(Benchmark) & _
            "&period1=" & StartDate & _
            "&period2=" & EndDate & _
            "&interval=" & Granularity & _
            "&includePrePost=true" & _
            "&corsDomain=finance.yahoo.com"
    Set response = HTTPRequest( _
        Agent:="XMLHttp", _
        URL:=URI, _
        Format:="json", _
        RequestHeader:=RequestHeader, _
        RequestBody:="", _
        User:="", _
        password:="", _
        filename:="", _
        filePath:="")
    
    Debug.Print URI
    
    'Create Object to store response data for the ticker
    Set Bmk = New clsPriceHistory
    
    With response.item("chart").item("result").item(1)
        'Reads the Bmk history meta data
        Bmk.Symbol = .item("meta").item("symbol")
        Bmk.Fields.Add .item("meta").item("currency"), "curncy"
        Bmk.Fields.Add .item("meta").item("exchangeName"), "exchange_name"
        Bmk.Fields.Add .item("meta").item("instrumentType"), "instrument_type"
        Bmk.Fields.Add .item("meta").item("gmtoffset"), "gmt_offset_seconds"
        Bmk.Fields.Add .item("meta").item("timezone"), "timezone"
        Bmk.Fields.Add .item("meta").item("exchangeTimezoneName"), "exchange_timezone_name"
        Bmk.Fields.Add Timestamp2Date(.item("meta").item("firstTradeDate"), .item("meta").item("gmtoffset")), "first_trade_date"
        Bmk.Fields.Add CdblOrNull(.item("meta").item("regularMarketPrice")), "last_price"
        Bmk.Fields.Add Timestamp2Date(.item("meta").item("regularMarketTime"), .item("meta").item("gmtoffset")), "last_time"
        Bmk.Fields.Add CdblOrNull(.item("meta").item("chartPreviousClose")), "close_before_start_date"
        i = 0
        
        'Read Bmks quotes data and store it in the Quotes Collection (Key = Timestamp truncated to the day)
        For Each Timestamp In .item("timestamp")
            i = i + 1
            'Create new element in Quotes collection to store the fields values
            Set BmkQt = New clsQuote
            If Intraday Then
                BmkQt.TimeKey = Timestamp
            Else
                BmkQt.TimeKey = CStr(Fix(Timestamp / 86400) * 86400) 'Truncate Timestamp to Day and sores it a search key;
            End If
            BmkQt.TimeKeyDateTime = Timestamp2Date(BmkQt.TimeKey) 'added as parameter for debugging
            BmkQt.IndexNo = i      'Useful to retrieve relative positions from start / end of series for calculations; MA for instance
            On Error Resume Next 'in case the key already exists in the collection
                AggregateResponseTimeKeys.Add BmkQt.TimeKey, CStr(BmkQt.TimeKey)  'Sets the date as search key
            On Error GoTo 0
            'Store additonal human readable quote Time Information in fields
            BmkQt.Fields.Add Timestamp, "gmt_timestamp"
            BmkQt.Fields.Add Timestamp2Date(.item("timestamp").item(i), Bmk.Fields("gmt_offset_seconds")), "market_datetime"
            BmkQt.Fields.Add Timestamp2Date(.item("timestamp").item(i), 7200), "local_datetime"
            With .item("indicators")
                BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("open").item(i)), "open"
                BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("high").item(i)), "high"
                BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("low").item(i)), "low"
                BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("close").item(i)), "close"
                On Error Resume Next ' in case there is no AdjClose (as for FX)
                    BmkQt.Fields.Add CdblOrNull(.item("adjclose").item(1).item("adjclose").item(i)), "adjclose"
                    On Error Resume Next
                    'Stores first value for percentage calculation
                    If i = 1 Then Bmk.Fields.Add CdblOrNull(BmkQt.Fields("adjclose")), "first_adjclose"
                On Error GoTo 0
                'Stores first value for percentage calculation
                If i = 1 Then
                    Bmk.Fields.Add CdblOrNull(BmkQt.Fields("close")), "first_close"
                    BmkQt.Fields.Add CdblOrNull(Bmk.Fields("close_before_start_date")), "previous_close"
                Else
                    BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("close").item(i - 1)), "previous_close"
                End If
                BmkQt.Fields.Add CdblOrNull(.item("quote").item(1).item("volume").item(i)), "volume"
                
                'Add some commonly requested fields
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("close") / BmkQt.Fields("previous_close") - 1)), "pct_close_previousclose"
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("open") / BmkQt.Fields("previous_close") - 1)), "pct_open_previousclose"
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("high") / BmkQt.Fields("low") - 1)), "pct_high_low"
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("close") / BmkQt.Fields("open") - 1)), "pct_close_open"
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("close") / Bmk.Fields("first_close"))), "idx_close"
                BmkQt.Fields.Add CdblOrNull((BmkQt.Fields("adjclose") / Bmk.Fields("first_adjclose"))), "idx_adjclose"
            End With
            
            'Calculate and stores additional indicators
            'ToDo
            
            'Store Quote object in Bmk object with Timestamp as key
            Bmk.Quotes.Add item:=BmkQt, key:=BmkQt.TimeKey
            'debug.print "BmkQt TimeKey: " & Timestamp2Date(BmkQt.TimeKey) & "(" & BmkQt.TimeKey & ")" _
                        & " - " & "Timestamp: " & Timestamp2Date(Timestamp) & "(" & Timestamp & ")"
            Set BmkQt = Nothing
        Next Timestamp
    End With
    'Add all relevant Relative fields to Ticker fields
    For Each Tk In TickersRequests.Tickers
        RVcloseAdjFactor = CdblOrNull(Bmk.Fields("first_close") / Tk.Fields("first_close"))
        RVadjCloseAdjFactor = CdblOrNull(Bmk.Fields("first_adjclose") / Tk.Fields("first_adjclose"))
        Tk.Fields.Add CdblOrNull(Tk.Fields("last_price") / Bmk.Fields("last_price") * RVcloseAdjFactor), "rv_last_price"
        Tk.Fields.Add CdblOrNull(Tk.Fields("close_before_start_date") / Bmk.Fields("close_before_start_date") * RVcloseAdjFactor), "rv_close_before_start_date"
        i = 0
        For Each qt In Tk.Quotes
            i = i + 1
            qt.Fields.Add item:=CdblOrNull(qt.Fields("open") / Bmk.Quotes(qt.TimeKey).Fields("open") * RVcloseAdjFactor), key:="rv_open"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("high") / Bmk.Quotes(qt.TimeKey).Fields("high") * RVcloseAdjFactor), key:="rv_high"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("low") / Bmk.Quotes(qt.TimeKey).Fields("low") * RVcloseAdjFactor), key:="rv_low"
            qt.Fields.Add item:=CdblOrNull(qt.Fields("close") / Bmk.Quotes(qt.TimeKey).Fields("close") * RVcloseAdjFactor), key:="rv_close"
               
            On Error Resume Next ' in case there is no AdjClose (as for Bmk)
                qt.Fields.Add item:=CdblOrNull(qt.Fields("adjclose") / Bmk.Quotes(qt.TimeKey).Fields("adjclose") * RVadjCloseAdjFactor), key:="rv_adjclose"
            On Error GoTo 0
                'Stores first value for percentage calculation
                If i = 1 Then
                    qt.Fields.Add CdblOrNull(Tk.Fields("rv_close_before_start_date")), "rv_previous_close"
                Else
                    qt.Fields.Add CdblOrNull(Tk.Quotes.item(i - 1).Fields("rv_close")), "rv_previous_close"
                End If

            'Add some commonly requested fields
            qt.Fields.Add CdblOrNull((qt.Fields("rv_close") / qt.Fields("rv_previous_close") - 1)), "rv_pct_close_previousclose"
            qt.Fields.Add CdblOrNull((qt.Fields("rv_open") / qt.Fields("rv_previous_close") - 1)), "rv_pct_open_previousclose"
            qt.Fields.Add CdblOrNull((qt.Fields("rv_high") / qt.Fields("rv_low") - 1)), "rv_pct_high_low"
            qt.Fields.Add CdblOrNull((qt.Fields("rv_close") / qt.Fields("rv_open") - 1)), "rv_pct_close_open"
        Next qt
    Next Tk

End If

'Step: Build return Array "AggregateResponse" with Rows for Dates (+ Label if required) from AggregateResponseTimeKeys and Tickers x Requestfields column
'Sort the TimeKeys collection in ascending order
Set AggregateResponseTimeKeys = SortCollection(AggregateResponseTimeKeys, True)

Columns = TickersRequests.Tickers.Count * QuoteFields.Count + 1
If OutputFieldsLables = True Then
    Rows = AggregateResponseTimeKeys.Count + 1
    ReDim AggregateResponse(1 To Rows, 1 To Columns)
    AggregateResponse(1, 1) = "Date"
    r = 2
    c = 2
    For Each Tk In TickersRequests.Tickers
        For Each Fld In QuoteFields
            AggregateResponse(1, c) = Tk.Symbol & " (" & DataFields(Fld).Label & ")"
            c = c + 1
        Next Fld
    Next Tk
    c = 1
Else
    Rows = AggregateResponseTimeKeys.Count
    Columns = TickersRequests.Tickers.Count * QuoteFields.Count + 1
    ReDim AggregateResponse(1 To Rows, 1 To Columns)
    r = 1
    c = 1
End If

For Each Tm In AggregateResponseTimeKeys
    AggregateResponse(r, c) = Timestamp2Date(Tm, 7200)
    r = r + 1
Next Tm

c = 1
For Each Tk In TickersRequests.Tickers
    For Each Fld In QuoteFields
        If OutputFieldsLables = True Then r = 1 Else r = 0
        c = c + 1
        On Error GoTo NotFound
        
        For Each Tm In AggregateResponseTimeKeys
            r = r + 1
            If DataFields(Fld).Scope = "Ticker" Then
                AggregateResponse(r, c) = Null2NA(Tk.Fields(Fld))
            Else
                AggregateResponse(r, c) = Null2NA(Tk.Quotes(Tm).Fields(Fld))
            End If
        Next Tm
        GoTo Found:
NotFound:
        AggregateResponse(r, c) = CVErr(xlErrNA)
        Resume Next
Found:
        On Error GoTo 0
    Next Fld
Next Tk

If UDF Then
    YahooHistory = AggregateResponse
Else
    Array2Range AggregateResponse, OutputFirstCell, , RangeName
End If

CleanUp:
Set DataFields = Nothing

End Function

Private Sub Array2Range(ByVal objArray, FirstCell, Optional Transpose As Boolean = False, _
                        Optional RangeName As String = "", Optional FormatAsTable As Boolean = True)
Dim r, c As Integer
Dim Destination As Range

'Determine rows and columns of array
r = UBound(objArray, 1) - LBound(objArray, 1) + 1
c = UBound(objArray, 2) - LBound(objArray, 2) + 1
'Check or transform Firstcell to range
If Not IsObject(FirstCell) Then Set FirstCell = Range(FirstCell)
If Transpose Then
    Set Destination = FirstCell.Cells(1, 1).Resize(c, r)
    Destination.value = Application.Transpose(objArray)
Else
    Set Destination = FirstCell.Cells(1, 1).Resize(r, c)
    Destination.value = objArray
End If

If FormatAsTable Then
    On Error Resume Next
    ActiveSheet.ListObjects.Add(xlSrcRange, Destination, , xlYes).Name = "ArrayTable"
    If RangeName <> "" Then
        ActiveSheet.ListObjects("ArrayTable").Name = RangeName
    End If
    On Error GoTo 0
ElseIf RangeName <> "" Then
    ActiveWorkbook.Names.Add RangeName, Destination
End If
End Sub



Function ScaricaFile(ByVal myURL As String, ByVal NomeUtente As String, ByVal password As String, ByVal filePath As String, ByVal filename As String, ByVal Overwrite As Boolean)

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False, NomeUtente, password
    WinHttpReq.Send

    myURL = WinHttpReq.responseBody
    'MsgBox myURL
    If myURL = "" Then
        MsgBox "Purtroppo non � stato possibile scaricare il file " & filename & " o il file risulta vuoto.L'operazione di salvataggio del file � stata interrota.", vbCritical, "Errore nello scaricmanteo del file da Internet"
        Exit Function
    End If
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        If Overwrite Then
            oStream.SaveToFile filePath & filename, 2
        Else
            ' 1 = no overwrite, 2 = overwrite
            oStream.SaveToFile filePath & filename, 1
        End If
        oStream.Close
    End If

End Function

Function ScaricaFileEncoded(ByVal myURL As String, ByVal NomeUtente As String, ByVal password As String, ByVal filePath As String, ByVal filename As String, ByVal Overwrite As Boolean, ByVal Encoded As Boolean)

    Dim xmlhttp As New MSXML2.ServerXMLHTTP60
    xmlhttp.Open "GET", myURL, False, NomeUtente, password
    If Encoded = True Then
        xmlhttp.SetRequestHeader "Content-Type", "text/json"
        xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
        xmlhttp.SetRequestHeader "Authorization", "Basic " + Base64Encode(NomeUtente + ":" + password)
    End If
    xmlhttp.Send

    myURL = xmlhttp.responseBody
    'MsgBox myURL
    If myURL = "" Then
        MsgBox "Purtroppo non � stato possibile scaricare il file " & filename & " o il file risulta vuoto.L'oprazione di salvataggio del file � stata interrota.", vbCritical, "Errore nello scaricmanteo del file da Internet"
        Exit Function
    End If
    If xmlhttp.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write xmlhttp.responseBody
        If Overwrite Then
            oStream.SaveToFile filePath & filename, 2
        Else
            ' 1 = no overwrite, 2 = overwrite
            oStream.SaveToFile filePath & filename, 1
        End If
        oStream.Close
    End If

End Function

Function wgetDownload(cookieName, LogInString, AuthURL, filePath, filename, downloadUrl, Optional ExistingCookieFullPath = "")

    'Windows Script object parameters

    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, DownloadRequest, AutheticationRequest, Request As String

    'wget Paramenters - Wget must be installed and in "User path"

    BaseRequest = "cmd.exe /S /C "

    If ExistingCookieFullPath = "" Then
        Poststring = "wget --no-check-certificate --cookies=on --save-cookies " & cookieName & ".txt --keep-session-cookies --post-data " & """" & LogInString & """" & " " & """" & AuthURL & """"
        GetString = "wget --no-check-certificate --cookies=on --load-cookies " & cookieName & ".txt --keep-session-cookies -O """ & filePath & filename & """ " & """" & downloadUrl & """"
        Request = BaseRequest & Poststring & " && " & GetString
    Else
        Poststring = ""
        GetString = "wget --no-check-certificate --cookies=on --load-cookies " & ExistingCookieFullPath & " --keep-session-cookies -O """ & filePath & filename & " " & """" & " " & """" & downloadUrl & """"
        Request = BaseRequest & GetString
    End If

    Debug.Print Request

    Set WScr = VBA.CreateObject("WScript.Shell")

    WScr.Run Request, windowStyle, waitOnReturn
    'WScr.Run DownloadRequest, windowStyle, waitOnReturn


End Function
Function ExportEdgeCookies_old(DomainName As String) As String
    Dim conn As Object
    Dim rs As Object
    Dim outputPath As String
    Dim fso As Object
    Dim txtFile As Object
    Dim cookiesPath As String
    
    ' Potential cookie database locations
    Dim cookiesLocations(1) As String
    cookiesLocations(0) = Environ$("LOCALAPPDATA") & "\Microsoft\Edge\User Data\Default\Network\Cookies"
    cookiesLocations(1) = Environ$("LOCALAPPDATA") & "\Microsoft\Edge\User Data\Default\Cookies"
    
    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Find existing cookies database
    Dim i As Integer
    Dim foundCookiesPath As Boolean
    foundCookiesPath = False
    
    For i = 0 To UBound(cookiesLocations)
        If fso.FileExists(cookiesLocations(i)) Then
            cookiesPath = cookiesLocations(i)
            foundCookiesPath = True
            Exit For
        End If
    Next i
    
    ' Exit if no cookies database found
    If Not foundCookiesPath Then
        MsgBox "Cookies database not found in expected locations!"
        ExportEdgeCookies = ""
        Exit Function
    End If
    
    ' Replace dot with dash for filename
    Dim formattedDomainName As String
    formattedDomainName = Replace(DomainName, ".", "-")
    
    ' Set output path
    outputPath = "D:\Cloud\OneDrive\Cookies\" & formattedDomainName & ".txt"
    
    ' Create output text file
    Set txtFile = fso.CreateTextFile(outputPath, True)
    
    ' Write Netscape cookie file header
    txtFile.WriteLine "# Netscape HTTP Cookie File"
    txtFile.WriteLine "# http://curl.haxx.se/rfc/cookie_spec.html"
    txtFile.WriteLine "# Generated by VBA Script"
    
    ' Open SQLite database using ADO
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Driver={SQLite3 ODBC Driver};Database=" & cookiesPath & ";"
    conn.Open
    
    ' Execute query to get all columns explicitly
    Set rs = conn.Execute("SELECT * FROM cookies WHERE host_key LIKE '%" & DomainName & "%'")
    
    ' Write Netscape cookie format
    Do While Not rs.EOF
        Dim cookieLine As String
        Dim Domain As String
        Dim isHttpOnly As String
        Dim ExpiresUTC As Long
        Dim isSecure As Boolean
        
        ' Safely retrieve field values
        On Error Resume Next
        
        ' Determine domain
        Domain = CStr(rs.Fields("host_key").value)
        If Err.Number <> 0 Then
            Err.Clear
            rs.MoveNext
            GoTo NextIteration
        End If
        
        isHttpOnly = IIf(Left(Domain, 1) = ".", "#HttpOnly_" & Domain, "#HttpOnly_" & Domain)
        
        ' Safely get path
        Dim cookiePath As String
        cookiePath = CStr(rs.Fields("path").value)
        If Err.Number <> 0 Then
            Err.Clear
            cookiePath = "/"
        End If
        
        ' Safely get expires_utc
        On Error Resume Next
        ExpiresUTC = CLng(rs.Fields("expires_utc").value / 1000000 - 11644473600#)
        If Err.Number <> 0 Then
            Err.Clear
            ExpiresUTC = CLng(DateDiff("s", "1970-01-01", DateAdd("d", 30, Now())) + 1)
        End If
        
        ' Safely get is_secure
        On Error Resume Next
        isSecure = CBool(rs.Fields("is_secure").value)
        If Err.Number <> 0 Then
            Err.Clear
            isSecure = False
        End If
        
        ' Safely get name and value with more explicit retrieval
        Dim cookieName As String
        Dim cookieValue As String
        
        On Error Resume Next
        cookieName = CStr(rs.Fields("name").value)
        cookieValue = CStr(rs.Fields("encrypted_value").value)
        
        ' If encrypted_value is empty, try value column
        If cookieValue = "" Then
            cookieValue = CStr(rs.Fields("value").value)
        End If
        
        If Err.Number <> 0 Or cookieValue = "" Then
            Err.Clear
            rs.MoveNext
            GoTo NextIteration
        End If
        
        ' Construct cookie line
        cookieLine = isHttpOnly & vbTab & _
                     "TRUE" & vbTab & _
                     cookiePath & vbTab & _
                     IIf(isSecure, "TRUE", "FALSE") & vbTab & _
                     CStr(ExpiresUTC) & vbTab & _
                     cookieName & vbTab & _
                     cookieValue
        
        ' Write the cookie line
        txtFile.WriteLine cookieLine
        
NextIteration:
        rs.MoveNext
        On Error GoTo 0
    Loop
    
    ' Close everything
    txtFile.Close
    rs.Close
    conn.Close
    
    Set txtFile = Nothing
    Set rs = Nothing
    Set conn = Nothing
    Set fso = Nothing
    
    ExportEdgeCookies = outputPath
End Function

Sub Test()
ExportEdgeCookies "investing.com"
End Sub

Function Curl_CookieFile_Download(ByVal URI As String, ByVal CookieFileFullPath As String, ByVal filePath As String, ByVal filename As String, Optional ByVal allowOverwrite As Boolean = False) As String
    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, FullPathString As String
    Dim Request As String
    Dim FinalFileName As String
    
    ' If not allowing overwrite, modify filename with datetime suffix
    If Not allowOverwrite Then
        Dim FileExt As String
        Dim FileNameNoExt As String
        Dim DateTimeSuffix As String
        
        ' Generate datetime suffix
        DateTimeSuffix = "_" & Format(Now, "YYYYMMDD-HHMMSS")
        
        ' Split filename and extension
        If InStrRev(filename, ".") > 0 Then
            FileNameNoExt = Left(filename, InStrRev(filename, ".") - 1)
            FileExt = Mid(filename, InStrRev(filename, "."))
        Else
            FileNameNoExt = filename
            FileExt = ""
        End If
        
        ' Create new filename with datetime suffix
        FinalFileName = FileNameNoExt & DateTimeSuffix & FileExt
    Else
        FinalFileName = filename
    End If
    
    BaseRequest = "cmd /S /C "
    FullPathString = "curl """ & URI & """ " & _
        "-b """ & CookieFileFullPath & """ " & _
        "-H ""accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"" " & _
        "-H ""accept-language: en,it;q=0.9,en-GB;q=0.8,en-US;q=0.7,it-CH;q=0.6,de-CH;q=0.5,de;q=0.4,fr-FR;q=0.3,fr-CH;q=0.2,fr;q=0.1,de-DE;q=0.1,hr;q=0.1"" " & _
        "-H ""referer: https://www.investing.com/portfolio/?portfolioID=Z2M+aTVjPmthNGhnMGFlYw=="" " & _
        "-H ""user-agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"" " & _
        "-o """ & filePath & "\" & FinalFileName & """"
    
    Request = BaseRequest & FullPathString
    Debug.Print Request
    
    Set WScr = VBA.CreateObject("WScript.Shell")
    WScr.Run Request, windowStyle, waitOnReturn
    
    ' Return the final filename used
    Curl_CookieFile_Download = FinalFileName
End Function

Function Aria2c_CookieFile_Download(URI, CookieFileFullPath, filePath, filename, allowOverwrite)

    'Windows Script object parameters

    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, FullPathString As String
    Dim Request As String

    'Aria2 Paramenters - Aria2c must be installed and in "User path"

    BaseRequest = "cmd-aria.exe /S /C "
    FullPathString = "aria2c --load-cookies " & """" & CookieFileFullPath & """" & " --dir " & """" & filePath & """" & " --out " & """" & filename & """" & " --allow-overwrite=" & """" & allowOverwrite & """" & " " & """" & URI & """"
    '& " - -allow - Overwrite = " & """ & allowOverwrite & """"
    Request = BaseRequest & FullPathString

    Debug.Print Request

    Set WScr = VBA.CreateObject("WScript.Shell")

    WScr.Run Request, windowStyle, waitOnReturn

End Function

Function Test_CookieFile()
    Aria2c_CookieFile_Download "https://boerse.raiffeisen.ch/raiffeisen/export/export.xls", "D:\Stefano\Cloud\OneDrive\Cookies\RFborse.txt", "B:\Posizioni", "RFBorse.xls", "true"
End Function

Public Function Aria2c_CliGet_Download(CliGetString, filePath, filename, allowOverwrite As Boolean)

    'Windows Script object parameters

    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, FullPathString As String

    'Aria2 Paramenters - Aria2c must be installed and in "User path"

    BaseRequest = "cmd-aria.exe /S /C "

    FullPathString = " --dir " & filePath & " --out " & filename & " --allow-overwrite=" & allowOverwrite
    Request = BaseRequest & CliGetString & " && " & FullPathString

    Debug.Print Request

    Set WScr = VBA.CreateObject("WScript.Shell")

    WScr.Run Request, windowStyle, waitOnReturn

End Function

Public Function Aria2c_InputFile_Download(InputFileFullPath As String)

    'Windows Script object parameters

    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, Aria2cString As String

    'Aria2 Paramenters - Aria2c must be installed and in "User path"

    BaseRequest = "cmd.exe /S /C "
    Aria2cString = "aria2c --input-file "
    Request = BaseRequest & Aria2cString & """" & InputFileFullPath & """"

    Debug.Print Request

    Set WScr = VBA.CreateObject("WScript.Shell")

    WScr.Run Request, windowStyle, waitOnReturn

End Function

Public Function RunAgent(AgentFilePath As String)

    'Windows Script object parameters

    Dim WScr As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim BaseRequest, RunAgentString, Request As String

    'RunAgent.exe must be installed and in "User path"

    BaseRequest = "cmd.exe /S /C "
    RunAgentString = "RunAgent.exe " & AgentFilePath
    Request = BaseRequest & RunAgentString
    
    Debug.Print Request

    Set WScr = VBA.CreateObject("WScript.Shell")

    WScr.Run Request, windowStyle, waitOnReturn

End Function
Private Function Test_Agent()
Dim response As Integer
response = RunAgent("D:\Cloud\OneDrive\Agents\MS_Containers_Components\MS_Containers_Components")
Debug.Print response
End Function

Private Function TestIputFile()
    Aria2c_InputFile_Download "B:\Posizioni\AriaURLs.txt"
End Function

Private Function TestInvesting()
    wgetDownload _
        cookieName:="investing_com", _
        LogInString:="email=morniroli.stefano%40gmail.com&password=inver68filtro&logintoken=ef41aa5524cd55f8eca4ee4a601cb951&onAuthCompleteAction=%7B%22type%22%3A%22topBar%22%2C%22location%22%3A%7B%22path%22%3A%22%2F%22%2C%22search%22%3A%22%22%2C%22hash%22%3A%22%22%7D%2C%22mmID%22%3A1%7D", _
        AuthURL:="https://www.investing.com/members-admin/auth/signInByEmail/", _
        filePath:="B:\Posizioni\", _
        filename:="InvCom-Cur.csv", _
        downloadUrl:="https://www.investing.com/portfolio/service/Export?portfolio_id=3083029&exptype=csv&selected_tab=overview"
End Function

Private Function TestFT()
    wgetDownload _
        cookieName:="Ft_com", _
        LogInString:="_csrf=fwHO5mk3-Bp6hd7KfDPaBAZ3rvXcYpOx9TwQ&formType=enter-email&location=https%3A%2F%2Fwww.ft.com&email=morniroli.stefano%40gmail.com&_notUsed=ftcom%26r6filtro", _
        AuthURL:="https://accounts.ft.com/login?location=https%3A%2F%2Fwww.ft.com", _
        filePath:="B:\Posizioni\", _
        filename:="FtmCom.csv", _
        downloadUrl:="https://markets.ft.com/data/portfolio/ajax/export?data={%22type%22:%22Portfolio%22,%22containerId%22:%2236237a43-f400-4b54-8fa4-bf09b2986038%22,%22viewName%22:%225543e6de-0e5b-4ca6-80fe-719872039c82%22,%22startDate%22:%222018/06/17%22,%22endDate%22:%222018/07/02%22}", _
        ExistingCookieFullPath:="D:\Stefano\Cloud\OneDrive\Cookies\Ft-com.txt"

End Function

Function HTTPRequest( _
    Agent As String, _
    URL As String, _
    Format As String, _
    Optional ByVal RequestHeader = Nothing, _
    Optional ByVal RequestBody As String, _
    Optional User As String, _
    Optional password As String, _
    Optional filename = "", _
    Optional filePath = "") _
As Object

'If you need to pass the full HTML of the web page, read this
'https://www.ozgrid.com/forum/index.php?thread/149293-web-scraping-city-code-question-specific-site/

Dim Method As String
Dim response As Object
Dim http As MSXML2.ServerXMLHTTP60
Dim header As Variant ' Iterator
Dim HTMLDoc As Object



'Manage Input Parmeters
If RequestBody <> "" Then Method = "POST" Else Method = "GET"
If filePath = "" And filename <> "" Then
    filePath = MORFunctions.Informational.SpecialFolderPath("Desktop")
End If


'Dclare and Instanciate Relavant Response collectors
If Format = "html" Then
    'Late binding; Early Binding doesn't support the write method.
    'For further information refer to: https://www.ozgrid.com/forum/index.php?thread/149293-web-scraping-city-code-question-specific-site/
    Set HTMLDoc = CreateObject("HTMLfile")
ElseIf Format = "json" Then
    'Prepare for Late Binding of response
    Dim JSON As Object
Else
     MsgBox "The provided Format is not supported." & Chr(10) & "Try 'html' or 'json'"
End If

'Determine request Agent
If Agent = "XMLHttp" Then
    'Early Binding; you need to have a reference to Microsoft XML, Version 6.0
    Set http = New MSXML2.ServerXMLHTTP60
    
    'Make the request with .Open method
    'the syntax is: xmlhttp.Open Method, URL, async(true or false)
    http.Open Method, URL, False
    
    'Pass request Headers (if needed)
    'Request Header should be passed as a single string or a 2 dim collection
    For Each header In RequestHeader
        http.SetRequestHeader header(0), header(1)
    Next header
    
    'Additional Headers for POST request
    If Method = "POST" Then
        RequestHeader.Add Array("Accept-Encoding", "gzip, deflate, br")
    End If
    
    'http.setRequestHeader "Content-Type", "text/json"
    'http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    'http.setRequestHeader "User-Agent", "Mozilla/5.0 (iPad; U; CPU OS 3_2_1 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Mobile/7B405"
    
    
    'Pass log-in Information if Simple Authentication is supported and required
    'Log in information must be provided base 64 encoded;
    'In this file there is a utility function to encode in base 64
    If User <> "" Then
        http.SetRequestHeader "Authorization", "Basic " + encodeBase64(User + ":" + password)
    End If
    
    http.Send   'Send request
    
    Do Until http.readyState = 4    ' Wait for completion
        DoEvents
    Loop
    
    'Check if error
    If http.Status <> 200 Then
        MsgBox http.Status & " - " & http.statusText
        Exit Function
    End If
    Set response = http

ElseIf Agent = "IE" Then

    'Early binding to Internet Explorer Agent: yu need t add reference to "Microsoft Internet Controls"
    'If yoiu prefer late Binding: non eference and substitute following line with
    'Set http = CreateObject("InternetExplorer.Application")
    Dim ie As New InternetExplorer
    'ie request
    ie.Visible = True
    'navigate to website
    ie.navigate URL
    'Wait till Page loads
    Do While ie.readyState = 4: DoEvents: Loop
    Do Until ie.readyState = 4: DoEvents: Loop
    While ie.Busy
        DoEvents
    Wend
    Set response = ie.document
Else
    MsgBox "The requested Agent is not supported." & Chr(10) & "Try 'XMLHttp' or 'IE'"
    GoTo fine
End If

If Format = "html" Then
    HTMLDoc.Open
    If Agent = "XMLHttp" Then
        HTMLDoc.Write response.responseText
    ElseIf Agent = "IE" Then
        Set HTMLDoc = response
        'HTMLDoc.Write Response.body.innerHTML 'Response.body.outerHTML & Response.body.innerHTML
    Else
        MsgBox "I could not save the HTMLDoc file"
    End If
    HTMLDoc.Close
    Set HTTPRequest = HTMLDoc
    If filename <> "" Then 'the file should be saved
        filename = filename & "." & Format
        DumpText2File HTMLDoc.DocumentElement.innerHTML, filePath & "\" & filename
        OpenFile (filePath & "\" & filename)
    End If
    
ElseIf Format = "json" Then
    Set JSON = ParseJson(response.responseText)
    Set HTTPRequest = JSON
    If filename <> "" Then 'the file should be saved
        filename = filename & "." & Format
        DumpText2File response.responseText, filePath & "\" & filename
        OpenFile (filePath & "\" & filename)
    End If
Else
    MsgBox "The specified Format is not supported"
    GoTo fine
End If

fine:
   
End Function

Function Test_JSON_Download()
Dim RequestHeader As New Collection

'Set Header Collection
RequestHeader.Add Array("Content-Type", "text/json")
RequestHeader.Add Array("Accept", "application/json, text/javascript, */*; q=0.01")
RequestHeader.Add Array("Accept-Language", "it-CH,en;q=0.5")
RequestHeader.Add Array("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
RequestHeader.Add Array("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:75.0) Gecko/20100101 Firefox/75.0")


HTTPRequest _
    Agent:="XMLHttp", _
    URL:="https://query2.finance.yahoo.com/v8/finance/chart/BCVN.SW?&interval=1d&period1=1554940800&period2=1586563200&events=split%7Csplit&corsDomain=finance.yahoo.com", _
    Format:="json", _
    RequestHeader:=RequestHeader, _
    RequestBody:="", _
    User:="", _
    password:="", _
    filename:="Test", _
    filePath:=""
    
'HTTPRequest _
'    Agent:="IE", _
'    URL:="https://www.investing.com/equities/novartis-historical-data", _
'    Format:="html", _
'    RequestHeader:=RequestHeader, _
'    RequestBody:="", _
'    User:="", _
'    Password:="", _
'    Filename:="Test", _
'    Filepath:=""
    
'Yahoo finance Price History (json)
'https://query2.finance.yahoo.com/v8/finance/chart/BCVN.SW?formatted=true&crumb=%2Ffo1BA5Ov7w&lang=en-US&region=US&interval=1d&period1=1554940800&period2=1586563200&events=split%7Csplit&corsDomain=finance.yahoo.com"
'https://query2.finance.yahoo.com/v8/finance/chart/BCVN.SW?&fields=volume&interval=1d&period1=1554940800&period2=1586563200&events=split%7Csplit&corsDomain=finance.yahoo.com"


'Yahoo finance Valuation mesures
'https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/BCVN.SW?lang=en-US&region=US&symbol=BCVN.SW&padTimeSeries=true&type=quarterlyMarketCap%2CtrailingMarketCap%2CquarterlyEnterpriseValue%2CtrailingEnterpriseValue%2CquarterlyPeRatio%2CtrailingPeRatio%2CquarterlyForwardPeRatio%2CtrailingForwardPeRatio%2CquarterlyPegRatio%2CtrailingPegRatio%2CquarterlyPsRatio%2CtrailingPsRatio%2CquarterlyPbRatio%2CtrailingPbRatio%2CquarterlyEnterprisesValueRevenueRatio%2CtrailingEnterprisesValueRevenueRatio%2CquarterlyEnterprisesValueEBITDARatio%2CtrailingEnterprisesValueEBITDARatio&merge=false&period1=493590046&period2=1586636203&corsDomain=finance.yahoo.com



'Yahoo finance statistics
'https://query2.finance.yahoo.com/v7/finance/quote?formatted=false&crumb=%2Ffo1BA5Ov7w&lang=en-US&region=US&symbols=BCVN.SW&fields=messageBoardId%2ClongName%2CshortName%2CmarketCap%2CunderlyingSymbol%2CunderlyingExchangeSymbol%2CheadSymbolAsString%2CregularMarketPrice%2CregularMarketChange%2CregularMarketChangePercent%2CregularMarketVolume%2Cuuid%2CregularMarketOpen%2CfiftyTwoWeekLow%2CfiftyTwoWeekHigh%2CtoCurrency%2CfromCurrency%2CtoExchange%2CfromExchange&corsDomain=finance.yahoo.com

'Full Price History - Json Response
'https://query2.finance.yahoo.com/v8/finance/chart/BCVN.SW?formatted=true&crumb=%2Ffo1BA5Ov7w&lang=en-US&region=US&interval=1d&period1=1554940800&period2=1586563200&events=split%7Csplit&corsDomain=finance.yahoo.com


'Price hostory
'"https://it.finance.yahoo.com/quote/NOVN.SW/history?period1=1577836800&period2=1584921600&interval=1d&filter=history&frequency=1d", _
    "NOVN_Yahoo.html"

'&events=div&events=split&events=earn

End Function

Function Test_HTML_Download()

Dim RequestHeader As New Collection

'Set Header Collection
RequestHeader.Add Array("Content-Type", "text/html;charset=UTF-8")
RequestHeader.Add Array("Accept", "text/html, */*; q=0.01")
RequestHeader.Add Array("Accept-Language", "it-CH,en;q=0.5")
RequestHeader.Add Array("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:75.0) Gecko/20100101 Firefox/75.0")
'-------
RequestHeader.Add Array("Referer", "https://www.marketscreener.com")
RequestHeader.Add Array("DNT", "1")
RequestHeader.Add Array("Connection", "keep-alive")
RequestHeader.Add Array("Upgrade-Insecure-Requests", "1")

HTTPRequest _
    Agent:="XMLHttp", _
    URL:="https://www.marketscreener.com/search/instruments/?aComposeInputSearch=s_KR7005930003", _
    Format:="html", _
    RequestHeader:=RequestHeader, _
    RequestBody:="", _
    User:="", _
    password:="", _
    filename:="Test", _
    filePath:=""

End Function
