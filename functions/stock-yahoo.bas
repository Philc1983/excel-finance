' (c) 2013 Fincluster (fincluster.com)
' Usage
'	=GetTickerData(TICKER, DATE, FIELD)
'   Example
'		=GetTickerData(AAPL, "2013-1-7", "Close")

Public Function GetTickerData(ticker As String, qDate As Date, qField As String)
    Dim occUrl As String
     occUrl = "http://ichart.finance.yahoo.com/table.csv?s=" & ticker & _
        "&a=" & (Month(qDate) - 1) & "&b=" & Day(qDate) & "&c=" & Year(qDate) & _
        "&d=" & (Month(qDate) - 1) & "&e=" & Day(qDate) & "&f=" & Year(qDate) & _
        "&g=d&ignore=.csv"
    Debug.Print occUrl
    
    Dim tableText As String
    
    tableText = HTTPGet(occUrl, "")
    
    Dim lines() As String, fields() As String, values() As String
    Dim nOfCols As Integer, nOfRows As Integer, i As Integer
    lines = Split(tableText, vbLf)
    nOfCols = UBound(Split(lines(0), ","))
  nOfRows = UBound(lines) - 1
    Dim result As String
    
    if nOfRows > 0 Then
      fields = Split(lines(0), ",")
      values = Split(lines(1), ",")
      For i = 0 To nOfCols
          If fields(i) = qField Then
              result = values(i)
          End If
      Next i
    End If
    GetTickerData = result
End Function

Public Function HTTPGet(sUrl As String, sQuery As String) As String
    Dim sResult As String
	Dim xml As Object
	Set xml = CreateObject("Microsoft.XMLHTTP")
	xml.Open "GET", sUrl, False
	xml.send
	sResult = xml.ResponseText
	Set xml = Nothing
    HTTPGet = sResult
End Function
