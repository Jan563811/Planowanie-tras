Attribute VB_Name = "ExportReadablePairs"
Option Explicit

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef phHash As LongPtr, _
    ByRef pbHashObject As Any, _
    ByVal cbHashObject As Long, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbInput As Any, _
    ByVal cbInput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Type GeoPoint
    Address As String
    Latitude As Double
    Longitude As Double
End Type

Private Const CACHE_URL As String = _
    "https://raw.githubusercontent.com/Jan563811/Planowanie-tras/main/dm_pairs_cache.json"
Private Const OUTPUT_SHEET As String = "PairsReadable"
Private Const BUFFER_ROWS As Long = 2000

Public Sub ExportReadablePairs()
    Dim sourceTable As ListObject
    Dim addressColumn As Long, latitudeColumn As Long, longitudeColumn As Long
    Dim statusColumn As Long
    Dim points() As GeoPoint
    Dim pointCount As Long, rowIndex As Long
    Dim rowRange As Range
    Dim cache As Object
    Dim outputSheet As Worksheet
    Dim algorithm As LongPtr, hashObjectLength As Long
    Dim outputBuffer() As Variant, bufferedRows As Long, outputRow As Long
    Dim pointA As Long, pointB As Long
    Dim raw As String, pairHash As String
    Dim cached As Variant, distanceMeters As Variant, durationSeconds As Variant
    Dim oldScreenUpdating As Boolean, oldStatusBar As Variant
    Dim errorMessage As String

    oldScreenUpdating = Application.ScreenUpdating
    oldStatusBar = Application.StatusBar
    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set sourceTable = FindGeocodingTable(addressColumn, latitudeColumn, longitudeColumn, statusColumn)
    If sourceTable Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
            "Nie znalazłem tabeli Excela z kolumnami address, lat i lng. Załaduj wynik geokodowania do arkusza jako tabelę."
    End If
    If sourceTable.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1001, , "Tabela geokodowania nie zawiera wierszy."
    End If

    ReDim points(1 To sourceTable.DataBodyRange.Rows.Count)
    For rowIndex = 1 To sourceTable.DataBodyRange.Rows.Count
        Set rowRange = sourceTable.DataBodyRange.Rows(rowIndex)
        If IsAcceptedStatus(rowRange, statusColumn) Then
            If IsUsableNumber(rowRange.Cells(1, latitudeColumn).Value2) And _
               IsUsableNumber(rowRange.Cells(1, longitudeColumn).Value2) Then
                pointCount = pointCount + 1
                points(pointCount).Address = CStr(rowRange.Cells(1, addressColumn).Value2)
                points(pointCount).Latitude = ParseCoordinate(rowRange.Cells(1, latitudeColumn).Value2)
                points(pointCount).Longitude = ParseCoordinate(rowRange.Cells(1, longitudeColumn).Value2)
            End If
        End If
    Next rowIndex
    If pointCount < 2 Then Err.Raise vbObjectError + 1002, , "Potrzeba co najmniej dwóch poprawnych punktów."
    ReDim Preserve points(1 To pointCount)

    Application.StatusBar = "Pobieranie cache tras z GitHuba..."
    Set cache = DownloadDistanceCache()
    If cache.Count = 0 Then Err.Raise vbObjectError + 1003, , "Pobrany cache nie zawiera rozpoznanych rekordów."

    OpenSha256Provider algorithm, hashObjectLength
    Set outputSheet = GetOrCreateOutputSheet()
    PrepareOutputSheet outputSheet
    WriteHeaders outputSheet
    ReDim outputBuffer(1 To BUFFER_ROWS, 1 To 11)
    outputRow = 2

    For pointA = 1 To pointCount
        Application.StatusBar = "Dopasowywanie hashy: punkt " & pointA & " / " & pointCount
        For pointB = 1 To pointCount
            If pointA <> pointB Then
                raw = CoordinateText(points(pointA).Latitude) & "," & _
                      CoordinateText(points(pointA).Longitude) & "|" & _
                      CoordinateText(points(pointB).Latitude) & "," & _
                      CoordinateText(points(pointB).Longitude)
                pairHash = Sha256Hex(raw, algorithm, hashObjectLength)

                If cache.Exists(pairHash) Then
                    cached = cache(pairHash)
                    distanceMeters = cached(0)
                    durationSeconds = cached(1)
                    bufferedRows = bufferedRows + 1
                    outputBuffer(bufferedRows, 1) = points(pointA).Address
                    outputBuffer(bufferedRows, 2) = points(pointA).Latitude
                    outputBuffer(bufferedRows, 3) = points(pointA).Longitude
                    outputBuffer(bufferedRows, 4) = points(pointB).Address
                    outputBuffer(bufferedRows, 5) = points(pointB).Latitude
                    outputBuffer(bufferedRows, 6) = points(pointB).Longitude
                    outputBuffer(bufferedRows, 7) = distanceMeters
                    If IsNumeric(distanceMeters) Then outputBuffer(bufferedRows, 8) = CDbl(distanceMeters) / 1000
                    outputBuffer(bufferedRows, 9) = durationSeconds
                    If IsNumeric(durationSeconds) Then outputBuffer(bufferedRows, 10) = CDbl(durationSeconds) / 60
                    outputBuffer(bufferedRows, 11) = pairHash

                    If bufferedRows = BUFFER_ROWS Then
                        FlushOutput outputSheet, outputBuffer, bufferedRows, outputRow
                        outputRow = outputRow + bufferedRows
                        bufferedRows = 0
                    End If
                End If
            End If
        Next pointB
    Next pointA

    If bufferedRows > 0 Then
        FlushOutput outputSheet, outputBuffer, bufferedRows, outputRow
        outputRow = outputRow + bufferedRows
    End If

    Call BCryptCloseAlgorithmProvider(algorithm, 0)
    algorithm = 0
    FormatOutput outputSheet, outputRow - 1
    Application.StatusBar = oldStatusBar
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Gotowe. Dopasowano " & (outputRow - 2) & " par. Wynik jest w arkuszu '" & OUTPUT_SHEET & "'.", vbInformation
    Exit Sub

Failed:
    errorMessage = Err.Description
    If algorithm <> 0 Then Call BCryptCloseAlgorithmProvider(algorithm, 0)
    Application.StatusBar = oldStatusBar
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Makro nie zakończyło pracy: " & errorMessage, vbExclamation
End Sub

Private Function FindGeocodingTable( _
    ByRef addressColumn As Long, _
    ByRef latitudeColumn As Long, _
    ByRef longitudeColumn As Long, _
    ByRef statusColumn As Long) As ListObject

    Dim sheet As Worksheet, table As ListObject
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            addressColumn = FindColumn(table, "address")
            latitudeColumn = FindColumn(table, "lat")
            longitudeColumn = FindColumn(table, "lng")
            statusColumn = FindColumn(table, "status")
            If addressColumn > 0 And latitudeColumn > 0 And longitudeColumn > 0 Then
                Set FindGeocodingTable = table
                Exit Function
            End If
        Next table
    Next sheet
End Function

Private Function FindColumn(ByVal table As ListObject, ByVal columnName As String) As Long
    Dim index As Long
    For index = 1 To table.ListColumns.Count
        If StrComp(Trim$(table.ListColumns(index).Name), columnName, vbTextCompare) = 0 Then
            FindColumn = index
            Exit Function
        End If
    Next index
End Function

Private Function IsUsableNumber(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    IsUsableNumber = (Len(Trim$(CStr(value))) > 0)
End Function

Private Function IsAcceptedStatus(ByVal rowRange As Range, ByVal statusColumn As Long) As Boolean
    If statusColumn = 0 Then
        IsAcceptedStatus = True
    Else
        IsAcceptedStatus = (UCase$(Trim$(CStr(rowRange.Cells(1, statusColumn).Value2))) = "OK")
    End If
End Function

Private Function ParseCoordinate(ByVal value As Variant) As Double
    Dim text As String
    If IsNumeric(value) Then
        ParseCoordinate = CDbl(value)
    Else
        text = Trim$(CStr(value))
        text = Replace(text, ".", Application.DecimalSeparator)
        ParseCoordinate = CDbl(text)
    End If
End Function

Private Function CoordinateText(ByVal value As Double) As String
    CoordinateText = Format$(value, "0.0000000")
    CoordinateText = Replace(CoordinateText, Application.DecimalSeparator, ".")
End Function

Private Function DownloadDistanceCache() As Object
    Dim request As Object, expression As Object, matches As Object, item As Object
    Dim cache As Object, body As String, distanceValue As Variant, durationValue As Variant

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "GET", CACHE_URL, False
    request.Send
    If request.Status <> 200 Then
        Err.Raise vbObjectError + 1010, , "GitHub zwrócił HTTP " & request.Status & "."
    End If
    body = request.ResponseText

    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.IgnoreCase = True
    expression.Pattern = """([0-9a-f]{64})""\s*:\s*\{\s*""d""\s*:\s*(\d+|null)\s*,\s*""t""\s*:\s*(\d+|null)\s*\}"
    Set matches = expression.Execute(body)

    Set cache = CreateObject("Scripting.Dictionary")
    cache.CompareMode = vbTextCompare
    For Each item In matches
        distanceValue = JsonNumber(item.SubMatches(1))
        durationValue = JsonNumber(item.SubMatches(2))
        cache(item.SubMatches(0)) = Array(distanceValue, durationValue)
    Next item
    Set DownloadDistanceCache = cache
End Function

Private Function JsonNumber(ByVal value As String) As Variant
    If LCase$(value) = "null" Then
        JsonNumber = Empty
    Else
        JsonNumber = CDbl(value)
    End If
End Function

Private Sub OpenSha256Provider(ByRef algorithm As LongPtr, ByRef objectLength As Long)
    Dim algorithmName As String, propertyName As String, bytesReturned As Long
    Dim result As Long

    algorithmName = "SHA256"
    propertyName = "ObjectLength"
    result = BCryptOpenAlgorithmProvider(algorithm, StrPtr(algorithmName), 0, 0)
    If result <> 0 Then Err.Raise vbObjectError + 1020, , "Windows nie uruchomił SHA256 (status " & result & ")."

    result = BCryptGetProperty(algorithm, StrPtr(propertyName), objectLength, 4, bytesReturned, 0)
    If result <> 0 Or objectLength <= 0 Then
        Call BCryptCloseAlgorithmProvider(algorithm, 0)
        algorithm = 0
        Err.Raise vbObjectError + 1021, , "Nie udało się odczytać parametrów SHA256."
    End If
End Sub

Private Function Sha256Hex(ByVal text As String, ByVal algorithm As LongPtr, ByVal objectLength As Long) As String
    Dim hashObject() As Byte, inputBytes() As Byte, outputBytes(0 To 31) As Byte
    Dim hashHandle As LongPtr, result As Long, index As Long

    ReDim hashObject(0 To objectLength - 1)
    inputBytes = StrConv(text, vbFromUnicode)
    result = BCryptCreateHash(algorithm, hashHandle, hashObject(0), objectLength, 0, 0, 0)
    If result <> 0 Then Err.Raise vbObjectError + 1030, , "Nie udało się utworzyć hasha SHA256."

    result = BCryptHashData(hashHandle, inputBytes(0), UBound(inputBytes) + 1, 0)
    If result = 0 Then result = BCryptFinishHash(hashHandle, outputBytes(0), 32, 0)
    Call BCryptDestroyHash(hashHandle)
    If result <> 0 Then Err.Raise vbObjectError + 1031, , "Nie udało się obliczyć hasha SHA256."

    For index = 0 To 31
        Sha256Hex = Sha256Hex & Right$("0" & LCase$(Hex$(outputBytes(index))), 2)
    Next index
End Function

Private Function GetOrCreateOutputSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateOutputSheet = ThisWorkbook.Worksheets(OUTPUT_SHEET)
    On Error GoTo 0
    If GetOrCreateOutputSheet Is Nothing Then
        Set GetOrCreateOutputSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateOutputSheet.Name = OUTPUT_SHEET
    End If
End Function

Private Sub PrepareOutputSheet(ByVal sheet As Worksheet)
    Dim index As Long
    For index = sheet.ListObjects.Count To 1 Step -1
        sheet.ListObjects(index).Delete
    Next index
    sheet.Cells.Clear
End Sub

Private Sub WriteHeaders(ByVal sheet As Worksheet)
    Dim headers(1 To 1, 1 To 11) As Variant
    headers(1, 1) = "address_a"
    headers(1, 2) = "lat_a"
    headers(1, 3) = "lng_a"
    headers(1, 4) = "address_b"
    headers(1, 5) = "lat_b"
    headers(1, 6) = "lng_b"
    headers(1, 7) = "distance_m"
    headers(1, 8) = "distance_km"
    headers(1, 9) = "duration_s"
    headers(1, 10) = "duration_min"
    headers(1, 11) = "hash"
    sheet.Range("A1").Resize(1, 11).Value = headers
End Sub

Private Sub FlushOutput( _
    ByVal sheet As Worksheet, _
    ByRef buffer() As Variant, _
    ByVal rowCount As Long, _
    ByVal firstRow As Long)

    If rowCount = BUFFER_ROWS Then
        sheet.Cells(firstRow, 1).Resize(rowCount, 11).Value = buffer
    Else
        Dim partial() As Variant, rowIndex As Long, columnIndex As Long
        ReDim partial(1 To rowCount, 1 To 11)
        For rowIndex = 1 To rowCount
            For columnIndex = 1 To 11
                partial(rowIndex, columnIndex) = buffer(rowIndex, columnIndex)
            Next columnIndex
        Next rowIndex
        sheet.Cells(firstRow, 1).Resize(rowCount, 11).Value = partial
    End If
End Sub

Private Sub FormatOutput(ByVal sheet As Worksheet, ByVal lastRow As Long)
    If lastRow < 1 Then Exit Sub
    With sheet.Range("A1:K1")
        .Font.Bold = True
    End With
    sheet.Range("A1:K" & lastRow).AutoFilter
    sheet.Columns("A:K").AutoFit
    If lastRow > 1 Then
        sheet.Range("H2:H" & lastRow).NumberFormat = "0.00"
        sheet.Range("J2:J" & lastRow).NumberFormat = "0.0"
        sheet.Range("K2:K" & lastRow).NumberFormat = "@"
    End If
End Sub