Attribute VB_Name = "LinkCheck"
Option Explicit

' API declaration for rate limiting
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Tunables
Private Const TIMEOUT_MS As Long = 10000
Private Const SHOW_STATUSBAR As Boolean = True
Private Const DEBUG_LOG As Boolean = True
Private Const CONTEXT_CHARS As Long = 10
Private Const MAX_LINKS_TO_CHECK As Long = 1000 ' Safety limit
Private Const DELAY_BETWEEN_REQUESTS As Long = 100 ' Milliseconds between unique URL checks
Private Const STATUSBAR_UPDATE_INTERVAL As Long = 5 ' Update status every N links

' HTTP timeout constants
Private Const RESOLVE_TIMEOUT As Long = 3000
Private Const CONNECT_TIMEOUT As Long = 4000
Private Const SEND_TIMEOUT As Long = 4000
Private Const RECEIVE_TIMEOUT As Long = 7000

Private mCache As Object
Private mCacheHits As Long
Private mUseHeadFirst As Boolean ' Try HEAD before GET for performance

Public Sub ValidateHyperlinks()
    Dim doc As Word.Document, wsNewDoc As Word.Document, tbl As Word.Table
    Dim results As Collection
    Dim totalCount As Long, processedCount As Long, uniqueCount As Long
    Dim fn As Word.Footnote, hl As Word.Hyperlink
    Dim item As Variant
    Dim lastErrNum As Long, lastErrDesc As String
    
    ' Word settings to restore
    Dim origScreenUpdating As Boolean, origDisplayAlerts As WdAlertLevel
    Dim origSpellCheck As Boolean, origGrammarCheck As Boolean

    On Error GoTo CleanFail

    Set doc = ActiveDocument

    ' Store original settings
    origScreenUpdating = Application.ScreenUpdating
    origDisplayAlerts = Application.DisplayAlerts
    With Application.Options
        origSpellCheck = .CheckSpellingAsYouType
        origGrammarCheck = .CheckGrammarAsYouType
    End With

    ' Optimize Word for performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    With Application.Options
        .CheckSpellingAsYouType = False
        .CheckGrammarAsYouType = False
    End With

    ' Initialize cache
    Set mCache = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    mCache.RemoveAll
    On Error GoTo CleanFail
    mCacheHits = 0
    mUseHeadFirst = True ' Enable HEAD optimization

    Set results = New Collection

    ' Count total for progress
    totalCount = doc.Hyperlinks.Count
    For Each fn In doc.Footnotes
        totalCount = totalCount + fn.Range.Hyperlinks.Count
    Next fn
    
    ' Safety check
    If totalCount > MAX_LINKS_TO_CHECK Then
        If MsgBox("Document contains " & totalCount & " links. This may take a while. Continue?", vbYesNo + vbQuestion, "Many Links Detected") = vbNo Then
            GoTo CleanExit
        End If
    End If

    If SHOW_STATUSBAR Then Application.StatusBar = "Validating hyperlinks (0/" & totalCount & ")..."

    ' Pre-collect all links with their metadata for deduplication
    Dim linkInfo As Object ' Dictionary to store unique URL -> collection of locations
    Set linkInfo = CreateObject("Scripting.Dictionary")
    
    Dim tmpItem As Object
    Dim normalizedAddr As String
    Dim locationInfo As Object
    
    ' Gather main body hyperlinks
    For Each hl In doc.Hyperlinks
        normalizedAddr = NormalizeURL(hl.Address)
        If Len(normalizedAddr) > 0 Then
            Set locationInfo = CreateObject("Scripting.Dictionary")
            locationInfo.Add "Range", hl.Range
            locationInfo.Add "DisplayText", hl.TextToDisplay
            locationInfo.Add "Address", hl.Address
            
            If Not linkInfo.Exists(normalizedAddr) Then
                Dim locCollection As Collection
                Set locCollection = New Collection
                linkInfo.Add normalizedAddr, locCollection
            End If
            linkInfo(normalizedAddr).Add locationInfo
        End If
    Next hl

    ' Gather footnote hyperlinks
    For Each fn In doc.Footnotes
        For Each hl In fn.Range.Hyperlinks
            normalizedAddr = NormalizeURL(hl.Address)
            If Len(normalizedAddr) > 0 Then
                Set locationInfo = CreateObject("Scripting.Dictionary")
                locationInfo.Add "Range", hl.Range
                locationInfo.Add "DisplayText", hl.TextToDisplay
                locationInfo.Add "Address", hl.Address
                
                If Not linkInfo.Exists(normalizedAddr) Then
                    Set locCollection = New Collection
                    linkInfo.Add normalizedAddr, locCollection
                End If
                linkInfo(normalizedAddr).Add locationInfo
            End If
        Next hl
    Next fn

    uniqueCount = linkInfo.Count
    If SHOW_STATUSBAR Then Application.StatusBar = "Found " & totalCount & " links (" & uniqueCount & " unique). Checking..."

    ' Now check each unique URL once and apply result to all instances
    Dim urls As Variant, url As Variant
    Dim checkResult As Object
    Dim loc As Variant
    Dim uniqueProcessed As Long
    
    urls = linkInfo.Keys
    uniqueProcessed = 0
    
    For Each url In urls
        ' Allow user cancellation
        If uniqueProcessed Mod 10 = 0 Then
            DoEvents
            If GetAsyncKeyState(27) < 0 Then ' ESC key
                If MsgBox("Cancel validation?", vbYesNo + vbQuestion, "Cancel?") = vbYes Then
                    GoTo Finalise
                End If
            End If
        End If
        
        ' Check this unique URL
        Set checkResult = CheckUniqueURL(CStr(url), uniqueProcessed, uniqueCount)
        
        ' Apply result to all instances of this URL
        For Each loc In linkInfo(url)
            Set tmpItem = CreateResultItem(doc, loc("Range"), CStr(loc("DisplayText")), CStr(loc("Address")), checkResult)
            If Not tmpItem Is Nothing Then
                results.Add tmpItem
                processedCount = processedCount + 1
            End If
        Next loc
        
        uniqueProcessed = uniqueProcessed + 1
        
        ' Update status bar less frequently
        If SHOW_STATUSBAR And (uniqueProcessed Mod STATUSBAR_UPDATE_INTERVAL = 0 Or uniqueProcessed = uniqueCount) Then
            Application.StatusBar = "Checked " & uniqueProcessed & "/" & uniqueCount & " unique URLs (" & processedCount & " total links)..."
        End If
        
        ' Rate limiting - only delay for unique URLs not in cache
        If Not mCache.Exists(CStr(url)) Then
            Sleep DELAY_BETWEEN_REQUESTS
        End If
    Next url

    ' Sort results: Red (invalid symbol), Yellow (other fail), Green (valid)
    Dim sorted As New Collection
    Dim colorOrder As Variant, clr As Variant
    colorOrder = Array("R", "Y", "G")

    For Each clr In colorOrder
        For Each item In results
            If Not item Is Nothing Then
                On Error Resume Next
                If item("ColorCode") = clr Then sorted.Add item
                If Err.Number <> 0 Then
                    Debug.Print "Sorting skip error: "; Err.Number; Err.Description
                    Err.Clear
                End If
                On Error GoTo CleanFail
            End If
        Next item
    Next clr

    ' Output to new document
    Set wsNewDoc = Documents.Add()

    If sorted.Count = 0 Then
        With wsNewDoc.Range
            .Text = "Hyperlink validation completed. No reportable links were found or all links were skipped." & vbCrLf & vbCrLf & _
                    "Diagnostic: Links processed = " & processedCount & ", unique URLs = " & uniqueCount & ", cache hits = " & mCacheHits
            .Font.Bold = True
        End With
    Else
        ' Build table with optimizations (including header row)
        On Error Resume Next
        Set tbl = wsNewDoc.Tables.Add(wsNewDoc.Range, sorted.Count + 1, 4)
        If Err.Number <> 0 Then
            Debug.Print "Table.Add failed: "; Err.Number; Err.Description
            Err.Clear
            ' Fallback: write plain text report
            Dim outRng As Word.Range
            Set outRng = wsNewDoc.Range
            outRng.Text = "Could not create table. Writing plain-text report." & vbCrLf & vbCrLf
            Dim si As Long
            si = 1
            For Each item In sorted
                outRng.InsertAfter "Row " & si & vbCrLf
                outRng.InsertAfter "Display: " & item("DisplayText") & vbCrLf
                outRng.InsertAfter "Address: " & item("Address") & vbCrLf
                outRng.InsertAfter "Status: " & item("StatusText") & vbCrLf & vbCrLf
                si = si + 1
            Next item
            GoTo Finalise
        End If
        On Error GoTo CleanFail

        ' Optimize table settings
        tbl.AllowAutoFit = False
        
        ' Add header row first
        On Error Resume Next
        tbl.Cell(1, 1).Range.Text = "Display Text"
        tbl.Cell(1, 2).Range.Text = "URL"
        tbl.Cell(1, 3).Range.Text = "Context"
        tbl.Cell(1, 4).Range.Text = "Status"
        tbl.Rows(1).HeadingFormat = True
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows(1).Shading.BackgroundPatternColor = RGB(200, 200, 200)
        On Error GoTo CleanFail
        
        ' Batch populate table data (starting from row 2)
        Dim i As Long
        i = 2  ' Start from row 2 (row 1 is header)
        For Each item In sorted
            On Error Resume Next
            tbl.Cell(i, 1).Range.Text = item("DisplayText")
            tbl.Cell(i, 2).Range.Text = item("Address")
            If Err.Number <> 0 Then
                Debug.Print "Cell write error (2): "; Err.Number; Err.Description; " Row:"; i
                Err.Clear
            End If

            ' Hyperlink add may fail for malformed addresses; leave it non-fatal
            On Error Resume Next
            wsNewDoc.Hyperlinks.Add tbl.Cell(i, 2).Range, Address:=item("Address"), TextToDisplay:=item("Address")
            If Err.Number <> 0 Then
                Debug.Print "Hyperlink add failed: "; Err.Number; Err.Description; " Addr:"; item("Address")
                Err.Clear
            End If
            On Error GoTo CleanFail

            On Error Resume Next
            tbl.Cell(i, 3).Range.Text = item("Context")
            tbl.Cell(i, 4).Range.Text = item("StatusText")
            If Err.Number <> 0 Then
                Debug.Print "Cell write error (3/4): "; Err.Number; Err.Description; " Row:"; i
                Err.Clear
            End If
            On Error GoTo CleanFail

            ' Shading can fail on some environments; make non-fatal
            On Error Resume Next
            Select Case item("ColorCode")
                Case "G": tbl.Rows(i).Shading.BackgroundPatternColor = RGB(200, 255, 200)
                Case "R": tbl.Rows(i).Shading.BackgroundPatternColor = RGB(255, 153, 153)
                Case "Y": tbl.Rows(i).Shading.BackgroundPatternColor = RGB(255, 255, 102)
            End Select
            If Err.Number <> 0 Then
                Debug.Print "Row shading failed: "; Err.Number; Err.Description; " Row:"; i
                Err.Clear
            End If
            On Error GoTo CleanFail

            i = i + 1
        Next item
    End If

Finalise:
CleanExit:
    ' Restore Word settings
    Application.ScreenUpdating = origScreenUpdating
    Application.DisplayAlerts = origDisplayAlerts
    With Application.Options
        .CheckSpellingAsYouType = origSpellCheck
        .CheckGrammarAsYouType = origGrammarCheck
    End With
    
    If SHOW_STATUSBAR Then
        Application.StatusBar = "Done: " & results.Count & " links, " & uniqueCount & " unique, " & mCacheHits & " cache hits."
    End If
    
    Exit Sub

CleanFail:
    lastErrNum = Err.Number
    lastErrDesc = Err.Description
    Debug.Print "ValidateHyperlinks aborted: "; lastErrNum; lastErrDesc
    If SHOW_STATUSBAR Then Application.StatusBar = "Validation aborted: " & lastErrNum & " " & lastErrDesc
    
    ' Ensure a document exists so user sees something
    On Error Resume Next
    If Documents.Count = 0 Then Set wsNewDoc = Documents.Add()
    If Not wsNewDoc Is Nothing Then
        With wsNewDoc.Range
            .InsertAfter "Validation aborted with error: " & lastErrNum & " " & lastErrDesc & vbCrLf & vbCrLf
            .InsertAfter "Partial results may be available." & vbCrLf
        End With
    End If
    Resume CleanExit
End Sub

' Check a unique URL and return result dictionary
Private Function CheckUniqueURL(strURL As String, ByVal processed As Long, ByVal total As Long) As Object
    Dim statusText As String, isValid As Boolean, isInvalidSymbol As Boolean
    Dim dict As Object
    
    On Error GoTo CheckErr
    
    ' Use cached check if available
    isValid = CheckURL_Cached(strURL, isInvalidSymbol, statusText)
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "StatusText", statusText
    dict.Add "IsValid", isValid
    dict.Add "IsInvalidSymbol", isInvalidSymbol
    
    If isValid Then
        dict.Add "ColorCode", "G"
    ElseIf isInvalidSymbol Then
        dict.Add "ColorCode", "R"
    Else
        dict.Add "ColorCode", "Y"
    End If
    
    Set CheckUniqueURL = dict
    Exit Function
    
CheckErr:
    Debug.Print "CheckUniqueURL error: "; Err.Number; Err.Description; " URL:"; strURL
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "StatusText", "Check error: " & Err.Number & " " & Err.Description
    dict.Add "IsValid", False
    dict.Add "IsInvalidSymbol", False
    dict.Add "ColorCode", "Y"
    Set CheckUniqueURL = dict
End Function

' Create a result item from location info and check result
Private Function CreateResultItem(doc As Word.Document, rng As Variant, txt As String, addr As String, checkResult As Object) As Object
    On Error Resume Next
    
    Dim dict As Object
    Dim rngObj As Word.Range
    
    ' Handle Variant to Range conversion
    If IsObject(rng) Then
        Set rngObj = rng
    Else
        Set CreateResultItem = Nothing
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "DisplayText", SafeText(txt)
    dict.Add "Address", NzAddr(addr)
    dict.Add "StatusText", checkResult("StatusText")
    dict.Add "Context", GetContextSnippet(doc, rngObj)
    dict.Add "ColorCode", checkResult("ColorCode")
    
    If Err.Number <> 0 Then
        Debug.Print "CreateResultItem error: "; Err.Number; Err.Description
        Set CreateResultItem = Nothing
    Else
        Set CreateResultItem = dict
    End If
End Function

' Normalize URL: trim, replace spaces and control chars with percent-escapes, ensure scheme present for HTTP/HTTPS
Private Function NormalizeURL(s As String) As String
    Dim t As String, i As Long, ch As Long
    Dim chars() As String
    Dim validCharCount As Long
    
    t = Trim$(s)
    If Len(t) = 0 Then
        NormalizeURL = ""
        Exit Function
    End If

    ' If it looks like a protocol-relative or missing-scheme web link, force http
    If LCase$(Left$(t, 7)) <> "http://" And LCase$(Left$(t, 8)) <> "https://" And LCase$(Left$(t, 7)) <> "mailto:" Then
        ' don't automatically force a scheme for other protocols; only add http when it looks like a hostname or starts with www.
        If Left$(t, 4) = "www." Or InStr(t, ".") > 0 Then
            t = "http://" & t
        End If
    End If

    ' Optimized: use array instead of string concatenation
    ReDim chars(1 To Len(t) * 3) ' Worst case: every char becomes %XX
    validCharCount = 0
    
    For i = 1 To Len(t)
        ch = Asc(Mid$(t, i, 1))
        If ch < 32 Or ch = 32 Then
            validCharCount = validCharCount + 1
            chars(validCharCount) = "%" & Right$("0" & Hex(ch), 2)
        Else
            validCharCount = validCharCount + 1
            chars(validCharCount) = Mid$(t, i, 1)
        End If
    Next i

    ReDim Preserve chars(1 To validCharCount)
    NormalizeURL = Join(chars, "")
End Function

' Cached URL check
Private Function CheckURL_Cached(strURL As String, ByRef isInvalidSymbol As Boolean, ByRef statusText As String) As Boolean
    Dim key As String, arr As Variant, isValid As Boolean
    key = Trim$(strURL)

    If mCache.Exists(key) Then
        arr = mCache(key)
        isValid = arr(0)
        isInvalidSymbol = arr(1)
        statusText = arr(2)
        mCacheHits = mCacheHits + 1
        If DEBUG_LOG Then Debug.Print "(cached)"; key; "->"; statusText
        CheckURL_Cached = isValid
        Exit Function
    End If

    isValid = CheckURL(strURL, isInvalidSymbol, statusText)
    On Error Resume Next
    mCache.Add key, Array(isValid, isInvalidSymbol, statusText)
    If Err.Number <> 0 Then
        Debug.Print "Cache add failed: "; Err.Number; Err.Description; " Key len:"; Len(key)
        Err.Clear
    End If
    On Error GoTo 0
    CheckURL_Cached = isValid
End Function

' Live URL check with HEAD optimization and UNdocs bad-link detection
Private Function CheckURL(strURL As String, ByRef isInvalidSymbol As Boolean, ByRef statusText As String) As Boolean
    Dim http As Object, code As Long, body As String, okStatus As Boolean
    Dim useHead As Boolean

    On Error GoTo NetErr

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts RESOLVE_TIMEOUT, CONNECT_TIMEOUT, SEND_TIMEOUT, RECEIVE_TIMEOUT

    On Error Resume Next
    http.Option(6) = True ' follow redirects
    On Error GoTo NetErr

    ' Try HEAD first for performance (if enabled)
    useHead = mUseHeadFirst
    
TryRequest:
    If useHead Then
        http.Open "HEAD", strURL, False
    Else
        http.Open "GET", strURL, False
    End If
    
    http.SetRequestHeader "User-Agent", "Mozilla/5.0 WordVBA"
    http.Send

    code = 0
    On Error Resume Next
    code = http.Status
    If Err.Number <> 0 Then
        Debug.Print "http.Status read failed: "; Err.Number; Err.Description
        Err.Clear
    End If
    On Error GoTo NetErr

    ' Accept 2xx and 3xx as valid (redirects are already followed)
    okStatus = (code >= 200 And code < 400)

    If okStatus Then
        ' If we used HEAD and got success, we're done
        If useHead Then
            statusText = "HTTP " & CStr(code) & " (HEAD)"
            CheckURL = True
        Else
            ' For GET, check for UNdocs bad-link wrapper
            body = http.ResponseText
            Dim plainBody As String
            plainBody = LCase$(body)

            ' Detect UNdocs bad-link wrapper
            If InStr(plainBody, "<iframe") > 0 And _
               InStr(plainBody, "/api/symbol/access?s=") > 0 And _
               InStr(plainBody, "<title>document viewer</title>") > 0 Then

                isInvalidSymbol = True
                statusText = "200 OK, but UNdocs bad-link wrapper detected"
                CheckURL = False
            Else
                statusText = "HTTP " & CStr(code)
                CheckURL = True
            End If
        End If
    ElseIf useHead And code = 405 Then
        ' Method Not Allowed - server doesn't support HEAD, retry with GET
        useHead = False
        GoTo TryRequest
    Else
        statusText = "HTTP " & CStr(code) & " - " & http.statusText
        isInvalidSymbol = False
        CheckURL = False
    End If

Clean:
    Set http = Nothing
    Exit Function

NetErr:
    isInvalidSymbol = False
    statusText = "Request error/timeout: " & Err.Number & " " & Err.Description
    Debug.Print "NetErr for URL: "; strURL; " Err: "; Err.Number; Err.Description
    CheckURL = False
    Resume Clean
End Function

' Helpers
Private Function NzAddr(s As String) As String
    NzAddr = IIf(Len(s) = 0, "", Trim$(s))
End Function

Private Function SafeText(s As String) As String
    SafeText = IIf(Len(s) = 0, "(no text)", s)
End Function

Private Function GetContextSnippet(doc As Word.Document, rng As Word.Range) As String
    On Error Resume Next
    Dim startPos As Long, endPos As Long, snippet As String
    
    ' Default empty string
    GetContextSnippet = ""
    
    If rng Is Nothing Then Exit Function
    
    startPos = rng.Start - CONTEXT_CHARS
    If startPos < 0 Then startPos = 0
    endPos = rng.End + CONTEXT_CHARS
    If endPos > doc.Range.End Then endPos = doc.Range.End
    
    snippet = doc.Range(startPos, endPos).Text
    If Err.Number = 0 Then
        GetContextSnippet = Replace(Replace(snippet, vbCr, " "), vbLf, " ")
    End If
    
    On Error GoTo 0
End Function
