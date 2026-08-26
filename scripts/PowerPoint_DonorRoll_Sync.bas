Attribute VB_Name = "PowerPoint_DonorRoll_Sync"
' ==============================================================================
' 🏛️ FRALIN MUSEUM OF ART - LIVE DONOR RECOGNITION ROLL FOR POWERPOINT
' ==============================================================================
' Description: Fetches live donor data (donors.csv) from GitHub Pages and
'              generates a native PowerPoint text box with smooth, movie-style
'              upward rolling "Credits" animation.
'
' Usage:
' 1. Open PowerPoint and press Alt + F11 (Windows) or Option + F11 (Mac).
' 2. File > Import File > Select this .bas file (or paste into a new Module).
' 3. Run the macro "SyncDonorsAndCreateRollingList".
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' CONFIGURATION CONSTANTS
' ------------------------------------------------------------------------------
Public Const CSV_URL As String = "https://fralin-development.github.io/donors.csv"
Public Const HEADER_TITLE As String = "THANK YOU TO OUR GENEROUS DONORS"
Public Const SUBHEADER_TITLE As String = "Who Make Art Together Possible"
Public Const FONT_FAMILY As String = "Poppins"
Public Const SCROLL_DURATION_SECONDS As Single = 45#
Public Const ANIM_REPEAT_COUNT As Long = 10
Public Const TARGET_SLIDE_INDEX As Long = 1

' ==============================================================================
' MAIN MACRO: Sync Donors and Generate Native Rolling Credits
' ==============================================================================
Sub SyncDonorsAndCreateRollingList()
    Dim csvContent As String
    Dim donorRows() As Variant
    Dim donorSlide As Slide
    Dim donorBox As Shape
    Dim animEffect As Effect
    Dim i As Long
    Dim slideW As Single, slideH As Single
    Dim boxW As Single, boxH As Single
    Dim fullText As String
    
    On Error GoTo ErrorHandler

    ' 1. Fetch CSV Content from GitHub Pages
    csvContent = FetchUrlContent(CSV_URL & "?t=" & Format(Now, "yyyymmddhhnnss"))
    If Len(Trim(csvContent)) = 0 Then
        MsgBox "Failed to download donor data from:" & vbCrLf & CSV_URL, vbCritical, "Donor Sync Error"
        Exit Sub
    End If

    ' 2. Parse CSV
    donorRows = ParseCsvString(csvContent)
    If (UBound(donorRows, 1) - LBound(donorRows, 1)) < 1 Then
        MsgBox "The downloaded donor file contains no donor records.", vbExclamation, "No Donors Found"
        Exit Sub
    End If

    ' 3. Validate Slide Selection
    If TARGET_SLIDE_INDEX > ActivePresentation.Slides.Count Or TARGET_SLIDE_INDEX < 1 Then
        MsgBox "Slide " & TARGET_SLIDE_INDEX & " does not exist in the active presentation.", vbCritical, "Slide Error"
        Exit Sub
    End If
    Set donorSlide = ActivePresentation.Slides(TARGET_SLIDE_INDEX)

    ' 4. Clean up any previous donor boxes on this slide
    For i = donorSlide.Shapes.Count To 1 Step -1
        If donorSlide.Shapes(i).Name = "FralinLiveDonorRoll" Then
            donorSlide.Shapes(i).Delete
        End If
    Next i

    ' 5. Calculate Slide Geometry & Dimensions
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    boxW = slideW * 0.85
    boxH = slideH * 0.85

    ' 6. Build Text Block
    fullText = HEADER_TITLE & vbCrLf
    If Len(Trim(SUBHEADER_TITLE)) > 0 Then
        fullText = fullText & SUBHEADER_TITLE & vbCrLf
    End If
    fullText = fullText & String(30, "-") & vbCrLf & vbCrLf

    For i = 1 To UBound(donorRows, 1)
        Dim dName As String, dAmount As String, dMsg As String
        dName = Trim(CStr(donorRows(i, 0)))
        dAmount = Trim(CStr(donorRows(i, 1)))
        dMsg = Trim(CStr(donorRows(i, 2)))

        If Len(dName) > 0 Then
            fullText = fullText & dName
            If Len(dAmount) > 0 Then
                fullText = fullText & "  —  " & dAmount
            End If
            If Len(dMsg) > 0 Then
                fullText = fullText & vbCrLf & """" & dMsg & """"
            End If
            fullText = fullText & vbCrLf & vbCrLf
        End If
    Next i

    ' 7. Create Text Box
    Set donorBox = donorSlide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        (slideW - boxW) / 2, _
        (slideH - boxH) / 2, _
        boxW, _
        boxH)

    donorBox.Name = "FralinLiveDonorRoll"
    donorBox.TextFrame.WordWrap = msoTrue
    donorBox.TextFrame.MarginLeft = 20
    donorBox.TextFrame.MarginRight = 20
    donorBox.TextFrame.MarginTop = 20
    donorBox.TextFrame.MarginBottom = 20
    donorBox.TextFrame.TextRange.Text = fullText

    ' Format Typography & Color (Fralin Blue: #127CC2 -> RGB(18, 124, 194))
    With donorBox.TextFrame.TextRange
        .Font.Name = FONT_FAMILY
        .Font.Size = 22
        .Font.Bold = msoFalse
        .Font.Color.RGB = RGB(18, 124, 194)
        .ParagraphFormat.Alignment = ppAlignCenter
    End With

    ' Make Title larger and bold
    With donorBox.TextFrame.TextRange.Paragraphs(1)
        .Font.Bold = msoTrue
        .Font.Size = 28
    End With

    ' 8. Apply Native PowerPoint "Credits" Animation
    ' Clear existing animations on the slide
    For i = donorSlide.TimeLine.MainSequence.Count To 1 Step -1
        donorSlide.TimeLine.MainSequence(i).Delete
    Next i

    Set animEffect = donorSlide.TimeLine.MainSequence.AddEffect( _
        Shape:=donorBox, _
        EffectId:=msoAnimEffectCredits, _
        Trigger:=msoAnimTriggerWithPrevious)

    animEffect.Timing.Duration = SCROLL_DURATION_SECONDS
    animEffect.Timing.RepeatCount = ANIM_REPEAT_COUNT

    MsgBox "Success! Updated " & (UBound(donorRows, 1)) & " donors on Slide " & TARGET_SLIDE_INDEX & "." & vbCrLf & _
           "Press F5 (or Shift + F5) to start the slideshow and watch the rolling credits.", _
           vbInformation, "Fralin Donor Roll Updated"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during sync:" & vbCrLf & Err.Description, vbCritical, "VBA Error " & Err.Number
End Sub

' ==============================================================================
' DARK THEME VARIANT: Navy & Gold gala theme
' ==============================================================================
Sub SyncDonorsDarkMode()
    SyncDonorsAndCreateRollingList
    
    Dim donorSlide As Slide
    Dim donorBox As Shape
    Dim i As Long
    
    On Error Resume Next
    Set donorSlide = ActivePresentation.Slides(TARGET_SLIDE_INDEX)
    For i = 1 To donorSlide.Shapes.Count
        If donorSlide.Shapes(i).Name = "FralinLiveDonorRoll" Then
            Set donorBox = donorSlide.Shapes(i)
            Exit For
        End If
    Next i
    
    If Not donorBox Is Nothing Then
        ' Dark background on slide
        donorSlide.Background.Fill.Solid
        donorSlide.Background.Fill.ForeColor.RGB = RGB(7, 10, 18)
        
        ' Gold Text (#E5C158 -> RGB(229, 193, 88))
        donorBox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        donorBox.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RGB(229, 193, 88)
    End If
End Sub

' ==============================================================================
' HTTP HELPER: Cross-platform HTTP GET (Windows & Mac)
' ==============================================================================
Private Function FetchUrlContent(ByVal url As String) As String
    Dim http As Object
    Dim responseText As String
    
    #If Mac Then
        ' macOS VBA curl execution
        Dim scriptCmd As String
        Dim tmpFile As String
        tmpFile = "/tmp/fralin_donors.csv"
        scriptCmd = "do shell script ""curl -s -L '" & url & "' -o " & tmpFile & " && cat " & tmpFile & """"
        On Error Resume Next
        responseText = MacScript(scriptCmd)
        On Error GoTo 0
    #Else
        ' Windows ServerXMLHTTP / XMLHTTP
        On Error Resume Next
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
        If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
        If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        If Not http Is Nothing Then
            http.Open "GET", url, False
            http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerPoint-VBA"
            http.setRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"
            http.setRequestHeader "Pragma", "no-cache"
            http.Send
            If http.Status = 200 Then
                responseText = http.responseText
            End If
        End If
        On Error GoTo 0
    #End If
    
    FetchUrlContent = responseText
End Function

' ==============================================================================
' CSV PARSER: Handles standard CSV, multi-line values, and quotes
' ==============================================================================
Private Function ParseCsvString(ByVal csvRaw As String) As Variant
    Dim lines() As String
    Dim totalLines As Long
    Dim i As Long, j As Long
    Dim results() As String
    Dim rowCount As Long
    Dim cleanText As String
    
    cleanText = Replace(csvRaw, vbCrLf, vbLf)
    cleanText = Replace(cleanText, vbCr, vbLf)
    lines = Split(cleanText, vbLf)
    totalLines = UBound(lines) - LBound(lines) + 1
    
    ReDim results(0 To totalLines, 0 To 2)
    rowCount = 0
    
    For i = LBound(lines) To UBound(lines)
        Dim lineStr As String
        lineStr = Trim(lines(i))
        If Len(lineStr) > 0 Then
            Dim tokens As Collection
            Set tokens = TokenizeCsvLine(lineStr)
            
            If tokens.Count >= 1 Then
                results(rowCount, 0) = tokens(1) ' Name
                If tokens.Count >= 2 Then results(rowCount, 1) = tokens(2) ' Amount
                If tokens.Count >= 3 Then results(rowCount, 2) = tokens(3) ' Message
                rowCount = rowCount + 1
            End If
        End If
    Next i
    
    ' Resize array to actual non-empty row count
    If rowCount > 0 Then
        Dim finalResults() As String
        ReDim finalResults(0 To rowCount - 1, 0 To 2)
        For i = 0 To rowCount - 1
            For j = 0 To 2
                finalResults(i, j) = results(i, j)
            Next j
        Next i
        ParseCsvString = finalResults
    Else
        ParseCsvString = results
    End If
End Function

Private Function TokenizeCsvLine(ByVal line As String) As Collection
    Dim col As New Collection
    Dim curToken As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String, nextCh As String
    
    inQuotes = False
    curToken = ""
    
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If i < Len(line) Then nextCh = Mid(line, i + 1, 1) Else nextCh = ""
        
        If ch = """" Then
            If inQuotes And nextCh = """" Then
                curToken = curToken & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            col.Add Trim(curToken)
            curToken = ""
        Else
            curToken = curToken & ch
        End If
    Next i
    
    col.Add Trim(curToken)
    Set TokenizeCsvLine = col
End Function
