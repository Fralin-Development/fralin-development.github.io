' ==============================================================================
' 🏛️ FRALIN MUSEUM OF ART - LIVE DONOR RECOGNITION ROLL FOR POWERPOINT
' ==============================================================================
' Description: Fetches live donor data (donors.csv) from GitHub Pages and
'              generates a native PowerPoint text box with smooth, movie-style
'              upward rolling "Credits" animation.
'
' Usage:
' 1. Open PowerPoint and press Alt + F11 (Windows) or Option + F11 (Mac).
' 2. Click Insert > Module.
' 3. Copy and paste this ENTIRE code into the module window.
' 4. Run the macro "SyncDonorsAndCreateRollingList".
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' CONFIGURATION CONSTANTS
' ------------------------------------------------------------------------------
Public Const CSV_URL As String = "https://fralin-development.github.io/donors.csv"
Public Const HEADER_TITLE As String = "THANK YOU TO OUR GENEROUS DONORS"
Public Const SUBHEADER_TITLE As String = "Who Make Art Together Possible"
Public Const FONT_FAMILY As String = "Poppins"
Public Const SCROLL_SPEED_POINTS_PER_SEC As Single = 60  ' Scroll speed in points/sec (60 = smooth & readable)
Public Const ANIM_REPEAT_COUNT As Long = 1000            ' Infinite / continuous looping
Public Const TARGET_SLIDE_INDEX As Long = 1

' ==============================================================================
' 1. MAIN MACRO: Sync Donors and Generate Continuous Rolling Credits
' ==============================================================================
Sub SyncDonorsAndCreateRollingList()
    Dim csvContent As String
    Dim donorRows As Variant
    Dim donorSlide As Slide
    Dim donorBox As Shape
    Dim animEffect As Effect
    Dim i As Long
    Dim slideW As Single, slideH As Single
    Dim boxW As Single
    Dim fullText As String
    Dim dCount As Long
    Dim dName As String, dAmount As String, dMsg As String
    Dim totalTravelDistance As Single
    Dim dynamicDuration As Single

    On Error GoTo ErrorHandler

    ' 1. Fetch CSV Content from Live URL (cache-busted with timestamp)
    csvContent = FetchUrlContent(CSV_URL & "?t=" & Format(Now, "yyyymmddhhnnss"))
    If Len(Trim(csvContent)) = 0 Then
        MsgBox "Failed to download donor data from:" & vbCrLf & CSV_URL & vbCrLf & vbCrLf & _
               "Please check your internet connection.", vbCritical, "Donor Sync Error"
        Exit Sub
    End If

    ' 2. Parse CSV into 2D Array
    donorRows = ParseCsvString(csvContent)
    
    If Not IsArray(donorRows) Then
        MsgBox "Failed to parse donor CSV.", vbCritical, "Parser Error"
        Exit Sub
    End If
    
    dCount = UBound(donorRows, 1) - LBound(donorRows, 1) + 1
    If dCount <= 1 Then
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

    ' 5. Calculate Slide Dimensions
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    boxW = slideW * 0.85

    ' 6. Build Formatted Text Block
    fullText = HEADER_TITLE & vbCrLf
    If Len(Trim(SUBHEADER_TITLE)) > 0 Then
        fullText = fullText & SUBHEADER_TITLE & vbCrLf
    End If
    fullText = fullText & String(30, "-") & vbCrLf & vbCrLf

    ' Skip row 0 (Header: Name, Amount, Message)
    For i = 1 To UBound(donorRows, 1)
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

    ' Trailing empty buffer: gives a clean 3-second blank pause after the last donor
    fullText = fullText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

    ' 7. Create Text Box placed at top=0 so Credits starts immediately at slide bottom
    Set donorBox = donorSlide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        (slideW - boxW) / 2, _
        0, _
        boxW, _
        slideH)

    donorBox.Name = "FralinLiveDonorRoll"
    donorBox.TextFrame.WordWrap = msoTrue
    donorBox.TextFrame.MarginLeft = 20
    donorBox.TextFrame.MarginRight = 20
    donorBox.TextFrame.MarginTop = 20
    donorBox.TextFrame.MarginBottom = 20
    donorBox.TextFrame.TextRange.Text = fullText

    ' Typography & Fralin Blue (#127CC2)
    With donorBox.TextFrame.TextRange
        .Font.Name = FONT_FAMILY
        .Font.Size = 22
        .Font.Bold = msoFalse
        .Font.Color.RGB = RGB(18, 124, 194)
        .ParagraphFormat.Alignment = ppAlignCenter
    End With

    ' Large bold title
    On Error Resume Next
    With donorBox.TextFrame.TextRange.Paragraphs(1)
        .Font.Bold = msoTrue
        .Font.Size = 28
    End With
    On Error GoTo ErrorHandler

    ' Auto-fit height to all text content
    donorBox.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    donorBox.Width = boxW
    donorBox.Left = (slideW - boxW) / 2
    donorBox.Top = 0

    ' 8. Calculate Dynamic Duration based on TRUE Content Height
    totalTravelDistance = slideH + donorBox.Height
    dynamicDuration = totalTravelDistance / SCROLL_SPEED_POINTS_PER_SEC
    If dynamicDuration < 20 Then dynamicDuration = 20

    ' 9. Apply Native PowerPoint "Credits" Animation
    For i = donorSlide.TimeLine.MainSequence.Count To 1 Step -1
        donorSlide.TimeLine.MainSequence(i).Delete
    Next i

    Set animEffect = donorSlide.TimeLine.MainSequence.AddEffect( _
        Shape:=donorBox, _
        EffectId:=msoAnimEffectCredits, _
        Trigger:=msoAnimTriggerWithPrevious)

    animEffect.Timing.Duration = dynamicDuration
    animEffect.Timing.RepeatCount = ANIM_REPEAT_COUNT

    MsgBox "Success! Loaded " & (UBound(donorRows, 1)) & " donors onto Slide " & TARGET_SLIDE_INDEX & "." & vbCrLf & vbCrLf & _
           "• Animation Duration: " & Round(dynamicDuration) & " seconds" & vbCrLf & _
           "• Starting Position: Immediate entrance from bottom" & vbCrLf & vbCrLf & _
           "Press F5 (or Shift + F5) to start the presentation!", vbInformation, "Donor Roll Updated"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during sync:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "VBA Sync Error"
End Sub

' ==============================================================================
' 2. DARK THEME GALA VARIANT
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
        donorSlide.FollowMasterBackground = msoFalse
        donorSlide.Background.Fill.Solid
        donorSlide.Background.Fill.ForeColor.RGB = RGB(7, 10, 18)
        
        donorBox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        donorBox.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RGB(229, 193, 88)
    End If
End Sub

' ==============================================================================
' 3. HTTP HELPER: Cross-platform HTTP GET (Windows & Mac)
' ==============================================================================
Private Function FetchUrlContent(ByVal url As String) As String
    Dim responseText As String
    
    #If Mac Then
        ' Mac OS: Execute curl via AppleScript with SSL certificate bypass (-k)
        On Error Resume Next
        responseText = MacScript("do shell script ""curl -s -k -L '" & url & "'""")
        On Error GoTo 0
    #Else
        Dim http As Object
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
' 4. CSV PARSER
' ==============================================================================
Private Function ParseCsvString(ByVal csvRaw As String) As Variant
    Dim lines() As String
    Dim totalLines As Long
    Dim i As Long, j As Long
    Dim results() As Variant
    Dim rowCount As Long
    Dim cleanText As String
    Dim lineStr As String
    Dim tokens As Collection
    Dim finalResults() As Variant
    
    cleanText = Replace(csvRaw, vbCrLf, vbLf)
    cleanText = Replace(cleanText, vbCr, vbLf)
    lines = Split(cleanText, vbLf)
    totalLines = UBound(lines) - LBound(lines) + 1
    
    ReDim results(0 To totalLines, 0 To 2)
    rowCount = 0
    
    For i = LBound(lines) To UBound(lines)
        lineStr = Trim(lines(i))
        If Len(lineStr) > 0 Then
            Set tokens = TokenizeCsvLine(lineStr)
            
            If tokens.Count >= 1 Then
                results(rowCount, 0) = CStr(tokens(1)) ' Name
                If tokens.Count >= 2 Then
                    results(rowCount, 1) = CStr(tokens(2)) ' Amount
                Else
                    results(rowCount, 1) = ""
                End If
                If tokens.Count >= 3 Then
                    results(rowCount, 2) = CStr(tokens(3)) ' Message
                Else
                    results(rowCount, 2) = ""
                End If
                rowCount = rowCount + 1
            End If
        End If
    Next i
    
    If rowCount > 0 Then
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
