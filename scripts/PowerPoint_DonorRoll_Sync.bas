' ==============================================================================
' 🏛️ FRALIN MUSEUM OF ART - LIVE DONOR RECOGNITION ROLL FOR POWERPOINT
' ==============================================================================
' Description: Fetches live donor data (donors.csv) from GitHub Pages and
'              generates a native PowerPoint text box with smooth, continuous
'              upward Motion Path animation across ALL donor records.
'              Supports live mid-presentation auto-updates & background layering!
'
' Usage:
' 1. Open PowerPoint and press Alt + F11 (Windows) or Option + F11 (Mac).
' 2. Click Insert > Module.
' 3. Copy and paste this ENTIRE code into the module window.
' 4. Save your presentation as PowerPoint Macro-Enabled Presentation (.pptm).
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' CONFIGURATION CONSTANTS
' ------------------------------------------------------------------------------
Public Const CSV_URL As String = "https://fralin-development.github.io/donors.csv"
Public Const FONT_FAMILY As String = "Poppins"
Public Const FONT_SIZE As Single = 22
Public Const SCROLL_SPEED_POINTS_PER_SEC As Single = 120  ' Speed: 60 = relaxed, 120 = standard, 200 = brisk
Public Const ANIM_REPEAT_COUNT As Long = 1000            ' Infinite / continuous looping
Public Const TARGET_SLIDE_INDEX As Long = 1

' ==============================================================================
' 1. MAIN MACROS: Manual & Mid-Presentation Silent Auto-Sync
' ==============================================================================

' Manual sync (shows confirmation popup)
Sub SyncDonorsAndCreateRollingList()
    ExecuteDonorSync True
End Sub

' Silent sync for live presentation (no popup interruption)
Sub SyncDonorsSilently()
    ExecuteDonorSync False
End Sub

' Gala Dark Mode
Sub SyncDonorsDarkMode()
    ExecuteDonorSync True
    ApplyDarkModeStyling
End Sub

' ==============================================================================
' 2. MID-PRESENTATION AUTO-UPDATE HOOKS
' ==============================================================================

' Auto-syncs live donors the instant the presentation begins (F5)
Sub OnSlideShowBegin(ByVal Wn As SlideShowWindow)
    On Error Resume Next
    ExecuteDonorSync False
End Sub

' Auto-syncs live donors every time the presentation loops back to Slide 1
Sub OnSlideShowPageChange(ByVal Wn As SlideShowWindow)
    On Error Resume Next
    If Wn.View.Slide.SlideIndex = TARGET_SLIDE_INDEX Then
        ExecuteDonorSync False
    End If
End Sub

' ==============================================================================
' 3. CORE SYNC ENGINE
' ==============================================================================
Private Sub ExecuteDonorSync(ByVal showPopups As Boolean)
    Dim csvContent As String
    Dim donorRows As Variant
    Dim donorSlide As Slide
    Dim donorBox As Shape
    Dim animEffect As Effect
    Dim b As AnimationBehavior
    Dim i As Long
    Dim slideW As Single, slideH As Single
    Dim boxW As Single
    Dim fullText As String
    Dim dCount As Long
    Dim dName As String, dAmount As String, dMsg As String
    Dim totalTravelDistance As Single
    Dim dynamicDuration As Single
    Dim calculatedH As Single
    Dim paraCount As Long
    Dim relTravel As Single

    On Error GoTo ErrorHandler

    ' 1. Fetch CSV Content from Live URL or Local File
    csvContent = FetchUrlContent(CSV_URL & "?t=" & Format(Now, "yyyymmddhhnnss"))
    If Len(Trim(csvContent)) = 0 Then
        If showPopups Then
            MsgBox "No donor data could be loaded. Please check your internet connection.", vbCritical, "Donor Sync Error"
        End If
        Exit Sub
    End If

    ' 2. Parse CSV into 2D Array
    donorRows = ParseCsvString(csvContent)
    If Not IsArray(donorRows) Then Exit Sub
    
    dCount = UBound(donorRows, 1) - LBound(donorRows, 1) + 1
    If dCount <= 1 Then Exit Sub

    ' 3. Validate Slide Selection
    If TARGET_SLIDE_INDEX > ActivePresentation.Slides.Count Or TARGET_SLIDE_INDEX < 1 Then Exit Sub
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

    ' 6. Build Formatted Text Block (Pure donor names)
    fullText = ""

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

    ' Trailing empty buffer: clean blank pause after the last donor before seamless loop
    fullText = fullText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

    ' 7. Create Text Box placed right below the slide bottom (Top = slideH)
    Set donorBox = donorSlide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        (slideW - boxW) / 2, _
        slideH, _
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
        .Font.Size = FONT_SIZE
        .Font.Bold = msoFalse
        .Font.Color.RGB = RGB(18, 124, 194)
        .ParagraphFormat.Alignment = ppAlignCenter
    End With

    ' Auto-fit height to calculate exact rendered height
    donorBox.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    DoEvents

    ' Explicitly calculate full text height
    On Error Resume Next
    calculatedH = donorBox.TextFrame.TextRange.BoundHeight
    paraCount = donorBox.TextFrame.TextRange.Paragraphs.Count
    On Error GoTo ErrorHandler

    If calculatedH < (paraCount * FONT_SIZE * 1.5) Then
        calculatedH = (paraCount * FONT_SIZE * 1.6) + 200
    End If

    ' Set full height and starting position below bottom of screen
    donorBox.TextFrame.AutoSize = ppAutoSizeNone
    donorBox.Width = boxW
    donorBox.Height = calculatedH
    donorBox.Left = (slideW - boxW) / 2
    donorBox.Top = slideH

    ' 8. Z-Order: Send donor box to back so it sneaks under any pictures or logos
    donorBox.ZOrder msoSendToBack
    For i = 1 To donorSlide.Shapes.Count
        If donorSlide.Shapes(i).Name <> "FralinLiveDonorRoll" Then
            donorSlide.Shapes(i).ZOrder msoBringToFront
        End If
    Next i

    ' 9. Calculate Full Travel Distance and Dynamic Duration
    totalTravelDistance = slideH + calculatedH
    dynamicDuration = totalTravelDistance / SCROLL_SPEED_POINTS_PER_SEC
    If dynamicDuration < 5 Then dynamicDuration = 5

    ' 10. Apply Native PowerPoint Upward Motion Path
    For i = donorSlide.TimeLine.MainSequence.Count To 1 Step -1
        donorSlide.TimeLine.MainSequence(i).Delete
    Next i

    Set animEffect = donorSlide.TimeLine.MainSequence.AddEffect( _
        Shape:=donorBox, _
        EffectId:=msoAnimEffectPathUp, _
        Trigger:=msoAnimTriggerWithPrevious)

    animEffect.Timing.Duration = dynamicDuration
    animEffect.Timing.RepeatCount = ANIM_REPEAT_COUNT
    animEffect.Timing.SmoothStart = msoFalse
    animEffect.Timing.SmoothEnd = msoFalse

    ' Set travel distance using BOTH Mac ByY/ToY properties and VML Path
    relTravel = totalTravelDistance / slideH
    
    For Each b In animEffect.Behaviors
        If b.Type = msoAnimTypeMotion Then
            With b.MotionEffect
                .FromX = 0
                .FromY = 0
                .ToX = 0
                .ToY = -relTravel
                .ByX = 0
                .ByY = -relTravel
                On Error Resume Next
                .Path = "M 0 0 L 0 -" & Replace(Format(relTravel, "0.0000"), ",", ".")
                On Error GoTo ErrorHandler
            End With
        End If
    Next b

    If showPopups Then
        MsgBox "Success! Loaded " & (UBound(donorRows, 1)) & " donors onto Slide " & TARGET_SLIDE_INDEX & "." & vbCrLf & vbCrLf & _
               "• Total Donors: " & UBound(donorRows, 1) & vbCrLf & _
               "• Full Height: " & Round(calculatedH) & " points" & vbCrLf & _
               "• Layering: Set behind pictures / foreground elements" & vbCrLf & _
               "• Animation Duration: " & Round(dynamicDuration, 1) & " seconds" & vbCrLf & vbCrLf & _
               "Press F5 (or Shift + F5) to start the presentation!", vbInformation, "Donor Roll Updated"
    End If
    Exit Sub

ErrorHandler:
    If showPopups Then
        MsgBox "An error occurred during sync:" & vbCrLf & _
               "Error " & Err.Number & ": " & Err.Description, vbCritical, "VBA Sync Error"
    End If
End Sub

Private Sub ApplyDarkModeStyling()
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
        
        ' Crisp gold donor names on dark background
        donorBox.TextFrame.TextRange.Font.Color.RGB = RGB(229, 193, 88)
    End If
End Sub

' ==============================================================================
' 4. HTTP & FILE HELPER: Robust Cross-Platform Loading
' ==============================================================================
Private Function FetchUrlContent(ByVal url As String) As String
    Dim responseText As String
    
    #If Mac Then
        ' 1. Attempt Mac download via AppleScript
        On Error Resume Next
        responseText = MacScript("do shell script ""curl -s -k -L 'https://fralin-development.github.io/donors.csv'""")
        On Error GoTo 0
        
        ' 2. If blocked by macOS sandbox, auto-check local paths
        If Len(Trim(responseText)) = 0 Then
            ' Check same folder as presentation
            If Len(ActivePresentation.Path) > 0 Then
                responseText = ReadLocalFileText(ActivePresentation.Path & "/donors.csv")
            End If
            
            ' Check Downloads folder
            If Len(Trim(responseText)) = 0 Then
                Dim dlFolder As String
                On Error Resume Next
                dlFolder = MacScript("POSIX path of (path to downloads folder)")
                On Error GoTo 0
                If Len(dlFolder) > 0 Then
                    responseText = ReadLocalFileText(dlFolder & "donors.csv")
                End If
            End If
            
            ' 3. If still not found, prompt user with native Mac file picker
            If Len(Trim(responseText)) = 0 Then
                Dim chosenPath As String
                On Error Resume Next
                chosenPath = MacScript("try" & vbCr & "return POSIX path of (choose file with prompt ""Select donors.csv file:"")" & vbCr & "on error" & vbCr & "return """"" & vbCr & "end try")
                On Error GoTo 0
                If Len(chosenPath) > 0 Then
                    responseText = ReadLocalFileText(chosenPath)
                End If
            End If
        End If
    #Else
        ' Windows: Standard XMLHTTP request
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
        
        ' Windows Local Fallback
        If Len(Trim(responseText)) = 0 And Len(ActivePresentation.Path) > 0 Then
            responseText = ReadLocalFileText(ActivePresentation.Path & "\donors.csv")
        End If
    #End If
    
    FetchUrlContent = responseText
End Function

Private Function ReadLocalFileText(ByVal filePath As String) As String
    Dim fNum As Integer
    Dim fileText As String
    On Error Resume Next
    If Len(filePath) > 0 And Dir(filePath) <> "" Then
        fNum = FreeFile
        Open filePath For Binary Access Read As #fNum
        If LOF(fNum) > 0 Then
            fileText = Space$(LOF(fNum))
            Get #fNum, , fileText
        End If
        Close #fNum
    End If
    On Error GoTo 0
    ReadLocalFileText = fileText
End Function

' ==============================================================================
' 5. CSV PARSER
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
