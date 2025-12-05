Attribute VB_Name = "FolderScan_M"
Option Explicit

Private visited As Object  ' used by RecursiveFolderScan guard
Private Const MAX_SCAN_DEPTH As Long = 40  ' safety guard against runaway recursion

Private scanSteps As Long
Private Const SCAN_MAX_STEPS As Long = 200000   ' safety cutoff

Public Sub FolderScan()
    Const DO_SORT As Boolean = True
    Const DO_HYPERLINKS As Boolean = True
    Const DO_PAINT_MISSING As Boolean = True
    Const DO_REFRESH_ALL As Boolean = True
    Const DO_UPDATE_KW As Boolean = False
    Const DO_TEST_CTRL_FILENAME As Boolean = True

    Dim ws As Worksheet
    Dim loConfig As ListObject, loFiles As ListObject
    Dim found As Range
    Dim aLocalRoot As String, domainName As String, htmlFilePath As String
    Dim urlPrefix As String, urlSuffix As String

    Dim r As ListRow, rowToUpdate As ListRow
    Dim desc As FileDescriptor

    Dim objType As String, showVal As String
    Dim cellVal As Variant
    Dim maxNum As Long, nextNum As Long

    Dim colNumber As Long, colDateFound As Long, colKeywords As Long
    Dim colDomain As Long, colCategory As Long, colFolder As Long
    Dim colObjType As Long
    Dim colFileName As Long
    Dim colRelPath As Long, colShow As Long, colLink As Long
    Dim colError As Long

    ' repositories / sets
    Dim existingRows As Object
    Dim existingRelSet As Object
    Dim existingFolderSet As Object
    Dim showSettings As Object

    Dim allowed As Object
    Dim fsRelSet As Object

    Dim toKeep As Object
    Dim toAdd As Object

    Dim baseNewSet As Object
    Dim showDefaults As Object

    Dim keysArr As Variant, kRel As Variant
    Dim parentPath As String, p As Long
    Dim rp As String

    Dim blankNumberCells As Collection
    Dim newKeys As Variant
    Dim numRows As Long, numCols As Long, i As Long
    Dim batchData() As Variant, startRow As Long, targetRange As Range, c As Range

    Dim encRel As String, computedLink As String, isShortcut As Boolean
    Dim totalNew As Long, idx As Long, thisCount As Long
    Dim updSteps As Long

    Const NEWROW_COLOR As Long = &HF7EBDD   ' RGB(221,235,247)

    ' timing
    Dim tBlock As Double, tTotal As Double
    tTotal = Timer
    Debug.Print "FolderScan - START at " & Format$(Now, "hh:nn:ss")

    On Error GoTo CleanFail
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '— 1) tbConfig —
    tBlock = Timer
    Debug.Print "ENTER 1) tbConfig at " & Format$(Now, "hh:nn:ss")

    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loConfig = ws.ListObjects("tbConfig")
        On Error GoTo 0
        If Not loConfig Is Nothing Then Exit For
    Next ws
    If loConfig Is Nothing Then MsgBox "Configuration table 'tbConfig' not found.", vbCritical: GoTo CleanFail

    Set found = loConfig.ListColumns("Key").DataBodyRange.Find("Local Root", , xlValues, xlWhole)
    If found Is Nothing Then MsgBox "Key = 'Local Root' not found in tbConfig.", vbCritical: GoTo CleanFail
    aLocalRoot = found.Offset(0, loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value
    If Right$(aLocalRoot, 1) <> "\" Then aLocalRoot = aLocalRoot & "\"

    Set found = loConfig.ListColumns("Key").DataBodyRange.Find("Domain name", , xlValues, xlWhole)
    If found Is Nothing Then
        domainName = ""
    Else
        domainName = CStr(found.Offset(0, loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
    End If

    Set found = loConfig.ListColumns("Key").DataBodyRange.Find("Html Index file", , xlValues, xlWhole)
    If Not found Is Nothing Then
        htmlFilePath = CStr(found.Offset(0, loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
    End If

    Set found = loConfig.ListColumns("Key").DataBodyRange.Find("Url Prefix", , xlValues, xlWhole)
    If Not found Is Nothing Then
        urlPrefix = CStr(found.Offset(0, loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
    Else
        urlPrefix = ""
    End If
    Set found = loConfig.ListColumns("Key").DataBodyRange.Find("Url Suffix", , xlValues, xlWhole)
    If Not found Is Nothing Then
        urlSuffix = CStr(found.Offset(0, loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
    Else
        urlSuffix = ""
    End If

    Debug.Print "EXIT  1) tbConfig, secs=" & Format$(Timer - tBlock, "0.000")

    '— 2) tbFiles —
    tBlock = Timer
    Debug.Print "ENTER 2) tbFiles at " & Format$(Now, "hh:nn:ss")

    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loFiles = ws.ListObjects("tbFiles")
        On Error GoTo 0
        If Not loFiles Is Nothing Then Exit For
    Next ws
    If loFiles Is Nothing Then MsgBox "Output table 'tbFiles' not found.", vbCritical: GoTo CleanFail

    colNumber = loFiles.ListColumns("#").Index
    colDateFound = loFiles.ListColumns("Date found").Index
    colKeywords = loFiles.ListColumns("Keywords").Index
    colDomain = loFiles.ListColumns("Domain").Index
    colCategory = loFiles.ListColumns("Category").Index
    colFolder = loFiles.ListColumns("Folder").Index
    colObjType = loFiles.ListColumns("Object Type").Index
    colFileName = loFiles.ListColumns("Object name").Index
    colRelPath = loFiles.ListColumns("RelativePath").Index
    colLink = IIf(ColumnExists(loFiles, "Link"), loFiles.ListColumns("Link").Index, 0)
    colShow = IIf(ColumnExists(loFiles, "Show?"), loFiles.ListColumns("Show?").Index, 0)
    colError = IIf(ColumnExists(loFiles, "Error?"), loFiles.ListColumns("Error?").Index, 0)

    ' Clear visuals up-front
    If Not loFiles.DataBodyRange Is Nothing Then
        With loFiles.DataBodyRange
            .FormatConditions.Delete
            .Interior.Pattern = xlSolid
            .Interior.TintAndShade = 0
            .Interior.ColorIndex = xlColorIndexNone
        End With
    End If

    Debug.Print "EXIT  2) tbFiles, secs=" & Format$(Timer - tBlock, "0.000")

    '— 3) Snapshot Excel repository —
    tBlock = Timer
    Debug.Print "ENTER 3) Snapshot Excel repo at " & Format$(Now, "hh:nn:ss")

    Dim relKey As String
    Set existingRows = CreateObject("Scripting.Dictionary")
    Set existingRelSet = CreateObject("Scripting.Dictionary")
    Set existingFolderSet = CreateObject("Scripting.Dictionary")
    Set showSettings = CreateObject("Scripting.Dictionary")
    Set blankNumberCells = New Collection
    maxNum = 0

    For Each r In loFiles.ListRows
        If Not IsError(r.Range.Cells(1, colRelPath).Value) Then
            ' Normalize path (kills ZERO WIDTH SPACE etc) and write it back into Excel
            relKey = NormalizePath(CStr(r.Range.Cells(1, colRelPath).Value))
            r.Range.Cells(1, colRelPath).Value = relKey

            If existingRows.Exists(relKey) Then existingRows.Remove relKey
            existingRows.Add relKey, r
            If Not existingRelSet.Exists(relKey) Then existingRelSet.Add relKey, True

            objType = CStr(r.Range.Cells(1, colObjType).Value)
            If objType = "Category" Or objType = "Folder" Or objType = "Subfolder" Then
                If Not existingFolderSet.Exists(relKey) Then existingFolderSet.Add relKey, True
            End If
        End If

        If CStr(r.Range.Cells(1, colObjType).Value) = "Category" Or _
           CStr(r.Range.Cells(1, colObjType).Value) = "Folder" Or _
           CStr(r.Range.Cells(1, colObjType).Value) = "Subfolder" Then

            If colShow > 0 Then
                showVal = LCase$(Trim$(CStr(r.Range.Cells(1, colShow).Value)))
                If Len(showVal) > 0 Then
                    ' normalize textual value and update both dictionary and cell
                    Select Case showVal
                        Case "all"
                            showVal = "all"
                            r.Range.Cells(1, colShow).Value = "All"
                        Case "nothing"
                            showVal = "nothing"
                            r.Range.Cells(1, colShow).Value = "Nothing"
                        Case "1st level", "1stlevel", "1st-level"
                            showVal = "1st level"
                            r.Range.Cells(1, colShow).Value = "1st Level"
                        Case Else
                            ' unknown value: treat as "nothing"
                            showVal = "nothing"
                            r.Range.Cells(1, colShow).Value = "Nothing"
                    End Select

                    If showSettings.Exists(relKey) Then showSettings.Remove relKey
                    showSettings.Add relKey, showVal
                End If
            End If
        End If

        ' optional test for control codes in filename on existing rows ---
        If DO_TEST_CTRL_FILENAME And colError > 0 Then
            Dim fn As String
            Dim normName As String
            fn = CStr(r.Range.Cells(1, colFileName).Value)

            If Len(fn) > 0 Then
                normName = NormalizePath(fn)
                If normName <> fn Then
                    If Len(Trim$(CStr(r.Range.Cells(1, colError).Value))) = 0 Then
                        r.Range.Cells(1, colError).Value = "Control codes in filename"
                    End If
                End If
            End If
        End If

        cellVal = r.Range.Cells(1, colNumber).Value
        If Len(Trim$(CStr(cellVal))) = 0 Then
            blankNumberCells.Add r.Range.Cells(1, colNumber)
        ElseIf IsNumeric(cellVal) Then
            If CLng(cellVal) > maxNum Then maxNum = CLng(cellVal)
        End If
    Next r
    nextNum = maxNum + 1

    Debug.Print "EXIT  3) Snapshot Excel repo, secs=" & Format$(Timer - tBlock, "0.000")

    '— 4) Scan filesystem with Show? inheritance in recursion —
    tBlock = Timer
    Debug.Print "ENTER 4) Scan filesystem at " & Format$(Now, "hh:nn:ss")

    Set allowed = CreateObject("Scripting.Dictionary")
    ScanFilesystemWithShow aLocalRoot, domainName, showSettings, allowed

    Debug.Print "EXIT  4) Scan filesystem, secs=" & Format$(Timer - tBlock, "0.000")

    '— 5) sets: Excel vs FS —
    tBlock = Timer
    Debug.Print "ENTER 5) Build sets Excel vs FS at " & Format$(Now, "hh:nn:ss")

    Set fsRelSet = CreateObject("Scripting.Dictionary")
    For Each kRel In allowed.keys
        fsRelSet.Add NormalizePath(CStr(kRel)), True
    Next kRel

    Set toKeep = CreateObject("Scripting.Dictionary")
    Set toAdd = CreateObject("Scripting.Dictionary")

    For Each kRel In existingRelSet.keys
        If fsRelSet.Exists(CStr(kRel)) Then
            toKeep.Add CStr(kRel), True
        End If
        ' If FS does not have it, we will just paint it yellow at the end
    Next kRel

    For Each kRel In fsRelSet.keys
        If Not existingRelSet.Exists(CStr(kRel)) Then
            toAdd.Add CStr(kRel), True
        End If
    Next kRel

    Debug.Print "EXIT  5) Build sets Excel vs FS, secs=" & Format$(Timer - tBlock, "0.000")

    '— 6) defaults for new base folders (Show? for new folder rows) —
    tBlock = Timer
    Debug.Print "ENTER 6) Compute Show? defaults for new folders at " & Format$(Now, "hh:nn:ss")

    Set baseNewSet = CreateObject("Scripting.Dictionary")
    Set showDefaults = CreateObject("Scripting.Dictionary")
    keysArr = allowed.keys

    ' identify new folder paths whose parent is root or an existing folder row
    For Each kRel In keysArr
        rp = CStr(kRel)
        Set desc = allowed(rp)
        If desc.ObjectType = "Category" Or desc.ObjectType = "Folder" Or desc.ObjectType = "Subfolder" Then
            If Not existingRelSet.Exists(rp) Then
                p = InStrRev(rp, "\")
                If p > 0 Then parentPath = Left$(rp, p - 1) Else parentPath = ""
                If parentPath = "" Or existingFolderSet.Exists(parentPath) Then
                    If Not baseNewSet.Exists(rp) Then baseNewSet.Add rp, True
                End If
            End If
        End If
    Next kRel

    ' For each new folder, compute parent effective rule, and set default Show?
    Dim parentEffRule As String, parentEffPath As String
    Dim parentFromExplicit As Boolean
    Dim tmpParent As String, guard2 As Long, p2 As Long

    For Each kRel In baseNewSet.keys
        rp = CStr(kRel)
        p = InStrRev(rp, "\")
        If p > 0 Then parentPath = Left$(rp, p - 1) Else parentPath = ""

        parentEffRule = ""
        parentEffPath = ""
        parentFromExplicit = False
        tmpParent = parentPath
        guard2 = 0

        Do While Len(tmpParent) > 0
            guard2 = guard2 + 1
            If guard2 > 200 Then Err.Raise vbObjectError + 912, , "Guard tripped in parent Show?-walk: " & tmpParent

            If showSettings.Exists(tmpParent) Then
                parentEffRule = CStr(showSettings(tmpParent))
                parentEffPath = tmpParent
                parentFromExplicit = True
                Exit Do
            End If

            p2 = InStrRev(tmpParent, "\")
            If p2 > 0 Then
                tmpParent = Left$(tmpParent, p2 - 1)
            Else
                Exit Do
            End If
        Loop

        If Len(parentEffRule) = 0 Then
            ' No explicit rule found above ? implicit root rule = "1st level"
            parentEffRule = "1st level"
            parentEffPath = ""
            parentFromExplicit = False
        End If

        If Not showDefaults.Exists(rp) Then
            Select Case parentEffRule
                Case "all"
                    showDefaults.Add rp, "All"

                Case "1st level"
                    If parentFromExplicit Then
                        ' explicit 1st Level ? children get "Nothing"
                        showDefaults.Add rp, "Nothing"
                    Else
                        ' implicit root 1st Level ? default "1st Level"
                        showDefaults.Add rp, "1st Level"
                    End If

                Case "nothing"
                    showDefaults.Add rp, "Nothing"

                Case Else
                    showDefaults.Add rp, "Nothing"
            End Select
        End If
    Next kRel

    Debug.Print "EXIT  6) Compute Show? defaults for new folders, secs=" & Format$(Timer - tBlock, "0.000")

    '— 7) update keep —
    tBlock = Timer
    Debug.Print "ENTER 7) Update kept rows at " & Format$(Now, "hh:nn:ss")

    updSteps = 0

    For Each kRel In toKeep.keys
        rp = CStr(kRel)

        ' Safety: skip if we somehow lack the descriptor
        If existingRows.Exists(rp) And allowed.Exists(rp) Then
            Set rowToUpdate = existingRows(rp)
            Set desc = allowed(rp)

            With rowToUpdate.Range
                .Cells(1, colDomain).Value = desc.Domain
                .Cells(1, colCategory).Value = desc.Category
                .Cells(1, colFolder).Value = desc.folder
                .Cells(1, colObjType).Value = desc.ObjectType
                .Cells(1, colFileName).Value = desc.fileName
                .Cells(1, colRelPath).Value = desc.RelativePath
                .Cells(1, colKeywords).Value = CleanKeywords(.Cells(1, colKeywords).Value)
            End With
        End If

        updSteps = updSteps + 1
        If (updSteps Mod 200) = 0 Then
            DoEvents    ' lets Esc/cancel be processed
        End If
    Next kRel

    Debug.Print "EXIT  7) Update kept rows, rows=" & updSteps & _
                " secs=" & Format$(Timer - tBlock, "0.000")

    '— 8) numbering fill —
    tBlock = Timer
    Debug.Print "ENTER 8) Numbering fill at " & Format$(Now, "hh:nn:ss")

    If blankNumberCells.Count > 0 Then
        For Each c In blankNumberCells
            If Len(Trim$(CStr(c.Value))) = 0 Then
                c.Value = nextNum
                nextNum = nextNum + 1
            End If
        Next c
    End If

    Debug.Print "EXIT  8) Numbering fill, secs=" & Format$(Timer - tBlock, "0.000")

    '— 9) add new in batch & paint blue —
    tBlock = Timer
    Debug.Print "ENTER 9) Add new rows at " & Format$(Now, "hh:nn:ss")

    If toAdd.Count > 0 Then
        newKeys = toAdd.keys
        If Not IsEmpty(newKeys) Then
            If UBound(newKeys) > LBound(newKeys) Then
                QuickSortArray newKeys, LBound(newKeys), UBound(newKeys)
            End If
        End If

        numRows = toAdd.Count
        numCols = loFiles.ListColumns.Count
        ReDim batchData(1 To numRows, 1 To numCols)

        i = 1
        For Each kRel In newKeys
            rp = CStr(kRel)
            Set desc = allowed(rp)

            '--- compute link (shortcuts vs normal files) ---
            isShortcut = EndsWithCI(desc.fileName, ".lnk") Or _
                         EndsWithCI(desc.fileName, ".url")

            If isShortcut Then
                ' Resolve ONLY for shortcuts
                computedLink = ResolveShortcutLink(aLocalRoot & desc.RelativePath)
                If Len(Trim$(computedLink)) = 0 Then
                    computedLink = "no link"
                End If
            Else
                ' Normal file: build SharePoint URL from RelativePath
                encRel = Replace(desc.RelativePath, "\", "/")
                encRel = Replace(encRel, "%", "%25")
                encRel = Replace(encRel, " ", "%20")
                encRel = Replace(encRel, "#", "%23")
                encRel = Replace(encRel, "&", "%26")
                encRel = Replace(encRel, "?", "%3F")
                encRel = Replace(encRel, "+", "%2B")
                encRel = Replace(encRel, ";", "%3B")
                encRel = Replace(encRel, ",", "%2C")
                encRel = Replace(encRel, "'", "%27")
                encRel = Replace(encRel, """", "%22")
                encRel = Replace(encRel, "<", "%3C")
                encRel = Replace(encRel, ">", "%3E")
                encRel = Replace(encRel, "(", "%28")
                encRel = Replace(encRel, ")", "%29")
                encRel = Replace(encRel, "=", "%3D")

                computedLink = urlPrefix & encRel & urlSuffix
            End If

            '--- write main data into batch array ---
            batchData(i, colDomain) = desc.Domain
            batchData(i, colCategory) = desc.Category
            batchData(i, colFolder) = desc.folder
            batchData(i, colObjType) = desc.ObjectType
            batchData(i, colFileName) = desc.fileName
            batchData(i, colRelPath) = desc.RelativePath
            If colLink > 0 Then
                batchData(i, colLink) = computedLink
            End If
            batchData(i, colNumber) = nextNum
            batchData(i, colDateFound) = Date
            batchData(i, colKeywords) = ""

            If colShow > 0 Then
                If showDefaults.Exists(rp) Then
                    batchData(i, colShow) = showDefaults(rp)
                Else
                    batchData(i, colShow) = ""
                End If
            End If

            '--- new-row filename control-code test (Error? kept with priority) ---
            If colError > 0 Then
                If Len(Trim$(CStr(batchData(i, colError)))) = 0 Then
                    Dim normNameTest As String
                    normNameTest = NormalizePath(desc.fileName)
                    If normNameTest <> desc.fileName Then
                        batchData(i, colError) = "Control codes in filename"
                    End If
                End If
            End If

            nextNum = nextNum + 1
            i = i + 1
            If (i Mod 200) = 0 Then DoEvents
        Next kRel

        ' actually add the rows to tbFiles
        For i = 1 To numRows
            loFiles.ListRows.Add
        Next i

        startRow = loFiles.DataBodyRange.Rows.Count - numRows + 1
        Set targetRange = loFiles.DataBodyRange.Rows(startRow).Resize(numRows, numCols)
        targetRange.Value = batchData
        targetRange.Interior.Color = NEWROW_COLOR

        ' clean keywords for new rows
        If colKeywords > 0 Then
            For Each c In targetRange.Columns(colKeywords).Cells
                c.Value = CleanKeywords(c.Value)
            Next c
        End If
    End If

    Debug.Print "EXIT  9) Add new rows, secs=" & Format$(Timer - tBlock, "0.000")

    '— 10) fill Show? for existing folder rows that are blank —
    tBlock = Timer
    Debug.Print "ENTER 10) Fill blank Show? in existing folders at " & Format$(Now, "hh:nn:ss")

    If colShow > 0 Then
        Dim showOrig As String
        Dim rpFolder As String
        Dim tmpPath2 As String
        Dim effRule2 As String
        Dim guard3 As Long
        Dim p3 As Long
        Dim baseRule As String

        For Each r In loFiles.ListRows
            objType = CStr(r.Range.Cells(1, colObjType).Value)
            If objType = "Category" Or objType = "Folder" Or objType = "Subfolder" Then
                showOrig = Trim$(CStr(r.Range.Cells(1, colShow).Value))
                If Len(showOrig) = 0 Then
                    rpFolder = CStr(r.Range.Cells(1, colRelPath).Value)

                    effRule2 = ""
                    tmpPath2 = rpFolder
                    guard3 = 0

                    ' Walk up to find first non-blank rule. If "1st level" is found,
                    ' this path inherits "nothing" (deeper than first level)
                    Do While Len(tmpPath2) > 0
                        guard3 = guard3 + 1
                        If guard3 > 200 Then Exit Do

                        If showSettings.Exists(tmpPath2) Then
                            baseRule = CStr(showSettings(tmpPath2))
                            If baseRule = "1st level" Then
                                effRule2 = "nothing"
                            Else
                                effRule2 = baseRule
                            End If
                            Exit Do
                        End If

                        p3 = InStrRev(tmpPath2, "\")
                        If p3 > 0 Then
                            tmpPath2 = Left$(tmpPath2, p3 - 1)
                        Else
                            Exit Do
                        End If
                    Loop

                    ' If nothing found in ancestors, implicit root rule = "1st level"
                    If Len(effRule2) = 0 Then effRule2 = "1st level"

                    Select Case effRule2
                        Case "all"
                            r.Range.Cells(1, colShow).Value = "All"
                        Case "nothing"
                            r.Range.Cells(1, colShow).Value = "Nothing"
                        Case "1st level"
                            r.Range.Cells(1, colShow).Value = "1st Level"
                    End Select
                End If
            End If
        Next r
    End If

    Debug.Print "EXIT 10) Fill blank Show?, secs=" & Format$(Timer - tBlock, "0.000")

    '— 11) hyperlinks —
    tBlock = Timer
    Debug.Print "ENTER 11) Hyperlinks at " & Format$(Now, "hh:nn:ss")

    If DO_HYPERLINKS And colLink > 0 And Not loFiles.DataBodyRange Is Nothing Then
        ApplyHyperlinksOnLinkColumnFast loFiles, colLink
    End If

    Debug.Print "EXIT 11) Hyperlinks, secs=" & Format$(Timer - tBlock, "0.000")

    '— 12) sort —
    tBlock = Timer
    Debug.Print "ENTER 12) Sort at " & Format$(Now, "hh:nn:ss")

    If DO_SORT Then
        With loFiles.Sort
            .SortFields.Clear
            .SortFields.Add key:=loFiles.ListColumns("RelativePath").DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If

    Debug.Print "EXIT 12) Sort, secs=" & Format$(Timer - tBlock, "0.000")

    '— 13) paint missing (Excel rows not in allowed FS set) —
    tBlock = Timer
    Debug.Print "ENTER 13) Paint missing (yellow) at " & Format$(Now, "hh:nn:ss")

    If DO_PAINT_MISSING Then
        Dim missRanges As New Collection
        Dim rngUnion As Range
        Dim rk As String

        For Each r In loFiles.ListRows
            rk = NormalizePath(CStr(r.Range.Cells(1, colRelPath).Value))
            If Len(rk) > 0 Then
                If Not fsRelSet.Exists(rk) Then
                    missRanges.Add r.Range
                End If
            End If
        Next r

        Set rngUnion = UnionChunked(missRanges, 60)
        If Not rngUnion Is Nothing Then rngUnion.Interior.Color = vbYellow
    End If

    Debug.Print "EXIT 13) Paint missing, secs=" & Format$(Timer - tBlock, "0.000")

    '— 14) finalize —
    tBlock = Timer
    Debug.Print "ENTER 14) Finalize at " & Format$(Now, "hh:nn:ss")

    ActiveWorkbook.Worksheets("Cover").Range("cUpdateDate").Value = Date
    If DO_UPDATE_KW Then UpdateKeywordsTable
    If DO_REFRESH_ALL Then ActiveWorkbook.RefreshAll

    MsgBox "tbFiles synchronized.", vbInformation

    Debug.Print "EXIT 14) Finalize, secs=" & Format$(Timer - tBlock, "0.000")
    Debug.Print "FolderScan - END, total secs=" & Format$(Timer - tTotal, "0.000")

CleanUp:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    If Err.Number = 18 Then
        MsgBox "Stopped by user.", vbExclamation
    Else
        MsgBox "FolderScan failed: " & Err.Number & " — " & Err.Description, vbCritical
    End If
    Debug.Print "FolderScan - ERROR " & Err.Number & " after " & Format$(Timer - tTotal, "0.000") & " secs"
    Resume CleanUp
End Sub


' --- lightweight timing ---
Public Sub Lap(ByVal label As String)
Static t0 As Double
    If t0 = 0 Then
        t0 = Timer
        'Debug.Print "? start"
    End If
    Debug.Print "Time " & Format(Timer - t0, "0.000") & "s — " & label
End Sub

' Fast, idempotent hyperlinks, with "no link" ? Error? = "Invalid shortcut link" if blank
Private Sub ApplyHyperlinksOnLinkColumnFast(ByVal lo As ListObject, ByVal colLink As Long)
    On Error GoTo EH
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If colLink <= 0 Then Exit Sub

    Dim colError As Long
    Dim hasErrorCol As Boolean
    Dim deltaCol As Long

    Dim cell As Range
    Dim url As String
    Dim errCell As Range

    ' Detect "Error?" column if present
    hasErrorCol = ColumnExists(lo, "Error?")
    If hasErrorCol Then
        colError = lo.ListColumns("Error?").Index
        deltaCol = colError - colLink   ' offset from Link column to Error? column
    Else
        deltaCol = 0
    End If

    Application.DisplayAlerts = False

    For Each cell In lo.DataBodyRange.Columns(colLink).Cells
        url = CStr(cell.Value)

        ' --- Special handling for "no link" ---
        If LCase$(Trim$(url)) = "no link" Then
            ' Ensure there is no hyperlink
            If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete

            ' If we have an Error? column and it is blank, set our message
            If hasErrorCol Then
                Set errCell = cell.Offset(0, deltaCol)
                If Len(Trim$(CStr(errCell.Value))) = 0 Then
                    errCell.Value = "Invalid shortcut link"
                End If
            End If

        ' --- Normal non-empty URLs: maintain hyperlink ---
        ElseIf Len(url) > 0 Then
            If cell.Hyperlinks.Count = 0 Or cell.Hyperlinks(1).Address <> url Then
                If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete
                cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:=url
            End If

        ' --- Blank link cell: ensure no hyperlink ---
        Else
            If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete
        End If

        If (cell.Row Mod 400) = 0 Then DoEvents
    Next cell

    Application.DisplayAlerts = True
    Exit Sub

EH:
    Application.DisplayAlerts = True
End Sub


' Chunk-safe Union to avoid O(n^2) unions (NEW)
Private Function UnionChunked(ByVal rngs As Collection, Optional ByVal chunkSize As Long = 60) As Range
    Dim i As Long, j As Long
    Dim acc As Range, part As Range
    If rngs Is Nothing Or rngs.Count = 0 Then Exit Function

    For i = 1 To rngs.Count Step chunkSize
        Set part = Nothing
        For j = i To Application.Min(i + chunkSize - 1, rngs.Count)
            If part Is Nothing Then
                Set part = rngs(j)
            Else
                Set part = Application.Union(part, rngs(j))
            End If
        Next j
        If Not part Is Nothing Then
            If acc Is Nothing Then
                Set acc = part
            Else
                Set acc = Application.Union(acc, part)
            End If
        End If
        DoEvents
    Next i
    Set UnionChunked = acc
End Function

' === helpers used here ===
Private Function ColumnExists(lo As ListObject, colName As String) As Boolean
    On Error Resume Next
    ColumnExists = Not lo.ListColumns(colName) Is Nothing
    On Error GoTo 0
End Function

'================= Helpers for Link / Shortcuts / Hyperlinks =================

Private Function EndsWithCI(ByVal text As String, ByVal suffix As String) As Boolean
    EndsWithCI = (Len(text) >= Len(suffix)) And (StrComp(Right$(text, Len(suffix)), suffix, vbTextCompare) = 0)
End Function

' Return the text inside the last "(...)" just before extension (for .lnk/.url names).
' e.g., "My Doc (pdf).lnk" ? "pdf" ; "Note (text).url" ? "text"
Private Function DeriveShortcutTypeFromName(ByVal fileName As String) As String
    Dim base As String, pOpen As Long, pClose As Long
    Dim noExt As String
    If InStrRev(fileName, ".") > 0 Then
        noExt = Left$(fileName, InStrRev(fileName, ".") - 1)
    Else
        noExt = fileName
    End If
    pClose = InStrRev(noExt, ")")
    pOpen = InStrRev(noExt, "(")
    If pOpen > 0 And pClose = Len(noExt) And pClose > pOpen Then
        DeriveShortcutTypeFromName = Trim$(Mid$(noExt, pOpen + 1, pClose - pOpen - 1))
    Else
        DeriveShortcutTypeFromName = ""
    End If
End Function

' Resolve .lnk and .url to a URL/target string; return "" if not resolvable.
Private Function ResolveShortcutLink(ByVal fullPath As String) As String
    On Error GoTo EH

    If EndsWithCI(fullPath, ".url") Then
        Dim ff As Integer, s As String
        Dim rx As Object, m As Object
        Set rx = CreateObject("VBScript.RegExp")
        With rx
            .Pattern = "^\s*URL\s*=\s*(.+)$"  ' capture everything after =
            .IgnoreCase = True
            .Global = False
            .MultiLine = True                 ' so ^/$ work for each line inside s
        End With
    
        On Error Resume Next
        ff = FreeFile
        Open fullPath For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, s
            ' If this "line" actually contains multiple physical lines, normalize
            s = Replace$(Replace$(s, vbCrLf, vbLf), vbCr, vbLf)
            If rx.Test(s) Then
                Set m = rx.Execute(s)(0)
                ResolveShortcutLink = Trim$(m.SubMatches(0))
                Close #ff
                Exit Function
            End If
        Loop
        Close #ff
        On Error GoTo 0
        Exit Function
    End If

    If EndsWithCI(fullPath, ".lnk") Then
        Dim sh As Object, lnk As Object
        Set sh = CreateObject("WScript.Shell")
        Set lnk = sh.CreateShortcut(fullPath)
        Dim tgt As String, args As String
        tgt = CStr(lnk.TargetPath)
        args = CStr(lnk.Arguments)
        ' Prefer explicit URL in Arguments; else TargetPath
        If InStr(1, args, "http", vbTextCompare) > 0 Or InStr(1, args, "file://", vbTextCompare) > 0 Then
            ResolveShortcutLink = Trim$(args)
        Else
            ResolveShortcutLink = Trim$(tgt)
        End If
        Exit Function
    End If

    ResolveShortcutLink = ""
    Exit Function

EH:
    On Error Resume Next
    ResolveShortcutLink = ""
End Function

' Add clickable hyperlinks to every non-empty Link cell in tbFiles.
Private Sub ApplyHyperlinksOnLinkColumn(ByVal loFiles As ListObject, ByVal colLink As Long)
    On Error GoTo EH
    If loFiles Is Nothing Then Exit Sub
    If loFiles.DataBodyRange Is Nothing Then Exit Sub
    If colLink <= 0 Then Exit Sub

    Dim rng As Range, cell As Range
    Set rng = loFiles.DataBodyRange.Columns(colLink)

    Application.DisplayAlerts = False
    For Each cell In rng.Cells
        If Len(CStr(cell.Value)) > 0 Then
            If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete
            cell.Worksheet.Hyperlinks.Add Anchor:=cell, Address:=CStr(cell.Value), TextToDisplay:=CStr(cell.Value)
        Else
            If cell.Hyperlinks.Count > 0 Then cell.Hyperlinks.Delete
        End If
    Next cell
    Application.DisplayAlerts = True
    Exit Sub

EH:
    Application.DisplayAlerts = True
End Sub

' Build JSON for the given workbook, skipping yellow rows,
' and emitting: Repository, Location, ObjectName, PrimaryName, ObjectType,
' Description, Keywords, Reference, DateDoc, Approved, Relevance, Link
Private Function BuildJSONFromWorkbookIgnoreYellow(ByVal wb As Excel.Workbook) As String
    Const SMALL_SAMPLE As Boolean = False   ' << set True to emit only first 10 non-yellow rows

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim loConfig As ListObject
    Dim found As Range
    Dim shortName As String

    Dim json As String, sep As String
    Dim r As ListRow
    Dim taken As Long

    ' 1) Find tbFiles in this workbook
    Set lo = Nothing
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects("tbFiles")
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        BuildJSONFromWorkbookIgnoreYellow = "[]"
        Exit Function
    End If

    ' 2) Find tbConfig and read Short Name (Repository)
    Set loConfig = Nothing
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set loConfig = ws.ListObjects("tbConfig")
        On Error GoTo 0
        If Not loConfig Is Nothing Then Exit For
    Next ws

    If loConfig Is Nothing Then
        shortName = "..."
    Else
        With loConfig.DataBodyRange
            Set found = .Columns(loConfig.ListColumns("Key").Index) _
                        .Find(What:="Short Name", LookIn:=xlValues, LookAt:=xlWhole)
        End With

        If found Is Nothing Then
            shortName = "..."
        Else
            shortName = CStr(found.Offset(0, _
                loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
            If Len(shortName) = 0 Then shortName = "..."
        End If
    End If

    ' 3) Build JSON rows
    json = "[": sep = "": taken = 0

    Dim hasRel As Boolean, hasObjName As Boolean, hasObjType As Boolean
    Dim hasDesc As Boolean, hasKw As Boolean, hasRef As Boolean
    Dim hasDateDoc As Boolean, hasApproved As Boolean, hasLink As Boolean
    Dim hasRelevance As Boolean

    hasRel = ColumnExists(lo, "RelativePath")
    hasObjName = ColumnExists(lo, "Object name")
    hasObjType = ColumnExists(lo, "Object Type")
    hasDesc = ColumnExists(lo, "Description")
    hasKw = ColumnExists(lo, "Keywords")
    hasRef = ColumnExists(lo, "Reference")
    hasDateDoc = ColumnExists(lo, "Date Doc")
    hasApproved = ColumnExists(lo, "Approved?")
    hasLink = ColumnExists(lo, "Link")
    hasRelevance = ColumnExists(lo, "Relevance")

    Dim relPath As String, location As String
    Dim objectName As String, primaryName As String
    Dim objType As String, objTypeOut As String
    Dim lastSlash As Long, dotPos As Long
    Dim baseNoExt As String
    Dim cellHash As Range
    Dim relvRaw As String, relvOut As String

    For Each r In lo.ListRows
        ' Skip yellow rows (missing in filesystem) using the "#" cell style
        Set cellHash = r.Range.Cells(1, lo.ListColumns("#").Index)
        If cellHash.Interior.Color = vbYellow Then
            ' skip
        Else
            ' -- Relative path and pull Location/ObjectName --
            If hasRel Then
                relPath = CStr(r.Range.Cells(1, lo.ListColumns("RelativePath").Index).Value)
            Else
                relPath = ""
            End If

            If hasObjName Then
                objectName = CStr(r.Range.Cells(1, lo.ListColumns("Object name").Index).Value)
            Else
                objectName = ""
            End If

            ' Location = parent of RelativePath (empty if top-level)
            If Len(relPath) = 0 Then
                location = ""
            Else
                lastSlash = InStrRev(relPath, "\")
                If lastSlash > 0 Then
                    location = Left$(relPath, lastSlash - 1)
                Else
                    location = ""
                End If
            End If

            ' ObjectType from table
            If hasObjType Then
                objType = CStr(r.Range.Cells(1, lo.ListColumns("Object Type").Index).Value)
            Else
                objType = ""
            End If

            ' PrimaryName & ObjectTypeOut
            Select Case LCase$(objType)
                Case "category", "folder", "subfolder"
                    primaryName = objectName
                    objTypeOut = objType
                Case Else
                    dotPos = InStrRev(objectName, ".")
                    If dotPos > 0 Then
                        baseNoExt = Left$(objectName, dotPos - 1)
                    Else
                        baseNoExt = objectName
                    End If

                    If dotPos > 0 Then
                        Dim extLower As String
                        extLower = LCase$(Mid$(objectName, dotPos + 1))
                        If extLower = "url" Or extLower = "lnk" Then
                            baseNoExt = StripTrailingParenGroup(baseNoExt)
                        End If
                    End If
                    primaryName = baseNoExt

                    If Len(objType) > 0 Then
                        objTypeOut = objType
                    Else
                        If InStrRev(objectName, ".") = 0 Then
                            objTypeOut = "File"   ' no extension ? File
                        Else
                            objTypeOut = ""
                        End If
                    End If
            End Select

            ' Relevance (Low/Normal/High or blank?Medium)
            If hasRelevance Then
                relvRaw = Trim$(CStr(r.Range.Cells(1, lo.ListColumns("Relevance").Index).Value))
            Else
                relvRaw = ""
            End If
            Select Case LCase$(relvRaw)
                Case "", "n", "normal": relvOut = "Normal"
                Case "l", "low":       relvOut = "Low"
                Case "h", "high":      relvOut = "High"
                Case Else:             relvOut = "Normal"
            End Select

            ' Emit JSON row
            json = json & sep & "{"
            json = json & """Repository"":""" & j(shortName) & ""","
            json = json & """Location"":""" & j(location) & ""","
            json = json & """ObjectName"":""" & j(objectName) & ""","
            json = json & """PrimaryName"":""" & j(primaryName) & ""","
            json = json & """ObjectType"":""" & j(objTypeOut) & ""","
            If hasDesc Then json = json & """Description"":""" & j(r.Range.Cells(1, lo.ListColumns("Description").Index).Value) & """," Else json = json & """Description"":"""","
            If hasKw Then json = json & """Keywords"":""" & j(r.Range.Cells(1, lo.ListColumns("Keywords").Index).Value) & """," Else json = json & """Keywords"":"""","
            If hasRef Then json = json & """Reference"":""" & j(r.Range.Cells(1, lo.ListColumns("Reference").Index).Value) & """," Else json = json & """Reference"":"""","
            If hasDateDoc Then json = json & """DateDoc"":""" & j(r.Range.Cells(1, lo.ListColumns("Date Doc").Index).Value) & """," Else json = json & """DateDoc"":"""","
            If hasApproved Then json = json & """Approved"":""" & j(r.Range.Cells(1, lo.ListColumns("Approved?").Index).Value) & """," Else json = json & """Approved"":"""","
            json = json & """Relevance"":""" & j(relvOut) & ""","    ' << NEW FIELD
            If hasLink Then json = json & """Link"":""" & j(r.Range.Cells(1, lo.ListColumns("Link").Index).Value) & """" Else json = json & """Link"":"""" "
            json = json & "}"

            sep = ","
            taken = taken + 1
            If SMALL_SAMPLE And taken >= 10 Then Exit For
        End If
    Next r

    json = json & "]"
    BuildJSONFromWorkbookIgnoreYellow = json
End Function


' If s ends with " (...)" (a single parenthesized group) remove that suffix.
' Example: "File3 (SharePoint Video)" -> "File3"
Private Function StripTrailingParenGroup(ByVal s As String) As String
    Dim pClose As Long, pOpen As Long
    s = Trim$(s)
    If Len(s) = 0 Then
        StripTrailingParenGroup = s
        Exit Function
    End If

    If Right$(s, 1) <> ")" Then
        StripTrailingParenGroup = s
        Exit Function
    End If

    pClose = Len(s)
    pOpen = InStrRev(s, "(")
    If pOpen > 0 And pOpen < pClose Then
        ' Make sure there is exactly one space before "(" or it starts the string
        Dim leftPart As String, midPart As String
        leftPart = Left$(s, pOpen - 1)
        midPart = Mid$(s, pOpen, pClose - pOpen + 1)
        ' Heuristic: only strip if it's at the end and looks like a simple group
        If pOpen > 1 And Mid$(s, pOpen - 1, 1) = " " Then
            StripTrailingParenGroup = RTrim$(Left$(s, pOpen - 2 + 1))  ' remove space before "(" too
        Else
            StripTrailingParenGroup = Left$(s, pOpen - 1)
        End If
    Else
        StripTrailingParenGroup = s
    End If
End Function


' JSON-escape helper
Private Function j(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then v = ""
    Dim s As String
    s = CStr(v)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    j = s
End Function


Public Sub GenerateFileIndex()
    Dim wbPrimary   As Excel.Workbook
    Dim wbExtra     As Excel.Workbook
    Dim wbTest      As Excel.Workbook

    Dim ws          As Worksheet
    Dim loConfig    As ListObject

    Dim templatePath As String, outputPath As String, repoName As String
    Dim html As String, json As String, jsonAll As String
    Dim stm As Object ' ADODB.Stream late-bound
    Dim found As Range

    Dim alsoPaths(1 To 4) As String
    Dim i As Long, p As String

    ' Use the ACTIVE workbook as the primary
    Set wbPrimary = ActiveWorkbook
    If wbPrimary Is Nothing Then
        MsgBox "No active workbook.", vbCritical
        Exit Sub
    End If

    ' 1) tbConfig in PRIMARY
    For Each ws In wbPrimary.Worksheets
        On Error Resume Next
        Set loConfig = ws.ListObjects("tbConfig")
        On Error GoTo 0
        If Not loConfig Is Nothing Then Exit For
    Next ws
    If loConfig Is Nothing Then
        MsgBox "Configuration table 'tbConfig' not found in primary workbook.", vbCritical: Exit Sub
    End If

    ' Repository name (for the HTML title, not the per-row Repository column)
    With loConfig.DataBodyRange
        Set found = .Columns(loConfig.ListColumns("Key").Index) _
            .Find(What:="Repository name", LookIn:=xlValues, LookAt:=xlWhole)
    End With
    If found Is Nothing Then
        repoName = "Repository"
    Else
        repoName = CStr(found.Offset(0, _
          loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)
        If Len(repoName) = 0 Then repoName = "Repository"
    End If

    ' Html Template
    With loConfig.DataBodyRange
        Set found = .Columns(loConfig.ListColumns("Key").Index) _
            .Find(What:="Html Template", LookIn:=xlValues, LookAt:=xlWhole)
    End With
    If found Is Nothing Then
        MsgBox "No 'Html Template' in tbConfig (primary).", vbCritical: Exit Sub
    End If
    templatePath = CStr(found.Offset(0, _
      loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)

    ' Html Index file
    With loConfig.DataBodyRange
        Set found = .Columns(loConfig.ListColumns("Key").Index) _
            .Find(What:="Html Index file", LookIn:=xlValues, LookAt:=xlWhole)
    End With
    If found Is Nothing Then
        MsgBox "No 'Html Index file' in tbConfig (primary).", vbCritical: Exit Sub
    End If
    outputPath = CStr(found.Offset(0, _
      loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)

    ' Also Read 1..4 (full paths, may be missing or blank)
    For i = 1 To 4
        With loConfig.DataBodyRange
            Set found = .Columns(loConfig.ListColumns("Key").Index) _
                .Find(What:="Also Read " & CStr(i), LookIn:=xlValues, LookAt:=xlWhole)
        End With
        If Not found Is Nothing Then
            alsoPaths(i) = Trim$(CStr(found.Offset(0, _
              loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value))
        Else
            alsoPaths(i) = ""
        End If
    Next i

    ' 2) Read HTML template from disk
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' text
        .Charset = "utf-8"
        .Open
        .LoadFromFile templatePath
        html = .ReadText(-1)
        .Close
    End With

    ' 3) Build JSON for PRIMARY workbook
    jsonAll = BuildJSONFromWorkbookIgnoreYellow(wbPrimary)

    ' 4) For each "Also Read N" path, try to open and append JSON
    Dim baseName As String
    For i = 1 To 4
        p = alsoPaths(i)
        If Len(p) > 0 Then
            baseName = Dir$(p)
            Set wbExtra = Nothing

            ' Try to open the workbook (ignore macros inside extra workbooks)
            On Error Resume Next
            Application.Workbooks.Open fileName:=p, ReadOnly:=True
            ' If there was a real failure, Err.Number will be non-zero
            If Err.Number <> 0 Then
                Debug.Print "Failed to open extra workbook: " & p & "  Err " & Err.Number & ": " & Err.Description
                MsgBox "Failed to open extra workbook: " & p & "  Err " & Err.Number & ": " & Err.Description, vbCritical
                Err.Clear
                On Error GoTo 0
            Else
                ' Find the workbook object by FullName
                On Error GoTo 0
                For Each wbTest In Application.Workbooks
                    If StrComp(wbTest.Name, baseName, vbTextCompare) = 0 Then
                        Set wbExtra = wbTest
                        Exit For
                    End If
                Next wbTest

                If wbExtra Is Nothing Then
                    Debug.Print "Opened extra workbook but could not resolve object for: " & p
                Else
                    Dim jsonExtra As String
                    jsonExtra = BuildJSONFromWorkbookIgnoreYellow(wbExtra)

                    ' Merge JSON arrays: remove trailing ']' from jsonAll and leading '[' from jsonExtra
                    If jsonAll = "[]" Then
                        jsonAll = jsonExtra
                    ElseIf jsonExtra <> "[]" Then
                        jsonAll = Left$(jsonAll, Len(jsonAll) - 1) & _
                                  "," & Mid$(jsonExtra, 2)
                    End If

                    ' Optionally close extra workbook afterwards (no save)
                    wbExtra.Close SaveChanges:=False
                End If
            End If
        End If
    Next i

    ' If for some reason jsonAll ended empty, normalize
    If Len(jsonAll) = 0 Then jsonAll = "[]"

    ' 5) Replace REPNAME and timestamp
    html = Replace(html, "REPNAME", repoName, , , vbTextCompare)
    Dim stamp As String
    stamp = Format(Now, "d mmm HH:nn")
    html = Replace(html, "1 Jan 23:45", stamp, , , vbTextCompare)

    ' 6) Inject const fileData = ...;
    Dim startPos As Long, openPos As Long, endPos As Long
    Dim before As String, after As String
    startPos = InStr(1, html, "const fileData =", vbTextCompare)
    If startPos > 0 Then
        openPos = InStr(startPos, html, "[")
        endPos = InStr(openPos, html, "];")
        If openPos > 0 And endPos > 0 Then
            before = Left$(html, startPos - 1)
            after = Mid$(html, endPos + 2)
            html = before & "const fileData = " & jsonAll & ";" & after
        End If
    End If

    ' 7) Write output utf-8
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' text
        .Charset = "utf-8"
        .Open
        .WriteText html
        .SaveToFile outputPath, 2 ' adSaveCreateOverWrite
        .Close
    End With

    MsgBox "HTML index written to:" & vbCrLf & outputPath, vbInformation
End Sub



Private Function GetWorkbookByFullName(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set GetWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function


'-----------------------------------------------
' Helper: Write a UTF-8 file via ADODB.Stream
Private Sub WriteUTF8File(ByVal filePath As String, ByVal content As String)
    Dim stm As New ADODB.Stream
    With stm
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText content
        .SaveToFile filePath, adSaveCreateOverWrite
        .Close
    End With
End Sub



''' Sorts a 0-based array of strings in place by descending length
Private Sub SortKeysByLengthDesc(ByRef keys() As String)
    ' simple bubble for small N (N = number of ruled paths)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If Len(keys(i)) < Len(keys(j)) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
End Sub

''' Returns True for Domain/Category/Folder/Subfolder, False otherwise
Private Function isFolder(objType As String) As Boolean
    Select Case LCase(objType)
        Case "domain", "category", "folder", "subfolder"
            isFolder = True
        Case Else
            isFolder = False
    End Select
End Function

''' Ensures non-blank kw begins AND ends with ".", leaves blank as is
Public Function CleanKeywords(ByVal kw As String) As String
    kw = Trim(kw)
    If kw = "" Then
        CleanKeywords = ""
    Else
        If Left$(kw, 1) <> "." Then kw = "." & kw
        If Right$(kw, 1) <> "." Then kw = kw & "."
        CleanKeywords = kw
    End If
End Function

''' Rebuilds tbKeywords from all cleaned keywords in tbFiles
Public Sub UpdateKeywordsTable()
    Dim ws          As Worksheet
    Dim loFiles     As ListObject
    Dim loKeywords  As ListObject
    Dim r           As ListRow
    Dim dict        As Object
    Dim rawKW       As String
    Dim innerList   As Variant
    Dim kw          As Variant
    Dim newRow      As ListRow
    Dim kwCount     As Long

    ' 1) Find tbFiles
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loFiles = ws.ListObjects("tbFiles")
        On Error GoTo 0
        If Not loFiles Is Nothing Then Exit For
    Next ws
    If loFiles Is Nothing Then
        MsgBox "Table tbFiles not found.", vbCritical: Exit Sub
    End If

    ' 2) Find tbKeywords
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loKeywords = ws.ListObjects("tbKeywords")
        On Error GoTo 0
        If Not loKeywords Is Nothing Then Exit For
    Next ws
    If loKeywords Is Nothing Then
        MsgBox "Table tbKeywords not found.", vbCritical: Exit Sub
    End If

    ' 3) Clear tbKeywords
    If Not loKeywords.DataBodyRange Is Nothing Then
        loKeywords.DataBodyRange.Delete
    End If

    ' 4) Collect unique keywords into dict
    Set dict = CreateObject("Scripting.Dictionary")
    For Each r In loFiles.ListRows
        rawKW = Trim$(r.Range.Cells(1, loFiles.ListColumns("Keywords").Index).Value)
        rawKW = CleanKeywords(rawKW) ' ensures leading/trailing dots if non-blank
        If Len(rawKW) >= 2 Then
            ' strip leading/trailing "." then split
            innerList = Split(Mid$(rawKW, 2, Len(rawKW) - 2), ".")
            For Each kw In innerList
                If Len(kw) > 0 Then
                    If Not dict.Exists(kw) Then dict.Add kw, kw
                End If
            Next kw
        End If
    Next r

    ' 5) Populate tbKeywords
    If dict.Count > 0 Then
        For Each kw In dict.keys
            Set newRow = loKeywords.ListRows.Add
            newRow.Range.Cells(1, loKeywords.ListColumns("Keyword").Index).Value = kw
        Next kw
    End If

    ' 6) Optional: sort tbKeywords by Keyword — ONLY if there are 2+ rows
    kwCount = loKeywords.ListRows.Count
    If kwCount > 1 Then
        With loKeywords.Sort
            .SortFields.Clear
            .SortFields.Add key:=loKeywords.ListColumns("Keyword").DataBodyRange, _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
End Sub


' Helper: flatten MyContainer into a Collection of FileDescriptor
Private Sub ScanFoldersCollectDescriptors(root As String, coll As Collection)
    Dim fMap As MyContainer, fp As Variant
    Dim cont As MyContainer, key As Variant

    Set fMap = New MyContainer
    ScanFoldersAndFiles root, fMap

    For Each fp In fMap.keys
        Set cont = fMap.Item(fp)
        If cont.Count > 0 Then
            For Each key In cont.keys
                coll.Add cont.Item(key)
            Next key
        End If
    Next fp
End Sub


Public Sub QuickSortArray(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot As Variant
    Dim tmpSwap As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long

    tmpLow = inLow
    tmpHi = inHi

    pivot = vArray((inLow + inHi) \ 2)

    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend

        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend

    If (inLow < tmpHi) Then QuickSortArray vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSortArray vArray, tmpLow, inHi
End Sub

' Initialize visited guard before recursion (CHANGED)
Public Sub ScanFoldersAndFiles(ByVal rootFolderPath As String, ByRef containerOut As MyContainer)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject

    If Not fso.FolderExists(rootFolderPath) Then
        Err.Raise vbObjectError + 100, , "Folder not found: " & rootFolderPath
    End If

    ' reset visited set for this run
    Set visited = CreateObject("Scripting.Dictionary")

    Dim rootFolder As Scripting.folder
    Set rootFolder = fso.GetFolder(rootFolderPath)

    ' Skip root if hidden/system
    If (rootFolder.Attributes And 6) <> 0 Then Exit Sub

    RecursiveFolderScan rootFolder, containerOut, rootFolderPath, 1
End Sub

' Add re-visit guard at the very start (CHANGED)
Private Sub RecursiveFolderScan( _
        ByVal folder As Scripting.folder, _
        ByRef containerOut As MyContainer, _
        ByVal theRoot As String, _
        ByVal theLevel As Integer)

    If visited Is Nothing Then Set visited = CreateObject("Scripting.Dictionary")
    If visited.Exists(folder.Path) Then Exit Sub
    visited.Add folder.Path, True

    Dim fileContainer   As MyContainer
    Dim subfolder       As Scripting.folder
    Dim file            As Scripting.file
    Dim aDescriptor     As FileDescriptor
    Dim relPath         As String
    Dim parts()         As String
    Dim depth           As Long
    Dim basePath        As String

    Set fileContainer = New MyContainer

    ' Skip hidden/system
    If (folder.Attributes And 6) <> 0 Then Exit Sub

    '--- Process subfolders ---
    For Each subfolder In folder.SubFolders
        If (subfolder.Attributes And 6) = 0 Then
            Set aDescriptor = New FileDescriptor

            relPath = Mid$(subfolder.Path, Len(theRoot) + 1)
            parts = Split(relPath, "\")
            depth = UBound(parts)

            If depth >= 0 Then aDescriptor.Domain = parts(0) Else aDescriptor.Domain = ""
            If depth >= 1 Then aDescriptor.Category = parts(1) Else aDescriptor.Category = ""
            If depth >= 2 Then aDescriptor.folder = parts(2) Else aDescriptor.folder = ""

            Select Case depth
              Case 0: aDescriptor.ObjectType = "Domain"
              Case 1: aDescriptor.ObjectType = "Category"
              Case 2: aDescriptor.ObjectType = "Folder"
              Case Else: aDescriptor.ObjectType = "Subfolder"
            End Select

            aDescriptor.fileName = subfolder.Name

            basePath = theRoot
            If aDescriptor.Domain <> "" Then basePath = basePath & "\" & aDescriptor.Domain
            If aDescriptor.Category <> "" Then basePath = basePath & "\" & aDescriptor.Category
            If aDescriptor.folder <> "" Then basePath = basePath & "\" & aDescriptor.folder

            aDescriptor.objectName = Replace(Mid$(subfolder.Path, Len(basePath) + 1), "\", "/")
            aDescriptor.RelativePath = relPath

            fileContainer.Add aDescriptor, subfolder.Name
            RecursiveFolderScan subfolder, containerOut, theRoot, theLevel + 1
        End If
        If (subfolder.Files.Count Mod 200) = 0 Then DoEvents
    Next

    '--- Process files ---
    For Each file In folder.Files
        If (file.Attributes And 6) = 0 Then
            Set aDescriptor = New FileDescriptor

            relPath = Mid$(file.Path, Len(theRoot) + 1)
            parts = Split(relPath, "\")
            depth = UBound(parts)

            If depth >= 1 Then aDescriptor.Domain = parts(0) Else aDescriptor.Domain = ""
            If depth >= 2 Then aDescriptor.Category = parts(1) Else aDescriptor.Category = ""
            If depth >= 3 Then aDescriptor.folder = parts(2) Else aDescriptor.folder = ""

            aDescriptor.ObjectType = GetFileFormat(file.Name)
            aDescriptor.fileName = parts(depth)

            basePath = theRoot
            If aDescriptor.Domain <> "" Then basePath = basePath & "\" & aDescriptor.Domain
            If aDescriptor.Category <> "" Then basePath = basePath & "\" & aDescriptor.Category
            If aDescriptor.folder <> "" Then basePath = basePath & "\" & aDescriptor.folder

            aDescriptor.objectName = Replace(Mid$(file.Path, Len(basePath) + 1), "\", "/")
            aDescriptor.RelativePath = relPath

            fileContainer.Add aDescriptor, file.Name
        End If
        If (folder.Files.Count Mod 400) = 0 Then DoEvents
    Next

    containerOut.Add fileContainer, folder.Path
End Sub

'=========================================================
' ScanFilesystemWithShow
'   Entry point from FolderScan
'   - rootPath: tbConfig[Local Root]
'   - domainName: tbConfig[Domain name]
'   - showSettings: Dictionary(RelativePath -> "all"/"nothing"/"1st level")
'   - allowed: Dictionary(RelativePath -> FileDescriptor)
'=========================================================
Private Sub ScanFilesystemWithShow( _
    ByVal rootPath As String, _
    ByVal domainName As String, _
    ByVal showSettings As Object, _
    ByVal allowed As Object)

    Dim fso As Scripting.FileSystemObject
    Dim rootFolder As Scripting.folder
    Dim t0 As Double

    scanSteps = 0        ' reset step counter each run
    Debug.Print String(60, "-")
    Debug.Print "ScanFS ENTER " & Now & " root=" & rootPath
    t0 = Timer

    Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists(rootPath) Then
        Err.Raise vbObjectError + 100, , "Folder not found: " & rootPath
    End If

    Set rootFolder = fso.GetFolder(rootPath)

    ' reset visited guard for this run
    Set visited = CreateObject("Scripting.Dictionary")

    ' implicit root rule = "1st level", start at depth 0
    RecursiveScanFolderWithShow rootFolder, "", "1st level", showSettings, domainName, allowed

    Debug.Print "ScanFS EXIT  items=" & allowed.Count & _
                "  secs=" & Format$(Timer - t0, "0.0")
End Sub

Function GetFileFormat(fileName As String) As String
    Dim ext As String
    Dim dotPos As Long

    ' Find last dot in the filename
    dotPos = InStrRev(fileName, ".")
    
    If dotPos = 0 Then
        GetFileFormat = "" ' No extension
        Exit Function
    End If

    ' Extract extension and convert to lowercase
    ext = LCase(Mid(fileName, dotPos + 1))

    Select Case ext
        Case "ppt", "pptx", "pptm"
            GetFileFormat = "pptx"
        Case "doc", "docx"
            GetFileFormat = "docx"
        Case "xls", "xlsx", "xlsm"
            GetFileFormat = "xlsx"
        Case Else
            ' Capitalize first letter, rest lower case
            GetFileFormat = UCase(Left(ext, 1)) & LCase(Mid(ext, 2))
            GetFileFormat = LCase(ext)
    End Select
End Function




' Recursive scan with Show? inheritance.
' parentRule is the effective Show? rule inherited from the parent folder
' ("" at the very top ? implicit "1st level").
Private Sub RecurseFolderWithShow_TRASHNOTUSED( _
        ByVal folder As Scripting.folder, _
        ByVal rootPath As String, _
        ByVal relPath As String, _
        ByVal parentRule As String, _
        ByVal domainName As String, _
        ByVal showSettings As Object, _
        ByRef allowed As Object)

    Dim myRule As String
    Dim explicitRule As String
    Dim key As String
    Dim file As Scripting.file
    Dim subfolder As Scripting.folder
    Dim childRel As String
    Dim desc As FileDescriptor
    Dim parts() As String
    Dim depth As Long

    If visited Is Nothing Then Set visited = CreateObject("Scripting.Dictionary")
    If visited.Exists(folder.Path) Then Exit Sub
    visited.Add folder.Path, True

    ' Skip hidden/system
    If (folder.Attributes And 6) <> 0 Then Exit Sub

    key = relPath

    ' Determine effective rule for THIS folder
    If LCase$(parentRule) = "nothing" Then
        ' Once a parent is Nothing, whole subtree is effectively Nothing
        myRule = "nothing"
    Else
        If showSettings Is Nothing Then
            explicitRule = ""
        ElseIf showSettings.Exists(key) Then
            explicitRule = LCase$(CStr(showSettings(key)))
        Else
            explicitRule = ""
        End If

        If Len(explicitRule) > 0 Then
            Select Case explicitRule
                Case "all", "nothing", "1st level"
                    myRule = explicitRule
                Case Else
                    myRule = "nothing"
            End Select
        Else
            ' inherit from parent, or implicit root rule
            If Len(parentRule) > 0 Then
                myRule = LCase$(parentRule)
            Else
                myRule = "1st level"
            End If
        End If
    End If

    ' Normalize for safety
    Select Case myRule
        Case "all", "nothing", "1st level"
        Case Else
            myRule = "nothing"
    End Select

    ' Create descriptor for this folder and add to allowed
    Set desc = New FileDescriptor
    desc.Domain = domainName
    desc.RelativePath = relPath
    desc.fileName = folder.Name
    desc.objectName = folder.Name

    If Len(relPath) > 0 Then
        parts = Split(relPath, "\")
        depth = UBound(parts)
    Else
        depth = -1
    End If

    If depth >= 0 Then desc.Category = parts(0) Else desc.Category = ""
    If depth >= 1 Then desc.folder = parts(1) Else desc.folder = ""

    Select Case depth
        Case 0: desc.ObjectType = "Category"
        Case 1: desc.ObjectType = "Folder"
        Case Else: desc.ObjectType = "Subfolder"
    End Select

    If allowed.Exists(relPath) Then allowed.Remove relPath
    allowed.Add relPath, desc

    ' Decide how deep to go based on myRule
    Select Case myRule
        Case "nothing"
            ' Folder row itself is visible, but no children
            Exit Sub

        Case "1st level"
            ' Folder row, plus immediate children only; no grandchildren

            ' Files directly inside this folder
            For Each file In folder.Files
                If (file.Attributes And 6) = 0 Then
                    Set desc = CreateFileDescriptorSimple(file, rootPath, domainName)
                    If Not allowed.Exists(desc.RelativePath) Then
                        allowed.Add desc.RelativePath, desc
                    End If
                End If
                If (folder.Files.Count Mod 400) = 0 Then DoEvents
            Next file

            ' Subfolders directly inside this folder – show them as rows,
            ' but do not recurse into their children.
            For Each subfolder In folder.SubFolders
                If (subfolder.Attributes And 6) = 0 Then
                    childRel = Mid$(subfolder.Path, Len(rootPath) + 1)
                    ' Build descriptor for the child folder
                    Set desc = New FileDescriptor
                    desc.Domain = domainName
                    desc.RelativePath = childRel
                    desc.fileName = subfolder.Name
                    desc.objectName = subfolder.Name

                    If Len(childRel) > 0 Then
                        parts = Split(childRel, "\")
                        depth = UBound(parts)
                    Else
                        depth = -1
                    End If

                    If depth >= 0 Then desc.Category = parts(0) Else desc.Category = ""
                    If depth >= 1 Then desc.folder = parts(1) Else desc.folder = ""

                    Select Case depth
                        Case 0: desc.ObjectType = "Category"
                        Case 1: desc.ObjectType = "Folder"
                        Case Else: desc.ObjectType = "Subfolder"
                    End Select

                    If allowed.Exists(childRel) Then allowed.Remove childRel
                    allowed.Add childRel, desc
                End If
                If (subfolder.Files.Count Mod 200) = 0 Then DoEvents
            Next subfolder

        Case Else  ' "all"
            ' Folder row + full recursion

            ' Files
            For Each file In folder.Files
                If (file.Attributes And 6) = 0 Then
                    Set desc = CreateFileDescriptorSimple(file, rootPath, domainName)
                    If Not allowed.Exists(desc.RelativePath) Then
                        allowed.Add desc.RelativePath, desc
                    End If
                End If
                If (folder.Files.Count Mod 400) = 0 Then DoEvents
            Next file

            ' Subfolders – recurse normally
            For Each subfolder In folder.SubFolders
                If (subfolder.Attributes And 6) = 0 Then
                    childRel = Mid$(subfolder.Path, Len(rootPath) + 1)
                    RecurseFolderWithShow_TRASHNOTUSED subfolder, rootPath, childRel, myRule, domainName, showSettings, allowed
                End If
                If (subfolder.Files.Count Mod 200) = 0 Then DoEvents
            Next subfolder
    End Select
End Sub

' Build a FileDescriptor for a file, without resolving shortcuts.
' Shortcut Target/Link is resolved later only for new/changed rows.
Private Function CreateFileDescriptorSimple( _
        ByVal fileObj As Scripting.file, _
        ByVal rootPath As String, _
        ByVal domainName As String) As FileDescriptor

    Dim d As FileDescriptor
    Dim relPath As String
    Dim parts() As String
    Dim depth As Long
    Dim extLower As String
    Dim t As String
    Dim dotPos As Long

    Set d = New FileDescriptor

    relPath = Mid$(fileObj.Path, Len(rootPath) + 1)
    d.RelativePath = relPath
    d.Domain = domainName

    If Len(relPath) > 0 Then
        parts = Split(relPath, "\")
        depth = UBound(parts)
    Else
        depth = -1
    End If

    ' For files: first segment is Category, second is Folder
    If depth >= 1 Then d.Category = parts(0) Else d.Category = ""
    If depth >= 2 Then d.folder = parts(1) Else d.folder = ""

    d.fileName = fileObj.Name
    d.objectName = fileObj.Name

    ' Default ObjectType from extension
    d.ObjectType = GetFileFormat(fileObj.Name)

    ' Special handling for shortcuts: type from "(Hint)" in name, but NO link resolution here
    dotPos = InStrRev(fileObj.Name, ".")
    If dotPos > 0 Then
        extLower = LCase$(Mid$(fileObj.Name, dotPos + 1))
        If extLower = "url" Or extLower = "lnk" Then
            t = DeriveShortcutTypeFromName(fileObj.Name)
            If Len(t) > 0 Then
                d.ObjectType = t           ' preserve hint case, e.g. "SharePoint Video"
            Else
                d.ObjectType = "shortcut"
            End If
            ' d.Link stays blank; will be filled only for new/blank rows in FolderScan
        End If
    End If

    Set CreateFileDescriptorSimple = d
End Function




'---------------------------------------------------------
' Helper: add or replace a FileDescriptor in `allowed`
'---------------------------------------------------------
Private Sub AddOrReplaceAllowed(ByVal dict As Object, _
                                ByVal key As String, _
                                ByVal fd As FileDescriptor)
    If dict.Exists(key) Then
        Set dict(key) = fd
    Else
        dict.Add key, fd
    End If
End Sub

' This is weird, but I had folders containing character 8203 in its name (no idea how I got this)
Private Function NormalizePath(ByVal s As String) As String
    If Len(s) = 0 Then
        NormalizePath = ""
        Exit Function
    End If

    ' ZERO WIDTH SPACE U+200B
    s = Replace$(s, ChrW(8203), "")

    ' Optional extras, if you ever hit them:
    ' BOM U+FEFF
    s = Replace$(s, ChrW(65279), "")
    ' Non-breaking space -> normal space
    s = Replace$(s, ChrW(160), " ")

    NormalizePath = s
End Function

'---------------------------------------------------------
' Recursive folder scan with Show? semantics
'   - folder: current FSO folder
'   - relPath: RelativePath from LocalRoot ("" for root)
'   - inheritedRule: "all" / "1st level" / "nothing"
'   - showSettings: Dict(RelPath -> explicit rule)
'   - allowed: Dict(RelPath -> FileDescriptor)
'
' Rules:
'   * "all"       ? folder + ALL descendants
'   * "nothing"   ? folder row only if explicit, no descendants
'   * "1st level" ? folder + immediate children (files+folders);
'                    recurse into a child folder ONLY if it has
'                    its own explicit Show? rule
'
' Visited guard:
'   Avoid infinite loops through junctions / reparse points by
'   remembering each physical Folder.Path once.
'---------------------------------------------------------
Private Sub RecursiveScanFolderWithShow( _
    ByVal folder As Scripting.folder, _
    ByVal relPath As String, _
    ByVal inheritedRule As String, _
    ByVal showSettings As Object, _
    ByVal domainName As String, _
    ByVal allowed As Object)

    'Const SCAN_MAX_STEPS As Long = 500000
    'Const MAX_SCAN_DEPTH As Long = 40

    'Static scanSteps As Long

    Dim effRule As String
    Dim hasExplicit As Boolean
    Dim fd As FileDescriptor
    Dim subfolder As Scripting.folder
    Dim file As Scripting.file
    Dim childRel As String
    Dim iCount As Long
    Dim childRule As String
    Dim depth As Long

    '---------- visited guard (physical path) ----------
    If visited Is Nothing Then Set visited = CreateObject("Scripting.Dictionary")
    If visited.Exists(folder.Path) Then Exit Sub
    visited.Add folder.Path, True

    ' Normalize relPath so dictionary keys are clean and match Excel
    relPath = NormalizePath(relPath)

    ' Root call: reset step counter for each new full scan
    If Len(relPath) = 0 Then
        scanSteps = 0
    End If

    ' Compute depth from relPath (number of segments)
    If Len(relPath) = 0 Then
        depth = 0
    Else
        depth = UBound(Split(relPath, "\")) + 1
    End If

    '----- step guard / debug -----
    scanSteps = scanSteps + 1
    If (scanSteps Mod 1000) = 0 Then
        Debug.Print "RS step"; scanSteps; " rel='"; relPath; "' rule='"; inheritedRule; "'"
    End If
    If scanSteps > SCAN_MAX_STEPS Then
        Err.Raise vbObjectError + 999, , _
            "RecursiveScanFolderWithShow aborted after " & scanSteps & _
            " steps at folder: " & folder.Path & " relPath=" & relPath
    End If

    '----- hard depth guard -----
    If depth > MAX_SCAN_DEPTH Then
        Debug.Print "Guard: MAX_SCAN_DEPTH reached at relPath=" & relPath & _
                    " (depth=" & depth & ")"
        Exit Sub
    End If

    ' Optional: manual breakpoint for specific paths
    'If relPath = "03 Sourcebook" Or relPath = "03 Sourcebook\002 Adobe XMP Specification" Then
    '    Debug.Print ">>> BREAK in RecursiveScanFolderWithShow, relPath=" & relPath & _
    '                "  inheritedRule=" & inheritedRule
    '    Stop     ' VB6-style breakpoint
    '    Debug.Print "force stop"
    'End If

    '----- Determine effective rule for THIS folder -----
    effRule = inheritedRule
    hasExplicit = False

    ' Show? only lives on folder rows, keyed by RelativePath
    If Len(relPath) > 0 Then
        If Not showSettings Is Nothing Then
            If showSettings.Exists(relPath) Then
                effRule = CStr(showSettings(relPath))      ' "all" / "nothing" / "1st level"
                hasExplicit = True
            End If
        End If
    End If

    ' Normalize effRule defensively
    Select Case LCase$(effRule)
        Case "all", "nothing", "1st level"
            effRule = LCase$(effRule)
        Case Else
            ' If something weird arrives, treat as "nothing"
            effRule = "nothing"
    End Select

    '----- Add folder descriptor (except for root relPath="") -----
    If Len(relPath) > 0 Then
        ' For Show? = Nothing:
        '   - If rule is explicit here, include this folder row;
        '   - If it is inherited, keep this branch invisible
        '     (existing Excel rows will later go yellow).
        If effRule = "nothing" And Not hasExplicit Then
            ' inherited Nothing ? do NOT add this folder row
        Else
            Set fd = BuildFolderDescriptor(folder, relPath, domainName)
            AddOrReplaceAllowed allowed, relPath, fd
        End If
    End If

    ' If effective rule is "nothing": do not go deeper at all
    If effRule = "nothing" Then Exit Sub

    '=====================================================
    ' Rule "all" ? folder + ALL descendants
    '=====================================================
    If effRule = "all" Then
        '--- files in this folder ---
        iCount = 0
        For Each file In folder.Files
            If (file.Attributes And 6) = 0 Then   ' skip hidden/system
                If Len(relPath) = 0 Then
                    childRel = file.Name
                Else
                    childRel = relPath & "\" & file.Name
                End If

                childRel = NormalizePath(childRel)

                Set fd = BuildFileDescriptor(file, childRel, domainName)
                AddOrReplaceAllowed allowed, childRel, fd
            End If

            iCount = iCount + 1
            If (iCount Mod 200) = 0 Then DoEvents
        Next file

        '--- subfolders (recurse) ---
        iCount = 0
        For Each subfolder In folder.SubFolders
            If (subfolder.Attributes And 6) = 0 Then  ' skip hidden/system
                If Len(relPath) = 0 Then
                    childRel = subfolder.Name
                Else
                    childRel = relPath & "\" & subfolder.Name
                End If

                childRel = NormalizePath(childRel)

                ' inheritedRule "all" – explicit rules (if any) will override inside
                RecursiveScanFolderWithShow subfolder, childRel, "all", showSettings, domainName, allowed
            End If

            iCount = iCount + 1
            If (iCount Mod 100) = 0 Then DoEvents
        Next subfolder

        Exit Sub
    End If

    '=====================================================
    ' Rule "1st level"
    '   - Show this folder and its immediate children
    '   - Recurse ONLY into child folders with explicit Show? rule
    '=====================================================
    If effRule = "1st level" Then
        '--- files directly inside this folder ---
        iCount = 0
        For Each file In folder.Files
            If (file.Attributes And 6) = 0 Then
                If Len(relPath) = 0 Then
                    childRel = file.Name
                Else
                    childRel = relPath & "\" & file.Name
                End If

                childRel = NormalizePath(childRel)

                Set fd = BuildFileDescriptor(file, childRel, domainName)
                AddOrReplaceAllowed allowed, childRel, fd
            End If

            iCount = iCount + 1
            If (iCount Mod 200) = 0 Then DoEvents
        Next file

        '--- subfolders directly inside this folder ---
        iCount = 0
        For Each subfolder In folder.SubFolders
            If (subfolder.Attributes And 6) = 0 Then
                If Len(relPath) = 0 Then
                    childRel = subfolder.Name
                Else
                    childRel = relPath & "\" & subfolder.Name
                End If

                childRel = NormalizePath(childRel)

                ' Always show the immediate child folder row
                Set fd = BuildFolderDescriptor(subfolder, childRel, domainName)
                AddOrReplaceAllowed allowed, childRel, fd

                ' If this child has its OWN explicit Show? rule,
                ' recurse into it so that that rule (all / 1st level / nothing)
                ' is applied independently of the parent 1st level.
                If Not showSettings Is Nothing Then
                    If showSettings.Exists(childRel) Then
                        childRule = CStr(showSettings(childRel))
                        RecursiveScanFolderWithShow subfolder, childRel, childRule, showSettings, domainName, allowed
                    End If
                End If
            End If

            iCount = iCount + 1
            If (iCount Mod 100) = 0 Then DoEvents
        Next subfolder

        Exit Sub
    End If

    '=====================================================
    ' Any other rule (should not happen – already normalized)
    ' Treat as "nothing" from here down.
    '=====================================================
    ' Do not descend.
End Sub




'---------------------------------------------------------
' Build FileDescriptor for a FOLDER
'   relPath: "Cat", "Cat\Folder", "Cat\Folder\Sub1", ...
'---------------------------------------------------------
Private Function BuildFolderDescriptor( _
    ByVal folder As Scripting.folder, _
    ByVal relPath As String, _
    ByVal domainName As String) As FileDescriptor

    Dim d As FileDescriptor
    Dim parts() As String
    Dim n As Long

    Set d = New FileDescriptor

    d.Domain = domainName
    d.RelativePath = relPath
    d.fileName = folder.Name
    d.objectName = folder.Name   ' not used in Excel, but OK

    If Len(relPath) > 0 Then
        parts = Split(relPath, "\")
        n = UBound(parts)

        If n >= 0 Then d.Category = parts(0) Else d.Category = ""
        If n >= 1 Then d.folder = parts(1) Else d.folder = ""

        Select Case n
            Case 0: d.ObjectType = "Category"
            Case 1: d.ObjectType = "Folder"
            Case Else: d.ObjectType = "Subfolder"
        End Select
    Else
        ' Root should never be stored, but set something safe
        d.Category = ""
        d.folder = ""
        d.ObjectType = "Folder"
    End If

    Set BuildFolderDescriptor = d
End Function

'---------------------------------------------------------
' Build FileDescriptor for a FILE
'   relPath: "Cat\File.ext", "Cat\Folder\File.ext", ...
'---------------------------------------------------------
Private Function BuildFileDescriptor( _
    ByVal file As Scripting.file, _
    ByVal relPath As String, _
    ByVal domainName As String) As FileDescriptor

    Dim d As FileDescriptor
    Dim parts() As String
    Dim n As Long

    Set d = New FileDescriptor

    d.Domain = domainName
    d.RelativePath = relPath
    d.fileName = file.Name
    d.objectName = file.Name   ' We only persist fileName into tbFiles

    If Len(relPath) > 0 Then
        parts = Split(relPath, "\")
        n = UBound(parts)

        ' For files: Category = first segment, Folder = second
        If n >= 1 Then d.Category = parts(0) Else d.Category = ""
        If n >= 2 Then d.folder = parts(1) Else d.folder = ""
    Else
        d.Category = ""
        d.folder = ""
    End If

    d.ObjectType = GetFileFormat(file.Name)

    Set BuildFileDescriptor = d
End Function


